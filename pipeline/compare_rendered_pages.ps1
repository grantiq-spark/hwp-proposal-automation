param(
    [string]$SourceDocx = "C:\Users\UMTR\Downloads\KBio_ODA_FS_v2_0.docx",
    [string]$CandidateHwpx = "",
    [string]$WorkDir = "",
    [double]$Threshold = 0.90,
    [int]$Density = 144
)

$ErrorActionPreference = "Stop"

function Write-Step {
    param([string]$Message)
    Write-Host "[KBio VisualCompare] $Message"
}

function Resolve-RequiredPath {
    param([string]$Path, [string]$Name)
    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw "$Name path is required."
    }
    if (![System.IO.Path]::IsPathRooted($Path)) {
        $Path = Join-Path (Get-Location).Path $Path
    }
    if (!(Test-Path -LiteralPath $Path)) {
        throw "$Name not found: $Path"
    }
    return (Resolve-Path -LiteralPath $Path).Path
}

function Get-ToolPath {
    param([string]$Name)
    $cmd = Get-Command $Name -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    if ($Name -eq "magick") {
        $candidate = Get-ChildItem -LiteralPath "C:\Program Files" -Directory -Filter "ImageMagick*" -ErrorAction SilentlyContinue |
            ForEach-Object { Join-Path $_.FullName "magick.exe" } |
            Where-Object { Test-Path -LiteralPath $_ } |
            Select-Object -First 1
        if ($candidate) { return $candidate }
    }
    return $null
}

function Export-DocxToPdf {
    param([string]$Docx, [string]$Pdf)

    $word = $null
    $doc = $null
    try {
        Write-Step "Exporting DOCX to PDF..."
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $doc = $word.Documents.Open($Docx, $false, $true)
        $doc.ExportAsFixedFormat($Pdf, 17) | Out-Null
    }
    finally {
        if ($doc -ne $null) {
            $doc.Close($false) | Out-Null
            [Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
        }
        if ($word -ne $null) {
            $word.Quit() | Out-Null
            [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Export-HwpxToPdf {
    param([string]$Hwpx, [string]$Pdf)

    $hwp = $null
    try {
        Write-Step "Exporting HWPX to PDF with Hangul..."
        $hwp = New-Object -ComObject "HWPFrame.HwpObject"
        $hwp.XHwpWindows.Item(0).Visible = $true
        try { $hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule") | Out-Null } catch {}
        $hwp.Open($Hwpx, "HWPX", "") | Out-Null
        try { $hwp.HAction.Run("MoveDocEnd") | Out-Null } catch {}
        Start-Sleep -Seconds 2
        $hwp.SaveAs($Pdf, "PDF", "") | Out-Null
    }
    finally {
        if ($hwp -ne $null) {
            try { $hwp.Quit() | Out-Null } catch {}
            [Runtime.InteropServices.Marshal]::ReleaseComObject($hwp) | Out-Null
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Render-PdfWithMagick {
    param([string]$Pdf, [string]$OutDir, [string]$Prefix, [int]$Dpi)

    $magick = Get-ToolPath -Name "magick"
    if ($null -eq $magick) {
        throw "ImageMagick 'magick' was not found. Install ImageMagick with PDF support or Ghostscript, then rerun this script."
    }

    $gs = Get-ToolPath -Name "gswin64c"
    if ($null -eq $gs) {
        $gs = Get-ChildItem -LiteralPath "C:\Program Files\gs" -Recurse -Filter "gswin64c.exe" -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty FullName -First 1
    }
    if ($gs) {
        $gsDir = Split-Path -Parent $gs
        if ($env:PATH -notlike "*$gsDir*") {
            $env:PATH = "$gsDir;$env:PATH"
        }
    }

    New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
    $outPattern = Join-Path $OutDir "$Prefix-%03d.png"
    Write-Step "Rendering PDF pages: $Prefix"
    & $magick -density $Dpi $Pdf -background white -alpha remove -alpha off $outPattern
    if ($LASTEXITCODE -ne 0) {
        throw "ImageMagick failed to render $Pdf"
    }
}

function Get-PageImages {
    param([string]$Dir, [string]$Prefix)
    @(Get-ChildItem -LiteralPath $Dir -Filter "$Prefix-*.png" | Sort-Object Name)
}

function Get-BitmapSimilarity {
    param([string]$APath, [string]$BPath)

    Add-Type -AssemblyName System.Drawing
    $a = [System.Drawing.Bitmap]::FromFile($APath)
    $b = [System.Drawing.Bitmap]::FromFile($BPath)
    try {
        $w = 220
        $h = 311
        $ra = New-Object System.Drawing.Bitmap($w, $h)
        $rb = New-Object System.Drawing.Bitmap($w, $h)
        $ga = [System.Drawing.Graphics]::FromImage($ra)
        $gb = [System.Drawing.Graphics]::FromImage($rb)
        try {
            $ga.DrawImage($a, 0, 0, $w, $h)
            $gb.DrawImage($b, 0, 0, $w, $h)
        }
        finally {
            $ga.Dispose()
            $gb.Dispose()
        }

        $sumSq = 0.0
        $count = $w * $h
        for ($y = 0; $y -lt $h; $y++) {
            for ($x = 0; $x -lt $w; $x++) {
                $ca = $ra.GetPixel($x, $y)
                $cb = $rb.GetPixel($x, $y)
                $la = (0.299 * $ca.R) + (0.587 * $ca.G) + (0.114 * $ca.B)
                $lb = (0.299 * $cb.R) + (0.587 * $cb.G) + (0.114 * $cb.B)
                $d = ($la - $lb) / 255.0
                $sumSq += ($d * $d)
            }
        }

        $rmse = [Math]::Sqrt($sumSq / $count)
        return [Math]::Max(0.0, 1.0 - $rmse)
    }
    finally {
        if ($ra) { $ra.Dispose() }
        if ($rb) { $rb.Dispose() }
        $a.Dispose()
        $b.Dispose()
    }
}

if ([string]::IsNullOrWhiteSpace($CandidateHwpx)) {
    $CandidateHwpx = Join-Path $PSScriptRoot "KBio_ODA_FS_table_header_shaded.hwpx"
}

$SourceDocx = Resolve-RequiredPath -Path $SourceDocx -Name "Source DOCX"
$CandidateHwpx = Resolve-RequiredPath -Path $CandidateHwpx -Name "Candidate HWPX"

if ([string]::IsNullOrWhiteSpace($WorkDir)) {
    $WorkDir = Join-Path $PSScriptRoot "visual_compare_work"
}
elseif (![System.IO.Path]::IsPathRooted($WorkDir)) {
    $WorkDir = Join-Path (Get-Location).Path $WorkDir
}

if (Test-Path -LiteralPath $WorkDir) {
    Remove-Item -LiteralPath $WorkDir -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $WorkDir | Out-Null

$sourcePdf = Join-Path $WorkDir "source.docx.pdf"
$candidatePdf = Join-Path $WorkDir "candidate.hwpx.pdf"
$sourceImages = Join-Path $WorkDir "source_pages"
$candidateImages = Join-Path $WorkDir "candidate_pages"

Export-DocxToPdf -Docx $SourceDocx -Pdf $sourcePdf
Export-HwpxToPdf -Hwpx $CandidateHwpx -Pdf $candidatePdf
Render-PdfWithMagick -Pdf $sourcePdf -OutDir $sourceImages -Prefix "source" -Dpi $Density
Render-PdfWithMagick -Pdf $candidatePdf -OutDir $candidateImages -Prefix "candidate" -Dpi $Density

$sourcePages = Get-PageImages -Dir $sourceImages -Prefix "source"
$candidatePages = Get-PageImages -Dir $candidateImages -Prefix "candidate"
$maxPages = [Math]::Max($sourcePages.Count, $candidatePages.Count)
$results = New-Object System.Collections.Generic.List[object]

for ($i = 0; $i -lt $maxPages; $i++) {
    $pageNo = $i + 1
    if ($i -ge $sourcePages.Count -or $i -ge $candidatePages.Count) {
        $results.Add([pscustomobject]@{
            Page = $pageNo
            Similarity = 0
            Pass = $false
            Source = $(if ($i -lt $sourcePages.Count) { $sourcePages[$i].FullName } else { "" })
            Candidate = $(if ($i -lt $candidatePages.Count) { $candidatePages[$i].FullName } else { "" })
        })
        continue
    }

    $sim = Get-BitmapSimilarity -APath $sourcePages[$i].FullName -BPath $candidatePages[$i].FullName
    $results.Add([pscustomobject]@{
        Page = $pageNo
        Similarity = [Math]::Round($sim, 4)
        Pass = ($sim -ge $Threshold)
        Source = $sourcePages[$i].FullName
        Candidate = $candidatePages[$i].FullName
    })
}

$csv = Join-Path $WorkDir "page_similarity.csv"
$json = Join-Path $WorkDir "page_similarity.json"
$results | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8
$results | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $json -Encoding UTF8

$passed = @($results | Where-Object { $_.Pass }).Count
$avg = if ($results.Count -gt 0) { [Math]::Round((($results | Measure-Object -Property Similarity -Average).Average), 4) } else { 0 }
Write-Step "Pages passing threshold: $passed/$($results.Count)"
Write-Step "Average similarity: $avg"
Write-Step "CSV: $csv"
Write-Step "JSON: $json"

if ($passed -lt $results.Count) {
    Write-Step "Pages below threshold:"
    $results | Where-Object { -not $_.Pass } | Select-Object Page,Similarity | Format-Table -AutoSize
    exit 2
}
