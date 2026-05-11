param(
    [string]$SourceDocx = "C:\Users\UMTR\Downloads\KBio_ODA_FS_v2_0.docx",
    [string]$OutputHwpx = ""
)

$ErrorActionPreference = "Stop"

function Write-Step {
    param([string]$Message)
    Write-Host "[KBio RichImport] $Message"
}

if (![System.IO.Path]::IsPathRooted($SourceDocx)) {
    $SourceDocx = Join-Path (Get-Location).Path $SourceDocx
}
if (!(Test-Path -LiteralPath $SourceDocx)) {
    throw "Source DOCX not found: $SourceDocx"
}

if ([string]::IsNullOrWhiteSpace($OutputHwpx)) {
    $OutputHwpx = Join-Path $PSScriptRoot "KBio_ODA_FS_word_rich_paste.hwpx"
}
elseif (![System.IO.Path]::IsPathRooted($OutputHwpx)) {
    $OutputHwpx = Join-Path (Get-Location).Path $OutputHwpx
}

if (Test-Path -LiteralPath $OutputHwpx) {
    Remove-Item -LiteralPath $OutputHwpx -Force
}

$word = $null
$doc = $null
$hwp = $null
try {
    Write-Step "Copying DOCX content from Word..."
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Open($SourceDocx, $false, $true)
    $doc.Content.Copy()
    Start-Sleep -Seconds 2

    Write-Step "Creating a new Hangul document..."
    $hwp = New-Object -ComObject "HWPFrame.HwpObject"
    $hwp.XHwpWindows.Item(0).Visible = $true
    try { $hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule") | Out-Null } catch {}
    $hwp.HAction.Run("FileNew") | Out-Null
    Start-Sleep -Seconds 1

    Write-Step "Pasting rich content into Hangul..."
    try {
        $hwp.HAction.Run("Paste") | Out-Null
    }
    catch {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.SendKeys]::SendWait("^v")
    }

    Write-Step "Waiting for Hangul import/layout..."
    Start-Sleep -Seconds 20
    try { $hwp.HAction.Run("MoveDocEnd") | Out-Null } catch {}
    Start-Sleep -Seconds 5

    Write-Step "Saving as HWPX: $OutputHwpx"
    $hwp.SaveAs($OutputHwpx, "HWPX", "") | Out-Null
    Write-Step "Done."
}
finally {
    if ($doc -ne $null) {
        try { $doc.Close($false) | Out-Null } catch {}
        [Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
    }
    if ($word -ne $null) {
        try { $word.Quit() | Out-Null } catch {}
        [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    if ($hwp -ne $null) {
        try { $hwp.Quit() | Out-Null } catch {}
        [Runtime.InteropServices.Marshal]::ReleaseComObject($hwp) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
