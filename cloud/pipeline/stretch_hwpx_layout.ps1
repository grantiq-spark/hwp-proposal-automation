param(
    [string]$InputHwpx = "",
    [string]$OutputHwpx = "",
    [double]$Factor = 1.5
)

$ErrorActionPreference = "Stop"

function Write-Step {
    param([string]$Message)
    Write-Host "[KBio StretchHWPX] $Message"
}

function Resolve-InputPath {
    param([string]$Path, [string]$DefaultName)
    if ([string]::IsNullOrWhiteSpace($Path)) {
        $Path = Join-Path $PSScriptRoot $DefaultName
    }
    elseif (![System.IO.Path]::IsPathRooted($Path)) {
        $Path = Join-Path (Get-Location).Path $Path
    }
    if (!(Test-Path -LiteralPath $Path)) {
        throw "File not found: $Path"
    }
    return (Resolve-Path -LiteralPath $Path).Path
}

function Resolve-OutputPath {
    param([string]$Path, [string]$DefaultName)
    if ([string]::IsNullOrWhiteSpace($Path)) {
        $Path = Join-Path $PSScriptRoot $DefaultName
    }
    elseif (![System.IO.Path]::IsPathRooted($Path)) {
        $Path = Join-Path (Get-Location).Path $Path
    }
    return $Path
}

function Scale-IntAttribute {
    param(
        [System.Xml.XmlElement]$Node,
        [string]$Name,
        [double]$Scale,
        [int]$Min = 1
    )
    if ($Node -eq $null -or !$Node.HasAttribute($Name)) { return }
    $value = 0
    if ([int]::TryParse($Node.GetAttribute($Name), [ref]$value)) {
        $Node.SetAttribute($Name, [string]([Math]::Max([int][Math]::Round($value * $Scale), $Min)))
    }
}

$InputHwpx = Resolve-InputPath -Path $InputHwpx -DefaultName "KBio_ODA_FS_word_rich_paste.hwpx"
$OutputHwpx = Resolve-OutputPath -Path $OutputHwpx -DefaultName "KBio_ODA_FS_stretched.hwpx"

$workDir = Join-Path $PSScriptRoot "stretch_hwpx_work"
if (Test-Path -LiteralPath $workDir) {
    Remove-Item -LiteralPath $workDir -Recurse -Force
}
New-Item -ItemType Directory -Force -Path $workDir | Out-Null
tar -xf $InputHwpx -C $workDir

$headerPath = Join-Path $workDir "Contents\header.xml"
$sectionPath = Join-Path $workDir "Contents\section0.xml"

[xml]$header = Get-Content -LiteralPath $headerPath -Raw -Encoding UTF8
[xml]$section = Get-Content -LiteralPath $sectionPath -Raw -Encoding UTF8

$headerNs = New-Object System.Xml.XmlNamespaceManager($header.NameTable)
$headerNs.AddNamespace("hh", "http://www.hancom.co.kr/hwpml/2011/head")
$headerNs.AddNamespace("hc", "http://www.hancom.co.kr/hwpml/2011/core")

Write-Step "Scaling paragraph line spacing by $Factor..."
foreach ($ls in $header.SelectNodes("//*[local-name()='lineSpacing']")) {
    Scale-IntAttribute -Node $ls -Name "value" -Scale $Factor -Min 100
}

Write-Step "Scaling table and cell heights by $Factor..."
foreach ($node in $section.SelectNodes("//*[local-name()='tbl']/*[local-name()='sz']")) {
    Scale-IntAttribute -Node $node -Name "height" -Scale $Factor -Min 1000
}
foreach ($node in $section.SelectNodes("//*[local-name()='cellSz']")) {
    Scale-IntAttribute -Node $node -Name "height" -Scale $Factor -Min 600
}

Write-Step "Removing stale line layout caches..."
$lineSegArrays = @($section.SelectNodes("//*[local-name()='linesegarray']"))
foreach ($node in $lineSegArrays) {
    $node.ParentNode.RemoveChild($node) | Out-Null
}

$settings = New-Object System.Xml.XmlWriterSettings
$settings.Encoding = New-Object System.Text.UTF8Encoding($false)
$settings.Indent = $false

$writer = [System.Xml.XmlWriter]::Create($headerPath, $settings)
try { $header.Save($writer) } finally { $writer.Close() }

$writer = [System.Xml.XmlWriter]::Create($sectionPath, $settings)
try { $section.Save($writer) } finally { $writer.Close() }

if (Test-Path -LiteralPath $OutputHwpx) {
    Remove-Item -LiteralPath $OutputHwpx -Force
}
$zipTemp = [System.IO.Path]::ChangeExtension($OutputHwpx, ".zip")
if (Test-Path -LiteralPath $zipTemp) {
    Remove-Item -LiteralPath $zipTemp -Force
}
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory($workDir, $zipTemp, [System.IO.Compression.CompressionLevel]::Optimal, $false)
Move-Item -LiteralPath $zipTemp -Destination $OutputHwpx -Force

Write-Step "Created: $OutputHwpx"
