param(
  [Parameter(Mandatory = $true)]
  [string]$InputHwpx,

  [Parameter(Mandatory = $true)]
  [string]$OutputText
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Web

$inputPath = (Resolve-Path -LiteralPath $InputHwpx).Path
$outPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputText)
$work = Join-Path ([System.IO.Path]::GetTempPath()) ("hwpx_text_" + [guid]::NewGuid().ToString("N"))

try {
  New-Item -ItemType Directory -Path $work -Force | Out-Null
  [System.IO.Compression.ZipFile]::ExtractToDirectory($inputPath, $work)

  $xmlFiles = Get-ChildItem -LiteralPath $work -Recurse -Filter *.xml |
    Where-Object { $_.FullName -match "\\Contents\\|section|content" } |
    Sort-Object FullName

  $lines = New-Object System.Collections.Generic.List[string]
  foreach ($file in $xmlFiles) {
    $raw = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8
    $matches = [regex]::Matches($raw, "<hp:t[^>]*>(.*?)</hp:t>", "Singleline")
    foreach ($match in $matches) {
      $text = [System.Web.HttpUtility]::HtmlDecode($match.Groups[1].Value)
      $text = ($text -replace "\s+", " ").Trim()
      if ($text.Length -gt 0) {
        $lines.Add($text)
      }
    }
  }

  if ($lines.Count -eq 0) {
    throw "No text was extracted from HWPX."
  }

  $dir = Split-Path -Parent $outPath
  if ($dir) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
  Set-Content -LiteralPath $outPath -Value ($lines -join [Environment]::NewLine) -Encoding UTF8
  Write-Host "Extracted $($lines.Count) text fragments from HWPX."
}
finally {
  if (Test-Path -LiteralPath $work) {
    Remove-Item -LiteralPath $work -Recurse -Force
  }
}
