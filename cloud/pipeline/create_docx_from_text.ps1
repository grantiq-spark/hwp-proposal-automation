param(
  [Parameter(Mandatory = $true)]
  [string]$InputText,

  [Parameter(Mandatory = $true)]
  [string]$OutputDocx
)

$ErrorActionPreference = "Stop"
$inputPath = (Resolve-Path -LiteralPath $InputText).Path
$outputPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDocx)
$outputDir = Split-Path -Parent $outputPath
if ($outputDir) { New-Item -ItemType Directory -Path $outputDir -Force | Out-Null }

$word = $null
$doc = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $doc = $word.Documents.Add()
  $selection = $word.Selection
  $lines = Get-Content -LiteralPath $inputPath -Encoding UTF8

  foreach ($line in $lines) {
    $text = $line.TrimEnd()
    if ($text -match "^###\s+(.+)$") {
      $selection.Style = "Heading 3"
      $selection.TypeText($matches[1])
    } elseif ($text -match "^##\s+(.+)$") {
      $selection.Style = "Heading 2"
      $selection.TypeText($matches[1])
    } elseif ($text -match "^#\s+(.+)$") {
      $selection.Style = "Heading 1"
      $selection.TypeText($matches[1])
    } elseif ($text.Length -eq 0) {
      $selection.Style = "Normal"
    } else {
      $selection.Style = "Normal"
      $selection.TypeText($text)
    }
    $selection.TypeParagraph()
  }

  $doc.SaveAs([ref]$outputPath, [ref]16)
  Write-Host "Created DOCX: $outputPath"
}
finally {
  if ($doc -ne $null) { $doc.Close([ref]$false) | Out-Null }
  if ($word -ne $null) { $word.Quit() | Out-Null }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
