param(
  [int]$Port = 8787
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
$Log = Join-Path $Root "server-live.log"
$NodeCandidates = @(
  "C:\Program Files\nodejs\node.exe",
  "C:\Program Files\WindowsApps\OpenAI.Codex_26.429.8261.0_x64__2p2nqsd0c76g0\app\resources\node.exe",
  "C:\Users\UMTR\AppData\Local\OpenAI\Codex\bin\node.exe",
  "node.exe"
)
$Node = $NodeCandidates | Where-Object {
  if ($_ -eq "node.exe") {
    [bool](Get-Command node.exe -ErrorAction SilentlyContinue)
  } else {
    Test-Path -LiteralPath $_
  }
} | Select-Object -First 1

if (-not $Node) {
  throw "Node.js executable was not found."
}

Set-Location $Root
$env:PORT = "$Port"

"& `"$Node`" .\server.js" | Out-File -FilePath $Log -Encoding utf8
& $Node .\server.js *>> $Log
