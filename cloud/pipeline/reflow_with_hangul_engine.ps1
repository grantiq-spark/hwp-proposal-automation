param(
    [string]$InputHwpx = "",
    [string]$OutputHwpx = ""
)

$ErrorActionPreference = "Stop"

function Write-Step {
    param([string]$Message)
    Write-Host "[KBio HWP Reflow] $Message"
}

if ([string]::IsNullOrWhiteSpace($InputHwpx)) {
    $InputHwpx = Join-Path $PSScriptRoot "KBio_ODA_FS_safe_plain.hwpx"
}
elseif (![System.IO.Path]::IsPathRooted($InputHwpx)) {
    $InputHwpx = Join-Path (Get-Location).Path $InputHwpx
}

if ([string]::IsNullOrWhiteSpace($OutputHwpx)) {
    $OutputHwpx = Join-Path $PSScriptRoot "KBio_ODA_FS_hangul_reflowed.hwpx"
}
elseif (![System.IO.Path]::IsPathRooted($OutputHwpx)) {
    $OutputHwpx = Join-Path (Get-Location).Path $OutputHwpx
}

if (!(Test-Path -LiteralPath $InputHwpx)) {
    throw "Input HWPX not found: $InputHwpx"
}

if (Test-Path -LiteralPath $OutputHwpx) {
    Remove-Item -LiteralPath $OutputHwpx -Force
}

$hwp = $null
try {
    Write-Step "Starting Hancom Hangul..."
    $hwp = New-Object -ComObject "HWPFrame.HwpObject"
    $hwp.XHwpWindows.Item(0).Visible = $true

    try {
        $hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule") | Out-Null
    }
    catch {
        Write-Step "FilePathCheck module registration skipped."
    }

    Write-Step "Opening: $InputHwpx"
    $hwp.Open($InputHwpx, "HWPX", "") | Out-Null

    Write-Step "Forcing document layout refresh..."
    foreach ($action in @("MoveDocBegin", "MoveDocEnd", "MoveDocBegin")) {
        try { $hwp.HAction.Run($action) | Out-Null } catch {}
    }

    Start-Sleep -Seconds 3

    Write-Step "Saving reflowed copy: $OutputHwpx"
    $hwp.SaveAs($OutputHwpx, "HWPX", "") | Out-Null
    Write-Step "Done."
}
finally {
    if ($hwp -ne $null) {
        try { $hwp.Quit() | Out-Null } catch {}
        [Runtime.InteropServices.Marshal]::ReleaseComObject($hwp) | Out-Null
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
