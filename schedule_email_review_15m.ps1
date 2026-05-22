param(
  [string]$TaskName = "Codex-EmailReview-VentaBooking-15m"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$runner = Join-Path $root "run_email_review.ps1"

if (-not (Test-Path -LiteralPath $runner)) {
  throw "No existe $runner"
}

$runnerEscaped = $runner.Replace('"', '\"')
$cmd = 'schtasks /Create /TN "' + $TaskName + '" /TR "powershell.exe -NoProfile -ExecutionPolicy Bypass -File \"' + $runnerEscaped + '\"" /SC MINUTE /MO 15 /F'

cmd /c $cmd | Out-Null
if ($LASTEXITCODE -ne 0) {
  throw "No se pudo crear/actualizar la tarea $TaskName"
}

$taskInfo = schtasks /Query /TN $TaskName /V /FO LIST
$taskInfo | Out-String | Write-Host
