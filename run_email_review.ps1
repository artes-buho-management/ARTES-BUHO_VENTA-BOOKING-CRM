param(
  [switch]$DryRun,
  [ValidateSet("none", "mx_a")]
  [string]$DnsMode = "none",
  [ValidateSet("off", "duck")]
  [string]$WebMode = "duck",
  [int]$MaxWebLookups = 2,
  [int]$MaxWebLinks = 0
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$script = Join-Path $root "tools\review_emails_ia.py"
$python = "python"
$report = "C:\Users\elrub\Desktop\CARPETA CODEX\04_TEMPORAL\email_review_report.json"

if (-not (Test-Path -LiteralPath $script)) {
  throw "No existe $script"
}

$args = @(
  $script,
  "--report-file", $report,
  "--dns-mode", $DnsMode,
  "--web-mode", $WebMode,
  "--max-web-lookups", $MaxWebLookups,
  "--max-web-links", $MaxWebLinks
)
if ($DryRun) {
  $args += "--dry-run"
}

function Invoke-PythonWithCapture {
  param(
    [string]$PythonExe,
    [string[]]$ArgumentList
  )

  $prev = $ErrorActionPreference
  try {
    $ErrorActionPreference = "Continue"
    $out = & $PythonExe @ArgumentList 2>&1
    return @{
      ExitCode = [int]$LASTEXITCODE
      Output = @($out | ForEach-Object { $_.ToString() })
    }
  } finally {
    $ErrorActionPreference = $prev
  }
}

$run = Invoke-PythonWithCapture -PythonExe $python -ArgumentList $args
$exitCode = $run.ExitCode
if ($run.Output) { $run.Output | ForEach-Object { Write-Host $_ } }
if ($exitCode -ne 0) {
  $joined = ($run.Output -join "`n")
  $isWritePermIssue =
    ($joined -match "The caller does not have permission") -or
    ($joined -match "insufficientFilePermissions") -or
    ($joined -match "403")

  if (-not $DryRun -and $isWritePermIssue) {
    Write-Warning "[EMAIL-IA] Sin permiso de escritura con la cuenta tecnica. Reintentando en modo SOLO AUDITORIA (DryRun)."
    $argsDry = @(
      $script,
      "--report-file", $report,
      "--dns-mode", $DnsMode,
      "--web-mode", $WebMode,
      "--max-web-lookups", $MaxWebLookups,
      "--max-web-links", $MaxWebLinks,
      "--dry-run"
    )
    $runDry = Invoke-PythonWithCapture -PythonExe $python -ArgumentList $argsDry
    $exitDry = $runDry.ExitCode
    if ($runDry.Output) { $runDry.Output | ForEach-Object { Write-Host $_ } }
    if ($exitDry -ne 0) {
      throw "review_emails_ia.py fallo tambien en DryRun con codigo $exitDry"
    }
    Write-Host "[EMAIL-IA] Fallback DryRun completado report=$report"
    exit 0
  }

  throw "review_emails_ia.py fallo con codigo $exitCode"
}

Write-Host "[EMAIL-IA] OK report=$report"
