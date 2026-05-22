param(
  [string]$Branch = "main",
  [int]$DebounceSeconds = 20
)

$ErrorActionPreference = "Stop"
$repoPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$gitExe = "git"
$nodeExe = "node"
$claspLocal = Join-Path $repoPath "node_modules\@google\clasp\build\src\index.js"
$pending = $false
$lastChange = Get-Date

function Write-Log {
  param([string]$Message)
  Write-Host ("[REALTIME] " + $Message)
}

function Invoke-ClaspPush {
  & $nodeExe $claspLocal push
  if ($LASTEXITCODE -ne 0) {
    throw "clasp push fallo"
  }
}

function Invoke-GitPush {
  & $gitExe -c "safe.directory=$repoPath" add -A
  if ($LASTEXITCODE -ne 0) {
    throw "git add fallo"
  }

  & $gitExe -c "safe.directory=$repoPath" diff --cached --quiet
  if ($LASTEXITCODE -eq 0) {
    Write-Log "Sin cambios para GitHub."
    return
  }

  $msg = "sync(realtime): " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
  & $gitExe -c "safe.directory=$repoPath" commit -m $msg
  if ($LASTEXITCODE -ne 0) {
    throw "git commit fallo"
  }

  & $gitExe -c "safe.directory=$repoPath" push origin $Branch
  if ($LASTEXITCODE -ne 0) {
    throw "git push fallo"
  }
}

if (-not (Test-Path -LiteralPath $claspLocal)) {
  throw "No se encontro clasp local en $claspLocal. Ejecuta npm.cmd install."
}

Write-Log "Vigilando cambios en $repoPath"
Write-Log "Debounce: $DebounceSeconds segundos"

$fsw = New-Object System.IO.FileSystemWatcher
$fsw.Path = $repoPath
$fsw.IncludeSubdirectories = $true
$fsw.EnableRaisingEvents = $true
$fsw.NotifyFilter = [System.IO.NotifyFilters]'FileName, LastWrite, DirectoryName'
$fsw.Filter = "*.*"

$callback = {
  param($sender, $eventArgs)
  $path = $eventArgs.FullPath
  if ($path -match "\\node_modules\\") { return }
  if ($path -match "\\.git\\") { return }
  if ($path -match "\\reports\\") { return }
  $script:pending = $true
  $script:lastChange = Get-Date
}

$handlers = @()
$handlers += Register-ObjectEvent -InputObject $fsw -EventName Changed -Action $callback
$handlers += Register-ObjectEvent -InputObject $fsw -EventName Created -Action $callback
$handlers += Register-ObjectEvent -InputObject $fsw -EventName Deleted -Action $callback
$handlers += Register-ObjectEvent -InputObject $fsw -EventName Renamed -Action $callback

try {
  while ($true) {
    Start-Sleep -Seconds 2
    if (-not $pending) { continue }
    $elapsed = (Get-Date) - $lastChange
    if ($elapsed.TotalSeconds -lt $DebounceSeconds) { continue }

    $pending = $false
    Write-Log "Cambio detectado. Publicando en Apps Script y GitHub..."
    try {
      Invoke-ClaspPush
      Invoke-GitPush
      Write-Log "Publicacion completada."
    } catch {
      Write-Log ("Error: " + $_.Exception.Message)
      $pending = $true
      $lastChange = Get-Date
    }
  }
}
finally {
  foreach ($h in $handlers) {
    Unregister-Event -SourceIdentifier $h.Name -ErrorAction SilentlyContinue
    Remove-Job -Id $h.Id -Force -ErrorAction SilentlyContinue
  }
  $fsw.Dispose()
}
