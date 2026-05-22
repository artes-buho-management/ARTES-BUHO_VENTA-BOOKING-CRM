param(
  [ValidateSet("create", "clone", "login", "status", "pull", "push", "open", "version", "deploy", "run", "watch", "logs")]
  [string]$Action = "status",
  [string]$SheetId = "",
  [string]$ScriptId = "",
  [string]$Title = "CRM VENTA-BOOKING",
  [string]$FunctionName = ""
)

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$claspLocal = Join-Path $repoRoot "node_modules\@google\clasp\build\src\index.js"

if (-not (Test-Path -LiteralPath $claspLocal)) {
  throw "No se encontro clasp local en $claspLocal. Ejecuta npm.cmd install."
}

function Resolve-NodeExe {
  $cmd = Get-Command node.exe -ErrorAction SilentlyContinue
  if ($cmd -and $cmd.Source) {
    return $cmd.Source
  }

  $candidates = @(
    "C:\Program Files\nodejs\node.exe",
    "C:\Program Files (x86)\nodejs\node.exe",
    "C:\Progra~1\nodejs\node.exe"
  )

  foreach ($candidate in $candidates) {
    if (Test-Path -LiteralPath $candidate) {
      return $candidate
    }
  }

  throw "No se encontro node.exe"
}

function Invoke-Clasp {
  param([string[]]$ClaspArgs)
  & $nodeExe $claspLocal @ClaspArgs
  if ($LASTEXITCODE -ne 0) {
    throw "clasp $($ClaspArgs -join ' ') fallo"
  }
}

$nodeExe = Resolve-NodeExe

switch ($Action) {
  "create" {
    if ([string]::IsNullOrWhiteSpace($SheetId)) {
      throw "Debes indicar -SheetId para crear y vincular el proyecto."
    }
    Invoke-Clasp -ClaspArgs @("create", "--title", $Title, "--parentId", $SheetId)
  }
  "clone" {
    if ([string]::IsNullOrWhiteSpace($ScriptId)) {
      throw "Debes indicar -ScriptId para clonar un proyecto existente."
    }
    Invoke-Clasp -ClaspArgs @("clone", $ScriptId)
  }
  "run" {
    if ([string]::IsNullOrWhiteSpace($FunctionName)) {
      throw "Debes indicar -FunctionName para -Action run."
    }
    Invoke-Clasp -ClaspArgs @("run", $FunctionName)
  }
  "watch" {
    Invoke-Clasp -ClaspArgs @("push", "--watch")
  }
  "logs" {
    Invoke-Clasp -ClaspArgs @("logs", "--watch")
  }
  default {
    Invoke-Clasp -ClaspArgs @($Action)
  }
}
