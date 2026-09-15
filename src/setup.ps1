param(
    [switch]$Force,
    [string]$VenvName = '.venv',
    [string]$PythonVersion = '3.13'
)

$ErrorActionPreference = 'Stop'
$ProjectRoot = Resolve-Path (Join-Path $PSScriptRoot '..')
$repo = Split-Path $ProjectRoot -Leaf
$venvPath = Join-Path $ProjectRoot $VenvName

if ($Force -and (Test-Path $venvPath)) {
    Write-Host "[$repo] Entferne bestehende $VenvName ..."
    Remove-Item $venvPath -Recurse -Force
}

if (-not (Test-Path $venvPath)) {
    Write-Host "[$repo] Erstelle $VenvName ..."
    & py -$PythonVersion -m venv $venvPath
    if ($LASTEXITCODE -ne 0) {
        throw "Python $PythonVersion konnte nicht für venv verwendet werden."
    }
} else {
    Write-Host "[$repo] $VenvName bereits vorhanden."
}

$pythonExe = Join-Path $venvPath 'Scripts\python.exe'
& $pythonExe -m pip install --upgrade pip
& $pythonExe -m pip install -r (Join-Path $PSScriptRoot 'requirements.txt')
Write-Host "[$repo] Setup abgeschlossen." -ForegroundColor Green
