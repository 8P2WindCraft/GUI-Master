# DocxTpl Automatisierung - App starten
# Voraussetzung: setup_env.ps1 wurde ausgefuehrt

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

& ".venv\Scripts\Activate.ps1"
$env:PYTHONUTF8 = "1"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
python app.py
