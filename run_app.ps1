# DocxTpl Automatisierung - App starten
# Voraussetzung: setup_env.ps1 wurde ausgefuehrt

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

& ".venv\Scripts\Activate.ps1"
python app.py
