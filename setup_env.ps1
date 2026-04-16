# DocxTpl Automatisierung - Umgebung einrichten
# Ausfuehren: .\setup_env.ps1

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ProjectRoot

Write-Host "=== Umgebung einrichten ===" -ForegroundColor Cyan

# Virtuelle Umgebung erstellen (falls nicht vorhanden)
if (-not (Test-Path ".venv")) {
    Write-Host "Erstelle virtuelle Umgebung .venv ..." -ForegroundColor Yellow
    python -m venv .venv
}

# Aktivieren und installieren
Write-Host "Aktiviere .venv und installiere Abhaengigkeiten ..." -ForegroundColor Yellow
& ".venv\Scripts\Activate.ps1"
pip install -r requirements.txt

Write-Host ""
Write-Host "=== Fertig ===" -ForegroundColor Green
Write-Host "App starten mit: .venv\Scripts\Activate.ps1; python app.py" -ForegroundColor White
