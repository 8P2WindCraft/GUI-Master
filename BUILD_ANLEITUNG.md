# GUI mit Themes - EXE Build Anleitung

## Übersicht
Diese Anleitung erklärt, wie Sie die GUI-Anwendung mit integrierten Themes in eine ausführbare EXE-Datei verwandeln.

## Voraussetzungen

### 1. Python-Umgebung
- Python 3.8 oder höher
- Alle Abhängigkeiten aus `requirements.txt` installiert

### 2. PyInstaller
```bash
pip install pyinstaller
```

### 3. Projektdateien
Stellen Sie sicher, dass folgende Dateien vorhanden sind:
- `app.py` - Hauptanwendung
- `dark.qss` - Dark Theme
- `girly.qss` - Girly Theme
- `Logo.ico` - Anwendungs-Icon
- `Pictures/` - Ordner mit Bildern
- `pyside7.0_gui.spec` - PyInstaller-Konfiguration

## Build-Optionen

### Option 1: Einfacher Build (Empfohlen)
```bash
# Doppelklick auf:
build_pyside7_gui.bat
```

### Option 2: Erweiterter Build mit Optionen
```bash
# Doppelklick auf:
build_advanced.bat
```

Wählen Sie dann eine der folgenden Optionen:
1. **Standard Build** - Optimiert für normale Nutzung
2. **Debug Build** - Mit Konsolenfenster für Fehlerdiagnose
3. **One-File Build** - Einzelne EXE-Datei (größer, aber einfacher zu verteilen)
4. **One-Directory Build** - Mehrere Dateien (schneller, aber komplexer)

### Option 3: Manueller Build
```bash
pyinstaller --clean pyside7.0_gui.spec
```

## Build-Prozess

### 1. Automatische Prüfungen
Das Build-Skript prüft automatisch:
- Python-Installation
- PyInstaller-Installation
- Vorhandensein aller Projektdateien
- Theme-Dateien (dark.qss, girly.qss)

### 2. Build-Ausführung
- Bereinigung vorheriger Builds
- Kompilierung der Anwendung
- Einbindung aller Ressourcen (Themes, Bilder, Icons)
- Erstellung der EXE-Datei

### 3. Ergebnis
Die fertige EXE-Datei befindet sich im `dist/`-Ordner:
- **Standard**: `DocxTpl_Automatisierung_v7.0.exe`
- **Debug**: `DocxTpl_Automatisierung_v7.0.exe` (mit Konsolenfenster)

## Enthaltene Features

### Themes
- **Light Theme** (Standard) - Integriert in die Anwendung
- **Dark Theme** - Aus `dark.qss` geladen
- **Girly Theme** - Aus `girly.qss` geladen

### Ressourcen
- Bilder-Ordner (`Pictures/`)
- Anwendungs-Icon (`Logo.ico`)
- Logo-Bild (`Logo.png`)

### Funktionalitäten
- Vollständige Desktop-GUI (PySide6)
- Excel-Datenverarbeitung
- Word-Dokumentenerstellung
- QR-Code-Generierung
- SVG-Konvertierung
- Theme-Wechsel zur Laufzeit

## Fehlerbehebung

### Häufige Probleme

#### 1. "PyInstaller nicht gefunden"
```bash
pip install pyinstaller
```

#### 2. "Fehlende Abhängigkeiten"
```bash
pip install -r requirements.txt
```

#### 3. "Theme-Dateien nicht gefunden"
Stellen Sie sicher, dass `dark.qss` und `girly.qss` im Projektordner liegen.

#### 4. "Python nicht im PATH"
Fügen Sie Python zum System-PATH hinzu oder verwenden Sie den vollständigen Pfad.

### Debug-Modus
Verwenden Sie Option 2 (Debug Build) im erweiterten Build-Skript, um Fehlermeldungen zu sehen.

## Verteilung

### Einzelne EXE-Datei
- Verwenden Sie Option 3 (One-File Build)
- Die EXE-Datei enthält alle Abhängigkeiten
- Größere Dateigröße, aber einfache Verteilung

### Ordner-basierte Verteilung
- Verwenden Sie Option 4 (One-Directory Build)
- Mehrere Dateien, aber kleinere Gesamtgröße
- Ordner `dist/DocxTpl_Automatisierung_v7.0/` komplett verteilen

## Performance-Optimierung

### Build-Optimierungen
- `--clean` - Bereinigt vorherige Builds
- `--onefile` - Erstellt einzelne EXE-Datei
- `--noconsole` - Versteckt Konsolenfenster
- `--icon` - Setzt Anwendungs-Icon

### Laufzeit-Optimierungen
- Themes werden zur Build-Zeit eingebunden
- Bilder werden komprimiert
- Unnötige Module werden ausgeschlossen

## Support

Bei Problemen:
1. Prüfen Sie die Fehlermeldungen im Debug-Modus
2. Stellen Sie sicher, dass alle Abhängigkeiten installiert sind
3. Überprüfen Sie die Python-Version (3.8+ empfohlen)
4. Testen Sie die Anwendung zuerst im Python-Modus

## Changelog

### Version 7.0
- Desktop-GUI mit Themes
- Integrierte Theme-Unterstützung
- Optimierte PyInstaller-Konfiguration
- Erweiterte Build-Skripte
- Verbesserte Fehlerbehandlung 