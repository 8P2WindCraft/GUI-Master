# Projektverwaltung & Team-Portabilität

## 🎯 Übersicht

Das GUI-Master Programm unterstützt jetzt vollständige Projektverwaltung mit automatischer Pfad-Normalisierung für Team-Sharing über SharePoint/OneDrive.

## 📁 Projektverwaltung

### Menü "Projekt"
- **Neu**: Setzt alle Pfade zurück, behält Einstellungen
- **Öffnen...**: Lädt `.dta.json` Projektdatei
- **Speichern**: Speichert an aktuellen Pfad
- **Speichern unter...**: Speichert unter neuem Namen

### Projektformat (.dta.json)
```json
{
    "version": "7.0.1",
    "paths": {
        "excel_path": "C:/Users/{username}/OneDrive/Projekt/data.xlsx",
        "vorlagen_ordner": "C:/Users/{username}/OneDrive/Projekt/templates",
        "bilder_ordner": "C:/Users/{username}/OneDrive/Projekt/images",
        "export_ordner": "C:/Users/{username}/OneDrive/Projekt/exports"
    },
    "settings": {
        "header_row": 3,
        "svg_scale": 5,
        "theme": "Light"
    },
    "categories": {
        "b_": "Beschilderung",
        "ba_": "Betriebsanweisung"
    },
    "saved_at": "2025-01-27 14:30:15"
}
```

## 🔄 Team-Portabilität

### Problem
Verschiedene Benutzer haben verschiedene Pfade:
- Jonas: `C:\Users\JonasBeseler\OneDrive\...`
- Besa: `C:\Users\BeyzaAygündüz\OneDrive\...`

### Lösung: Automatische Pfad-Normalisierung

#### 1. **Intelligente Normalisierung**
- **UI zeigt echte Pfade**: `C:/Users/JonasBeseler/OneDrive/...`
- **Speichern normalisiert automatisch**: `C:/Users/{username}/OneDrive/...`
- **Laden denormalisiert automatisch**: `{username}` → aktueller Benutzer

#### 2. **Button "🔄 Pfade normalisieren"**
- Normalisiert nur Pfade OHNE `{username}`
- Zeigt Feedback: "X Pfade wurden normalisiert"
- Intelligente Erkennung bereits normalisierter Pfade

#### 3. **Unterstützt Umlaute**
- Funktioniert mit: `BeyzaAygündüz`, `JörgMüller`, etc.
- UTF-8 Encoding in allen JSON-Operationen

## 🚀 Workflow für Teams

### Jonas erstellt Projekt:
1. Pfade einrichten → sieht: `C:/Users/JonasBeseler/OneDrive/Projekt/...`
2. "🔄 Pfade normalisieren" klicken (optional, automatisch beim Speichern)
3. **Projekt → Speichern unter...** → `WindparkXY.dta.json`

### Besa öffnet Projekt:
1. **Projekt → Öffnen...** → `WindparkXY.dta.json` wählen
2. ✨ **Automatisch**: Alle Pfade zeigen `C:/Users/BeyzaAygündüz/OneDrive/Projekt/...`
3. Sofort einsatzbereit!

## 💾 Persistente Einstellungen

### settings.json
- Speichert zuletzt verwendete Projekte
- Merkt sich aktuellen Projektpfad
- UTF-8 Encoding für korrekte Umlaute

```json
{
    "excel_path": "C:/Users/{username}/OneDrive/Projekt/data.xlsx",
    "recent_projects": [
        "C:/Users/{username}/OneDrive/Projekt/WindparkXY.dta.json",
        "C:/Users/{username}/OneDrive/Projekt/Projekt2.dta.json"
    ],
    "last_project_path": "C:/Users/{username}/OneDrive/Projekt/WindparkXY.dta.json"
}
```

## ✅ Vorteile

- **Team-Sharing**: Projekte funktionieren auf allen Rechnern
- **Benutzerfreundlich**: Echte Pfade in der UI, keine Platzhalter
- **Automatisch**: Keine manuelle Pfadanpassung nötig
- **Robust**: UTF-8, Umlaute, verschiedene Slash-Richtungen
- **Intelligent**: Nur notwendige Normalisierung

## 🎯 GitHub Repository

Alle Änderungen sind verfügbar unter:
**https://github.com/8P2WindCraft/GUI-Master**

### Letzte Features:
- ✅ Projektverwaltung (Neu/Öffnen/Speichern)
- ✅ Automatische Pfad-Normalisierung
- ✅ Team-Portabilität (SharePoint/OneDrive)
- ✅ UTF-8 Encoding (Umlaute)
- ✅ Intelligente UI-Updates
