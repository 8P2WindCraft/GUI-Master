# DocxTpl Automatisierung v7.0 - Bedienungsanleitung

## 📋 Inhaltsverzeichnis
1. [Übersicht](#übersicht)
2. [Installation und Start](#installation-und-start)
3. [Excel-Datei vorbereiten](#excel-datei-vorbereiten)
4. [Word-Vorlagen erstellen](#word-vorlagen-erstellen)
5. [GUI-Bedienung](#gui-bedienung)
6. [Themes und Anpassungen](#themes-und-anpassungen)
7. [Fehlerbehebung](#fehlerbehebung)
8. [Beispiele](#beispiele)

---

## 🎯 Übersicht

Die **DocxTpl Automatisierung v7.0** ist eine Software zur automatischen Erstellung von Word-Dokumenten aus Excel-Daten. Sie können:

- **Excel-Daten** als Quelle für Dokumente verwenden
- **Word-Vorlagen** mit Platzhaltern erstellen
- **Automatisch** Dokumente für jeden Excel-Eintrag generieren
- **Themes** (Hell, Dunkel, Girly) verwenden
- **Bilder und QR-Codes** einbetten

---

## 🚀 Installation und Start

### Voraussetzungen
- Windows 10/11
- Excel (für die Datenquelle)
- Word (für Vorlagen)

### Start der Anwendung
1. **EXE-Datei**: Doppelklick auf `DocxTpl_Automatisierung_v7.0.exe`
2. **Python-Modus**: `python pyside7.0_gui.py`

### Erste Schritte
1. Öffnen Sie die Anwendung
2. Navigieren Sie zum Tab **"Hauptsteuerung"**
3. Konfigurieren Sie die Pfade (siehe unten)

---

## 📊 Excel-Datei vorbereiten

### Struktur der Excel-Datei

Die Excel-Datei benötigt **zwei Blätter**:

#### Blatt 1: Hauptdaten (Anlagen)
- **Header-Zeile**: Zeile 3 (standardmäßig)
- **Spalten**: Jede Spalte wird zu einer Textmarke
- **Zeilen**: Jede Zeile mit `anlage_seriennummer` wird zu einem Dokument

#### Blatt 2: Fallback-Daten (Allgemein)
- **Spalte B**: Textmarken-Name
- **Spalte C**: Standard-Wert
- **Verwendung**: Wird verwendet, wenn Hauptdaten leer sind

### Beispiel-Excel-Struktur

#### Blatt 1 - Hauptdaten:
| A | B | C | D | E |
|---|---|---|---|---|
| 1 | 2 | 3 | 4 | 5 |
| 6 | 7 | 8 | 9 | 10 |
| **anlage_seriennummer** | **projekt_name** | **anlage_typ** | **standort** | **betreiber** |
| 1151028 | Windpark Nord | WEA-3.0 | Hamburg | Energie AG |
| 1151029 | Windpark Süd | WEA-3.0 | München | Stadtwerke |
| 1151030 | Solarpark Ost | PV-500kW | Berlin | Solar GmbH |

#### Blatt 2 - Fallback-Daten:
| A | B | C |
|---|---|---|
| 1 | **projekt_name** | Standard-Projekt |
| 2 | **betreiber** | Unbekannter Betreiber |
| 3 | **standort** | Unbekannter Standort |

### Wichtige Regeln für Excel-Daten

#### 1. Pflichtfeld
- **`anlage_seriennummer`**: Muss in jeder Zeile vorhanden sein
- **Format**: Text oder Zahl (wird automatisch zu Text konvertiert)

#### 2. Bild-Referenzen
- **Dateiname**: Nur der Dateiname (z.B. `logo.png`)
- **Unterstützte Formate**: `.png`, `.jpg`, `.jpeg`, `.gif`, `.svg`
- **Größe**: Über `_img_size` Platzhalter steuerbar

#### 3. QR-Code-Referenzen
- **Suffix**: `_qr` oder `_link`
- **Inhalt**: URL oder Text für QR-Code
- **Größe**: Über `_qr_size` oder `_link_size` steuerbar

#### 4. Datum und Zeit
- **Automatisch**: `datetime_utc` wird automatisch hinzugefügt
- **Format**: Konfigurierbar in der GUI

### Beispiel für erweiterte Excel-Daten

| anlage_seriennummer | projekt_name | logo_bild | logo_bild_size | qr_link | qr_link_size |
|-------------------|--------------|-----------|----------------|---------|--------------|
| 1151028 | Windpark Nord | logo.png | 10 | https://example.com/1151028 | 5 |
| 1151029 | Windpark Süd | logo.png | 8 | https://example.com/1151029 | 4 |

---

## 📝 Word-Vorlagen erstellen

### Grundlagen der Platzhalter

Platzhalter verwenden **Jinja2-Syntax**:

```jinja2
{{ variablen_name }}
{{ projekt_name }}
{{ anlage_seriennummer }}
```

### Ordnerstruktur für Vorlagen

```
Vorlagen-Ordner/
├── Anlagen/
│   ├── B_Beschilderung.docx
│   ├── BA_Betriebsanweisung.docx
│   └── ...
└── Allgemein/
    ├── Projekt_Übersicht.docx
    └── ...
```

### Kategorisierung durch Präfixe

| Präfix | Kategorie-Ordner | Verwendung |
|--------|------------------|------------|
| `B_` | Beschilderung | Anlagenspezifische Beschilderung |
| `BA_` | Betriebsanweisung | Betriebsanweisungen |
| `A_` | Allgemein | Allgemeine Dokumente |

### Platzhalter-Typen

#### 1. Text-Platzhalter
```jinja2
Projekt: {{ projekt_name }}
Anlage: {{ anlage_seriennummer }}
Standort: {{ standort }}
```

#### 2. Bild-Platzhalter
```jinja2
{{ logo_bild }}
{{ diagramm_bild }}
```

**Größensteuerung:**
```jinja2
{{ logo_bild_size }}  {# Größe in cm #}
```

#### 3. QR-Code-Platzhalter
```jinja2
{{ qr_link }}
{{ anmeldung_qr }}
```

**Größensteuerung:**
```jinja2
{{ qr_link_size }}  {# Größe in cm #}
```

#### 4. Automatische Platzhalter
```jinja2
{{ datetime_utc }}  {# Aktueller Zeitstempel #}
```

### Erweiterte Jinja2-Features

#### Bedingte Anzeige
```jinja2
{% if anlage_typ == "WEA" %}
    Windenergieanlage: {{ anlage_seriennummer }}
{% else %}
    Andere Anlage: {{ anlage_seriennummer }}
{% endif %}
```

#### Schleifen
```jinja2
{% for i in range(1, 6) %}
    Punkt {{ i }}: {{ "punkt_" + i|string }}
{% endfor %}
```

#### Formatierung
```jinja2
{{ projekt_name|upper }}  {# Großbuchstaben #}
{{ standort|title }}     {# Erste Buchstaben groß #}
```

### Beispiel-Word-Vorlage

**Datei: `B_Beschilderung.docx`**

```
BESCHILDERUNG - ANLAGE {{ anlage_seriennummer }}

Projekt: {{ projekt_name }}
Standort: {{ standort }}
Betreiber: {{ betreiber }}

{% if logo_bild %}
Logo: {{ logo_bild }}
{% endif %}

QR-Code für Anmeldung: {{ anmeldung_qr }}

Erstellt am: {{ datetime_utc }}
```

---

## 🖥️ GUI-Bedienung

### Tab 1: Hauptsteuerung

#### Pfade konfigurieren
1. **Excel-Datei**: Klicken Sie auf "..." und wählen Sie Ihre Excel-Datei
2. **Vorlagen-Ordner**: Wählen Sie den Ordner mit Ihren Word-Vorlagen
3. **Bilder-Ordner** (optional): Ordner mit Bildern für die Vorlagen
4. **Export-Ordner**: Wo die fertigen Dokumente gespeichert werden

#### Einstellungen
- **Header-Zeile**: In welcher Zeile stehen die Spaltenüberschriften? (Standard: 3)
- **SVG-Skala**: Skalierung für SVG-Bilder (Standard: 3)
- **PNG-Kompression**: Kompressionslevel für PNG-Bilder (Standard: -1)
- **Zeitstempel-Format**: Format für `datetime_utc` (Standard: `%Y-%m-%d %H:%M:%S UTC`)

#### Excel-Daten anzeigen
- Die Tabelle zeigt eine Vorschau Ihrer Excel-Daten
- Prüfen Sie, ob alle Daten korrekt geladen wurden

#### Verarbeitung starten
1. **Konfiguration prüfen**: Testet alle Einstellungen ohne Dokumente zu erstellen
2. **Start**: Erstellt alle Dokumente
3. **Abbrechen**: Bricht laufende Verarbeitung ab

### Tab 2: Kategorien

#### Kategorien verwalten
- **Präfix**: Text, der am Anfang des Vorlagennamens steht
- **Ordnername**: Name des Ausgabeordners

#### Standard-Kategorien
- `B_` → Beschilderung
- `BA_` → Betriebsanweisung

#### Neue Kategorie hinzufügen
1. Klicken Sie auf "+"
2. Geben Sie Präfix und Ordnername ein
3. Klicken Sie auf "Kategorien speichern"

### Tab 3: Logs

#### Log-Filter
- **ALLE**: Zeigt alle Log-Einträge
- **INFO**: Nur Informationsmeldungen
- **WARN**: Nur Warnungen
- **ERROR**: Nur Fehler
- **SUCCESS**: Nur Erfolgsmeldungen

#### Kontext-Menü
Rechtsklick auf Log-Einträge bietet:
- **Dateien öffnen**: Öffnet referenzierte Dateien
- **Ordner öffnen**: Öffnet Export-Ordner
- **Textmarken hinzufügen**: Fügt fehlende Textmarken zur Excel hinzu

---

## 🎨 Themes und Anpassungen

### Theme wechseln
1. **Menü** → **Ansicht** → **Theme**
2. Wählen Sie:
   - **Hell (Standard)**: Klassisches Design
   - **Dark Mode**: Dunkles Design für bessere Augen
   - **Girly Mode**: Rosa/helles Design

### Einstellungen speichern
- Alle Pfade und Einstellungen werden automatisch gespeichert
- Datei: `settings.json` im Anwendungsordner

### Export-Ordner öffnen
- Nach erfolgreicher Verarbeitung erscheint der Button "Export-Ordner öffnen"
- Klicken Sie darauf, um den Ordner mit den fertigen Dokumenten zu öffnen

---

## 🔧 Fehlerbehebung

### Häufige Probleme

#### 1. "Excel-Datei nicht gefunden"
- **Lösung**: Prüfen Sie den Pfad zur Excel-Datei
- **Tipp**: Verwenden Sie den "..."-Button zum Auswählen

#### 2. "Fehlende Platzhalter in Excel"
- **Problem**: Word-Vorlage verwendet Textmarken, die nicht in Excel stehen
- **Lösung**: 
  1. Führen Sie "Konfiguration prüfen" aus
  2. Fügen Sie fehlende Spalten zur Excel hinzu
  3. Oder verwenden Sie Rechtsklick im Log → "Textmarke(n) im Excel hinzufügen"

#### 3. "Bild nicht gefunden"
- **Problem**: Excel referenziert Bild, das nicht im Bilder-Ordner liegt
- **Lösung**: 
  1. Prüfen Sie den Bilder-Ordner-Pfad
  2. Stellen Sie sicher, dass alle referenzierten Bilder vorhanden sind

#### 4. "Vorlage konnte nicht gelesen werden"
- **Problem**: Word-Vorlage ist beschädigt oder hat ungültige Syntax
- **Lösung**: 
  1. Öffnen Sie die Vorlage in Word
  2. Prüfen Sie die Platzhalter-Syntax
  3. Speichern Sie die Vorlage neu

#### 5. "Keine Datensätze gefunden"
- **Problem**: Excel hat keine Zeilen mit `anlage_seriennummer`
- **Lösung**: 
  1. Prüfen Sie die Header-Zeile (Standard: 3)
  2. Stellen Sie sicher, dass `anlage_seriennummer` als Spaltenname existiert
  3. Prüfen Sie, ob die Spalte Daten enthält

### Debug-Modus
- Verwenden Sie "Konfiguration prüfen" für detaillierte Fehleranalyse
- Logs zeigen genau, wo Probleme auftreten

---

## 📋 Beispiele

### Beispiel 1: Einfache Beschilderung

**Excel-Daten:**
| anlage_seriennummer | projekt_name | standort |
|-------------------|--------------|----------|
| 1151028 | Windpark Nord | Hamburg |

**Word-Vorlage (`B_Beschilderung.docx`):**
```
ANLAGE {{ anlage_seriennummer }}

Projekt: {{ projekt_name }}
Standort: {{ standort }}

Erstellt: {{ datetime_utc }}
```

**Ergebnis:**
```
ANLAGE 1151028

Projekt: Windpark Nord
Standort: Hamburg

Erstellt: 2024-01-15 14:30:25 UTC
```

### Beispiel 2: Mit Bildern und QR-Codes

**Excel-Daten:**
| anlage_seriennummer | logo_bild | logo_bild_size | qr_link | qr_link_size |
|-------------------|-----------|----------------|---------|--------------|
| 1151028 | logo.png | 8 | https://example.com/1151028 | 4 |

**Word-Vorlage:**
```
ANLAGE {{ anlage_seriennummer }}

Logo: {{ logo_bild }}

QR-Code für Anmeldung: {{ qr_link }}

Erstellt: {{ datetime_utc }}
```

### Beispiel 3: Bedingte Anzeige

**Excel-Daten:**
| anlage_seriennummer | anlage_typ | leistung |
|-------------------|------------|---------|
| 1151028 | WEA | 3000 |
| 1151029 | PV | 500 |

**Word-Vorlage:**
```
ANLAGE {{ anlage_seriennummer }}

{% if anlage_typ == "WEA" %}
Windenergieanlage
Leistung: {{ leistung }} kW
{% elif anlage_typ == "PV" %}
Photovoltaikanlage
Leistung: {{ leistung }} kWp
{% else %}
Unbekannte Anlage
{% endif %}
```

---

## 📞 Support

### Bei Problemen
1. **Logs prüfen**: Tab "Logs" zeigt detaillierte Informationen
2. **Konfiguration prüfen**: Testet alle Einstellungen
3. **Debug-Modus**: Verwenden Sie erweiterten Build mit Konsolenfenster

### Nützliche Tipps
- **Backup**: Erstellen Sie Backups Ihrer Excel-Daten und Word-Vorlagen
- **Testen**: Testen Sie mit kleinen Datenmengen vor der Massenverarbeitung
- **Themes**: Wechseln Sie zwischen Themes für bessere Lesbarkeit
- **Kategorien**: Nutzen Sie Kategorien für bessere Organisation

### Version
- **Aktuelle Version**: 7.0
- **Letzte Änderung**: Januar 2024
- **Kompatibilität**: Windows 10/11, Excel 2016+, Word 2016+ 