# 🚀 Schnellstart-Anleitung - DocxTpl Automatisierung v7.0

## ⚡ In 5 Minuten zur ersten automatisierten Dokumentenerstellung

### Schritt 1: Excel-Datei erstellen

#### 1.1 Neue Excel-Datei öffnen
- Öffnen Sie Excel
- Erstellen Sie eine neue Datei
- Speichern Sie sie als `meine_daten.xlsx`

#### 1.2 Hauptdaten eingeben (Blatt 1)
```
A1: 1
A2: 2  
A3: anlage_seriennummer    B3: projekt_name    C3: standort    D3: betreiber
A4: 1151028                B4: Windpark Nord   C4: Hamburg    D4: Energie AG
A5: 1151029                B5: Windpark Süd    C5: München    D5: Stadtwerke
```

#### 1.3 Fallback-Daten eingeben (Blatt 2)
```
A1: 1
B1: projekt_name           C1: Standard-Projekt
A2: 2  
B2: betreiber             C2: Unbekannter Betreiber
```

### Schritt 2: Word-Vorlage erstellen

#### 2.1 Neue Word-Datei erstellen
- Öffnen Sie Word
- Erstellen Sie ein neues Dokument
- Speichern Sie es als `B_Beschilderung.docx`

#### 2.2 Platzhalter einfügen
```
BESCHILDERUNG - ANLAGE {{ anlage_seriennummer }}

Projekt: {{ projekt_name }}
Standort: {{ standort }}
Betreiber: {{ betreiber }}

Erstellt am: {{ datetime_utc }}
```

#### 2.3 Ordnerstruktur erstellen
```
Vorlagen-Ordner/
└── Anlagen/
    └── B_Beschilderung.docx
```

### Schritt 3: Software starten und konfigurieren

#### 3.1 Anwendung öffnen
- Doppelklick auf `DocxTpl_Automatisierung_v7.0.exe`

#### 3.2 Pfade einstellen
1. **Excel-Datei**: Klicken Sie auf "..." → Wählen Sie `meine_daten.xlsx`
2. **Vorlagen-Ordner**: Klicken Sie auf "..." → Wählen Sie `Vorlagen-Ordner`
3. **Export-Ordner**: Klicken Sie auf "..." → Wählen Sie `C:\Export`

#### 3.3 Konfiguration testen
- Klicken Sie auf **"Konfiguration prüfen"**
- Prüfen Sie die Logs im Tab "Logs"

### Schritt 4: Dokumente erstellen

#### 4.1 Verarbeitung starten
- Klicken Sie auf **"Start"**
- Warten Sie, bis die Verarbeitung abgeschlossen ist

#### 4.2 Ergebnisse prüfen
- Klicken Sie auf **"Export-Ordner öffnen"**
- Sie finden die erstellten Dokumente in `C:\Export\YYYY-MM-DD_HHMMSS_Standard-Projekt\`

---

## 📋 Häufige Anwendungsfälle

### Fall 1: Einfache Beschilderung

**Excel-Daten:**
| anlage_seriennummer | projekt_name | standort |
|-------------------|--------------|----------|
| 1151028 | Windpark Nord | Hamburg |

**Word-Vorlage:**
```
ANLAGE {{ anlage_seriennummer }}
Projekt: {{ projekt_name }}
Standort: {{ standort }}
```

### Fall 2: Mit Bildern

**Excel-Daten:**
| anlage_seriennummer | logo_bild | logo_bild_size |
|-------------------|-----------|----------------|
| 1151028 | logo.png | 8 |

**Word-Vorlage:**
```
ANLAGE {{ anlage_seriennummer }}
Logo: {{ logo_bild }}
```

**Bilder-Ordner:**
```
Bilder/
└── logo.png
```

### Fall 3: Mit QR-Codes

**Excel-Daten:**
| anlage_seriennummer | qr_link | qr_link_size |
|-------------------|---------|--------------|
| 1151028 | https://example.com/1151028 | 4 |

**Word-Vorlage:**
```
ANLAGE {{ anlage_seriennummer }}
QR-Code: {{ qr_link }}
```

### Fall 4: Bedingte Anzeige

**Excel-Daten:**
| anlage_seriennummer | anlage_typ | leistung |
|-------------------|------------|---------|
| 1151028 | WEA | 3000 |

**Word-Vorlage:**
```
ANLAGE {{ anlage_seriennummer }}

{% if anlage_typ == "WEA" %}
Windenergieanlage
Leistung: {{ leistung }} kW
{% else %}
Andere Anlage
{% endif %}
```

---

## 🔧 Schnelle Problemlösung

### Problem: "Excel-Datei nicht gefunden"
**Lösung:** Prüfen Sie den Pfad und verwenden Sie den "..."-Button

### Problem: "Fehlende Platzhalter"
**Lösung:** 
1. "Konfiguration prüfen" ausführen
2. Fehlende Spalten zur Excel hinzufügen

### Problem: "Keine Datensätze gefunden"
**Lösung:** 
1. Prüfen Sie, ob `anlage_seriennummer` in Zeile 3 steht
2. Prüfen Sie, ob die Spalte Daten enthält

### Problem: "Bild nicht gefunden"
**Lösung:** 
1. Prüfen Sie den Bilder-Ordner-Pfad
2. Stellen Sie sicher, dass alle Bilder vorhanden sind

---

## 📝 Nützliche Tipps

### Excel-Tipps
- **Header-Zeile**: Standardmäßig Zeile 3, änderbar in der GUI
- **Pflichtfeld**: `anlage_seriennummer` muss in jeder Zeile stehen
- **Fallback**: Blatt 2 für Standardwerte verwenden

### Word-Tipps
- **Präfixe**: `B_` für Beschilderung, `BA_` für Betriebsanweisung
- **Platzhalter**: Verwenden Sie `{{ variablen_name }}`
- **Bedingungen**: `{% if bedingung %}...{% endif %}`

### GUI-Tipps
- **Themes**: Wechseln Sie zwischen Hell, Dunkel und Girly
- **Logs**: Verwenden Sie den Log-Tab für Fehlerdiagnose
- **Kategorien**: Organisieren Sie Dokumente mit Präfixen

---

## 🎯 Nächste Schritte

Nach dem ersten erfolgreichen Test:

1. **Erweitern Sie die Excel-Daten** um weitere Spalten
2. **Erstellen Sie weitere Word-Vorlagen** für verschiedene Dokumenttypen
3. **Fügen Sie Bilder und QR-Codes** hinzu
4. **Nutzen Sie bedingte Anzeige** für komplexere Logik
5. **Organisieren Sie mit Kategorien** für bessere Übersicht

---

## 📞 Hilfe

- **Detaillierte Anleitung**: Siehe `BEDIENUNGSANLEITUNG.md`
- **Build-Anleitung**: Siehe `BUILD_ANLEITUNG.md`
- **Logs**: Verwenden Sie den Log-Tab für Fehlerdiagnose
- **Konfiguration prüfen**: Testet alle Einstellungen vor der Verarbeitung 