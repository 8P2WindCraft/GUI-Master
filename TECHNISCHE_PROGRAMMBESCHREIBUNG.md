# DocxTpl Automatisierung v7.0 – Technische Programmbeschreibung

## Kernzusammenfassung (Executive Summary)

**DocxTpl Automatisierung** ist eine Desktop-Anwendung (PySide6/Qt) zur **massenhaften Erstellung von Word-Dokumenten** aus Excel-Daten und Word-Vorlagen. Ziel ist die Automatisierung von Dokumentationsprozessen im Bereich technische Anlagen (z.B. Beschilderung, Betriebsanweisungen, Notfalldokumente).

**Kernfunktion:** Excel-Daten (je Zeile ein Datensatz) werden mit Word-Vorlagen (.docx) und Jinja2-Platzhaltern verknüpft. Das Programm erzeugt pro Datensatz und pro Vorlage ein fertiges Word-Dokument. Optional werden QR-Codes, Bilder (inkl. SVG→PNG), PDF-Export und kategorisierte Exportordner unterstützt.

**Architektur:** Desktop-GUI (PySide6) + `core/logic.py` (docxtpl, pandas) + optionales Supabase-Backend (Streamlit) für Datenmigration und Kunden-Dateneingabe.

**Zielgruppe der Beschreibung:** Programmierer, die das System erweitern, debuggen oder integrieren sollen.

---

## 1. Architektur-Übersicht

### 1.1 Komponenten

| Komponente | Technologie | Aufgabe |
|------------|-------------|---------|
| `app.py` | PySide6 | Einstieg, QApplication, MainWindow |
| `gui/main_window.py` | PySide6 | Hauptfenster, Tabs, Menüs, Worker-Steuerung |
| `core/logic.py` | docxtpl, pandas, qrcode | Excel-Laden, Vorlagenverarbeitung, Rendering |
| `core/utils.py` | PySide6 | `resource_path()`, SVG→PNG |
| `core/logging.py` | – | Log-Handler (HTML/Console) |
| `core/excel_com.py` | pywin32 (COM) | Excel lesen, wenn Datei gesperrt |
| `Backend/` | Streamlit, Supabase | Datenverwaltung, Kunden-Portal |

### 1.2 Datenfluss

```
Excel (Blatt 1 + 2)  →  lade_excel_daten()
                              ↓
         Datensätze + Fallback-Marken
                              ↓
         Vorlagen (anlagen/ + allgemein/)
                              ↓
         ersetze_platzhalter_mit_docxtpl()
                              ↓
         Export-Ordner (Zeitstempel_Projektname/Kategorie/)
```

### 1.3 Thread-Modell

- GUI läuft im Hauptthread.
- `verarbeite_vorlagen()` läuft in einem `Worker` (`QThread`).
- Signale: `log`, `progress`, `finished`, `current_file`.
- Abbruch via `worker.requestInterruption()` und `worker_thread.isInterruptionRequested()`.

---

## 2. Desktop-GUI – Detaillierte Funktionsbeschreibung

### 2.1 Hauptsteuerung (Tab 1)

| Element | Funktion |
|---------|----------|
| **Excel-Datei** | Pfad zur .xlsx/.xls. Blatt 1 = Datensätze, Blatt 2 = Fallback-Platzhalter. |
| **Vorlagen-Ordner** | Muss Unterordner `anlagen/` und `allgemein/` enthalten. |
| **Bilder-Ordner** | Optional. Bilder werden per Dateinamen aus Excel referenziert. |
| **Export-Ordner** | Basis-Ordner für generierte Dokumente. |
| **Projektname (Override)** | Überschreibt `projekt_name` aus Excel für den Export-Ordner-Namen. |
| **Header-Zeile** | Zeilennummer der Spaltenüberschriften (Standard: 3). |
| **SVG-Skala** | Skalierungsfaktor für SVG→PNG (1–10, Standard 3). |
| **PNG-Kompression** | 0 = maximale Qualität, -1 = Standard. |
| **Zeitstempel-Format** | Python-Format für `datetime_utc` (Standard: `%Y-%m-%d %H:%M:%S UTC`). |
| **Pfade normalisieren** | Ersetzt `C:/Users/<username>/` durch `C:/Users/{username}/` für Team-Portabilität. |
| **PDF umwandeln** | Nach Generierung alle .docx im Export in PDF konvertieren (docx2pdf, Windows). |
| **Start** | Startet `Worker` mit `verarbeite_vorlagen()`. |
| **Konfiguration prüfen** | Trockenlauf ohne Datei-Erstellung. |
| **Schließen** | Beenden oder Worker abbrechen. |

**Excel-Vorschau:** `PandasModel` zeigt Blatt 1 in `QTableView`. Bei gesperrter Datei wird `core.excel_com.read_excel_sheet1_via_com()` genutzt.

### 2.2 Kategorien (Tab 2)

Kategorien steuern die **Export-Ordnerstruktur**:

- **Präfix** (z.B. `b_`, `ba_`): Wenn eine Vorlage mit diesem Präfix beginnt, geht sie in den zugehörigen Kategorie-Ordner.
- **Ordnername** (z.B. „Beschilderung“, „Betriebsanweisung“): Name des Unterordners im Export.

Beispiel: Vorlage `b_Hinweisschild.docx` → Export in `ExportOrdner/Beschilderung/`.

- **+** / **-** : Kategorie hinzufügen/entfernen.
- **Kategorien speichern**: Übernimmt die Änderungen in `self.categories` und `settings.json`.

### 2.3 Bilder-Vorschau (Tab 3)

- Listet alle unterstützten Bilder im Bilder-Ordner (PNG, JPG, GIF, SVG, BMP).
- Bei Auswahl: Vorschau rechts (SVG wird temporär zu PNG konvertiert).
- Wird beim Wechsel zum Tab aktualisiert.

### 2.4 Logs (Tab 4)

- HTML-Logs mit farbcodierten Leveln: INFO, WARN, ERROR, SUCCESS, FATAL, SEP.
- **Filter:** ALLE, INFO, WARN, ERROR, SUCCESS, FATAL, SEP.
- **Kontextmenü:**
  - Pfade aus Log-Zeilen öffnen (Datei/Ordner/Vorlage).
  - Fehlende Platzhalter oder Size-Platzhalter aus Log in Excel eintragen.
  - Text kopieren, Log löschen.

### 2.5 Menüs

**Projekt:**

- Neu, Öffnen, Letzte Projekte, Speichern, Speichern unter.
- Projekte: `*.dta.json` mit Pfaden, Einstellungen, Kategorien.
- Pfad-Normalisierung: Beim Speichern `{username}`, beim Laden aktuelle Benutzername.

**Ansicht → Theme:**

- Light (Standard), Dark (`dark.qss`), Girly (`girly.qss`).

---

## 3. Kernlogik – `core/logic.py`

### 3.1 `lade_excel_daten(pfad, header_row, log_callback)`

- **Blatt 1:** Pandas DataFrame, Header in `header_row`.
- Nur Zeilen mit gültiger `anlage_seriennummer`.
- Spaltennamen werden gestrippt.
- **Blatt 2:** Fallback-Marken. Spalte B = Key, Spalte C = Wert.
- Bei `PermissionError`: Fallback auf `lade_excel_daten_via_com()`.

**Rückgabe:** `(Liste von Datensätzen, dict mit Fallback-Marken)`.

### 3.2 `ersetze_platzhalter_mit_docxtpl(doc_path, context, svg_png_map, log_callback)`

Verarbeitungsreihenfolge der Platzhalter:

1. **Bilder** (`*_img`): Dateiname aus Kontext → Pfad im Bilder-Ordner. SVG nutzt `svg_png_map`. Optional `*_img_size` / `*_Size` für Breite in cm (Standard 15).
2. **QR-Codes** (`*_qr`, `*_link`): URL/Text → QR-Bild. Optional `*_qr_size` / `*_link_size` (Standard 4 cm).
3. **Text**: Alle anderen Platzhalter per `doc.render(render_context)`.

`context['_bilder_ordner']` wird für Bild-Pfade genutzt.

### 3.3 `verarbeite_vorlagen(...)`

Hauptablauf:

1. Excel laden (Datensätze + Fallback).
2. SVG-Referenzen sammeln → temporär in PNG konvertieren (`tempfile.mkdtemp()`).
3. Vorlagen suchen in `anlagen/` und `allgemein/`.
4. Export-Ordner: `export_ordner/YYYY-MM-DD_HHMMSS_Projektname/`.
5. **Anlagen:** Pro Datensatz × Anlagen-Vorlage → Kontext = Datensatz + Fallback, `datetime_utc` setzen, rendern, speichern.
6. **Allgemein:** Einmal mit erstem Datensatz + Fallback.
7. Optional: alle .docx in PDF umwandeln.
8. Temp-Ordner löschen.

**Dateinamen:**
- Anlagen: `{seriennummer}_{vorlagenname ohne Präfix}.docx`
- Allgemein: `{vorlagenname ohne Präfix}.docx`

**Kategorisierung:** Vorlagen-Präfix bestimmt Unterordner (z.B. `b_` → Beschilderung).

### 3.4 `verarbeite_vorlagen_trockenlauf(...)`

- Kein Schreiben von Dateien.
- Prüft: Vorlagen-Platzhalter vs. Excel/Fallback, vorhandene Bild-Referenzen.
- Detaillierte Ausgabe zu `_size`-Platzhaltern und Standardgrößen.

---

## 4. Platzhalter-Syntax (Jinja2/docxtpl)

| Typ | Platzhalter | Kontext-Beispiel | Bemerkung |
|-----|-------------|------------------|-----------|
| Text | `{{platzhalter}}` | `{"platzhalter": "Wert"}` | Direkt ersetzt |
| Bild | `{{xy_img}}` | `{"xy_img": "logo.png"}` | Dateiname im Bilder-Ordner |
| Bildgröße | `{{xy_img_size}}` oder `{{xy_img_Size}}` | `{"xy_img_size": 10}` | Breite in cm |
| QR | `{{xy_qr}}` oder `{{xy_link}}` | `{"xy_link": "https://..."}` | URL oder Text |
| QR-Größe | `{{xy_qr_size}}` | `{"xy_qr_size": 5}` | Breite in cm |
| System | `{{datetime_utc}}` | automatisch | Aktuelles Datum |

Bilder: PNG, JPG, GIF, SVG (SVG wird intern zu PNG konvertiert).

---

## 5. Excel-Struktur

### Blatt 1 (Datensätze)

- Header in konfigurierbarer Zeile (Standard: Zeile 3).
- Mindestens Spalte `anlage_seriennummer`.
- Alle Spaltennamen werden zu Platzhaltern (gestrippt).
- Leere Zeilen werden ignoriert.

### Blatt 2 (Fallback)

- Zeile 1: Überschrift (optional).
- Ab Zeile 2: Spalte B = Platzhalter-Name, Spalte C = Wert.
- Wird genutzt, wenn ein Wert im Datensatz fehlt oder leer ist.

---

## 6. Vorlagen-Ordnerstruktur

```
vorlagen_ordner/
├── anlagen/           # Pro Datensatz einmal pro Vorlage
│   ├── b_Hinweisschild.docx
│   └── ba_Anleitung.docx
└── allgemein/         # Einmal mit erstem Datensatz + Fallback
    └── Projektinfo.docx
```

Präfixe wie `b_`, `ba_` steuern die Kategorie und werden aus dem Dateinamen im Export entfernt.

---

## 7. Projektdatei (`*.dta.json`)

```json
{
  "version": "7.0.1",
  "paths": {
    "excel_path": "C:/Users/{username}/...",
    "vorlagen_ordner": "...",
    "bilder_ordner": "...",
    "export_ordner": "..."
  },
  "settings": {
    "header_row": 3,
    "svg_scale": 3,
    "png_compression": -1,
    "theme": "Light",
    "datetime_utc_format": "%Y-%m-%d %H:%M:%S UTC",
    "projekt_name_override": "",
    "export_as_pdf": false
  },
  "categories": {
    "b_": "Beschilderung",
    "ba_": "Betriebsanweisung"
  },
  "saved_at": "2026-..."
}
```

`{username}` wird beim Laden durch den aktuellen Windows-Benutzernamen ersetzt.

---

## 8. Hilfsmodule

### 8.1 `core/utils.py`

- **`resource_path(relative_path)`:** Nutzt `sys._MEIPASS` für PyInstaller, sonst aktuelles Verzeichnis.
- **`svg_to_png_file_pyside(svg_path, png_path, log_callback, scale, compression)`:** SVG → PNG via QSvgRenderer, QImage, QPainter.

### 8.2 `core/excel_com.py`

- **`read_excel_sheet1_via_com(path, header_row)`:** Liest Blatt 1 per Windows COM, wenn die Datei gesperrt ist.
- **`lade_excel_daten_via_com(path, header_row, log_callback)`:** Vollständiges Laden per COM.

### 8.3 `core/logging.py`

- **`_log_handler(msg, level, callback, is_dark_mode)`:** Formatiert Logs (HTML) und ruft den Callback auf.

---

## 9. Backend (Supabase/Streamlit)

Separates Modul, nicht Teil der Desktop-GUI-Laufzeit.

**Admin-App (`Backend/streamlit_app/`):**

- Home: Metriken, Supabase-Verbindung
- Projekte, Anlagen, Fallback-Marken verwalten
- Dateien hochladen
- Kunden-Links (Token)
- Excel-Import nach Supabase
- Einstellungen (Kategorien)

**Kunden-Frontend (`Backend/kunden_frontend/`):**

- Token-basierter Zugang
- Bearbeitung von Anlagen-Daten (Key-Value)
- Datei-Upload (PNG, JPG, PDF, DOCX)

---

## 10. Abhängigkeiten (requirements.txt)

| Paket | Verwendung |
|-------|-----------|
| PySide6 | GUI |
| docxtpl | Word-Vorlagen, Jinja2 |
| python-docx | Word-API |
| pandas | Excel, Datensätze |
| openpyxl | Excel lesen/schreiben |
| qrcode[pil] | QR-Codes |
| docx2pdf | PDF-Export (Windows) |
| pywin32 | Excel COM (Windows) |
| jinja2 | Template-Syntax |
| lxml | XML |

---

## 11. Hinweise für Programmierer

1. **Worker-Abbruch:** Immer `worker_thread.isInterruptionRequested()` in Schleifen prüfen.
2. **Excel gesperrt:** `PermissionError` → `excel_com`-Fallback nutzen.
3. **Vorlagenstruktur:** `anlagen/` und `allgemein/` sind Pflicht.
4. **Platzhalter-Reihenfolge:** Bilder → QR → Text, um Konflikte zu vermeiden.
5. **SVG-Cache:** Pro Lauf werden alle einzigartigen SVGs einmal konvertiert.
6. **Theme:** `is_dark_mode` beeinflusst die Log-Farben.
7. **EXE-Build:** `pyinstaller pyside7.0_gui.spec`, Einstieg `app.py`.

---

## 12. Typischer Workflow

1. Projekt öffnen oder neu anlegen.
2. Pfade setzen (Excel, Vorlagen, Bilder, Export).
3. Optional: Kategorien anpassen.
4. „Konfiguration prüfen“ (Trockenlauf).
5. „Start“ → Dokumente werden erstellt.
6. Optional: „Export-Ordner öffnen“ oder PDF-Erstellung nutzen.
7. Projekt speichern für Wiederverwendung.

---

*Stand: v7.0.1, Build 2025-01-27*