import io
import os
import shutil
import tempfile
import traceback
from datetime import datetime

import pandas as pd
import qrcode
from docx.shared import Cm
from docxtpl import DocxTemplate, InlineImage
from jinja2.exceptions import TemplateSyntaxError
from openpyxl import load_workbook

from core.logging import _log_handler
from core.utils import svg_to_png_file_pyside


def lade_excel_daten(pfad, header_row=3, log_callback=None):
    """
    Lädt Daten aus einer Excel-Datei.
    - Liest das erste Tabellenblatt in einen Pandas DataFrame. Jede Zeile, die eine
      'anlage_seriennummer' enthält, wird als ein Datensatz (dictionary) behandelt.
    - Liest das zweite Tabellenblatt als Quelle für Fallback-Textmarken. Dies ist
      nützlich für allgemeine Informationen (z.B. Projektname, Datum), die für
      alle Dokumente gleich sind.

    Args:
        pfad (str): Pfad zur Excel-Datei.
        header_row (int): Die Zeilennummer, die die Spaltenüberschriften enthält.
        log_callback (callable, optional): Callback für Log-Nachrichten.

    Returns:
        tuple: Ein Tupel mit (Liste von Datensätzen, Dictionary mit Fallback-Marken).
    """
    _log_handler(f"Lade Excel-Daten aus: {pfad}", "INFO", log_callback)
    if not os.path.isfile(pfad):
        raise FileNotFoundError(f"FEHLER: Excel-Datei nicht gefunden: {pfad}")
    try:
        try:
            df = pd.read_excel(pfad, header=header_row - 1).dropna(how='all')
        except (PermissionError, OSError) as e:
            if getattr(e, 'errno', None) == 13 or isinstance(e, PermissionError):
                from core.excel_com import lade_excel_daten_via_com
                result = lade_excel_daten_via_com(pfad, header_row, log_callback)
                if result is not None:
                    return result
            raise
        df.columns = [str(col).strip() for col in df.columns]
        df = df.map(lambda x: x.strip() if isinstance(x, str) else x)
        datensaetze = [
            row.to_dict()
            for _, row in df.iterrows()
            if 'anlage_seriennummer' in row and pd.notna(row['anlage_seriennummer'])
        ]
        _log_handler(f"Erfolgreich {len(datensaetze)} Datensätze (Zeilen) geladen.", "SUCCESS", log_callback)

        fallback_marken = {}
        wb = None
        try:
            wb = load_workbook(pfad, read_only=True)
            if len(wb.sheetnames) > 1:
                ws = wb[wb.sheetnames[1]]
                for row in ws.iter_rows(min_row=2, values_only=True):
                    if row and len(row) >= 3 and row[1]:
                        key = str(row[1]).strip()
                        value = str(row[2]).strip() if row[2] is not None else ""
                        fallback_marken[key] = value
                _log_handler(
                    f"Erfolgreich {len(fallback_marken)} Fallback-Textmarken aus dem zweiten Blatt geladen.",
                    "SUCCESS",
                    log_callback
                )
        except (PermissionError, OSError) as e:
            if getattr(e, 'errno', None) == 13 or isinstance(e, PermissionError):
                from core.excel_com import lade_excel_daten_via_com
                result = lade_excel_daten_via_com(pfad, header_row, log_callback)
                if result is not None:
                    _, fallback_marken = result
            else:
                raise
        finally:
            if wb is not None:
                wb.close()

        return datensaetze, fallback_marken
    except Exception as e:
        raise Exception(f"FEHLER beim Lesen der Excel-Datei {pfad}: {e}") from e


def ersetze_platzhalter_mit_docxtpl(doc_path, context, svg_png_map, log_callback=None):
    """
    Ersetzt Platzhalter in einem Word-Dokument (.docx) mit den Daten aus dem Kontext.
    Nutzt docxtpl, was eine Jinja2-ähnliche Syntax erlaubt.

    Verarbeitungsreihenfolge (nach User-Wunsch für bessere Nachvollziehbarkeit):
    1. Bild-Platzhalter (_img)
    2. QR-Code-Platzhalter (_qr, _link)
    3. Alle anderen Text-Platzhalter

    Spezielle Logik:
    - Bilder: Platzhalter mit Suffix '_img' werden durch Bilder ersetzt. Die Bildgröße
      kann über einen zusätzlichen Platzhalter mit Suffix '_img_size' gesteuert werden.
    - QR-Codes: Platzhalter mit Suffix '_qr' oder '_link' werden in QR-Codes umgewandelt.
      Auch hier kann die Größe mit einem '_qr_size' oder '_link_size' Suffix angepasst werden.
    - SVG-Unterstützung: Verwendet die vorab konvertierten PNGs aus der svg_png_map.

    Args:
        doc_path (str): Pfad zur Word-Vorlagendatei.
        context (dict): Dictionary mit den Daten zum Ersetzen.
        svg_png_map (dict): Mapping von SVG-Pfaden zu temporären PNG-Pfaden.
        log_callback (callable, optional): Callback für Log-Nachrichten.

    Returns:
        DocxTemplate: Das bearbeitete Dokumentenobjekt.
    """
    doc = DocxTemplate(doc_path)
    try:
        undeclared_variables = doc.get_undeclared_template_variables()
    except Exception:
        undeclared_variables = set(context.keys())

    render_context = context.copy()
    bilder_ordner = context.get('_bilder_ordner')

    # Platzhalter nach Typ sortieren, um eine definierte Verarbeitungsreihenfolge zu gewährleisten.
    image_keys = sorted([k for k in undeclared_variables if k.endswith('_img')])
    qr_keys = sorted([k for k in undeclared_variables if k.endswith(('_qr', '_link'))])
    text_keys = sorted([k for k in undeclared_variables if k not in image_keys and k not in qr_keys])

    # Gewünschte Verarbeitungsreihenfolge: Bilder -> QR-Codes -> Text.
    processing_order = image_keys + qr_keys + text_keys

    for key in processing_order:
        value = context.get(key)
        if not isinstance(value, str) or not value.strip():
            continue

        # 1. Bild-Platzhalter verarbeiten
        if key in image_keys and bilder_ordner:
            width_cm = 15
            size_val = context.get(key + '_size', context.get(key + '_Size'))
            if pd.notna(size_val):
                try:
                    width_cm = float(size_val)
                except (ValueError, TypeError):
                    pass
            bild_pfad = os.path.join(bilder_ordner, value)
            if os.path.exists(bild_pfad):
                try:
                    render_context[key] = InlineImage(doc, svg_png_map.get(bild_pfad, bild_pfad), width=Cm(width_cm))
                except Exception as e:
                    render_context[key] = "[Bild-Fehler]"
                    _log_handler(f"Bild {value} Fehler: {e}", "ERROR", log_callback)
            else:
                render_context[key] = "[Bild nicht gefunden]"
                _log_handler(f"Bild {value} nicht gefunden", "WARN", log_callback)

        # 2. QR-Code-Platzhalter verarbeiten
        elif key in qr_keys:
            width_cm = 4
            size_val = context.get(key + '_size', context.get(key + '_Size'))
            if pd.notna(size_val):
                try:
                    width_cm = float(size_val)
                except (ValueError, TypeError):
                    pass
            try:
                qr_img = qrcode.make(value)
                img_bytes = io.BytesIO()
                qr_img.save(img_bytes, format='PNG')
                img_bytes.seek(0)
                render_context[key] = InlineImage(doc, img_bytes, width=Cm(width_cm))
            except Exception as e:
                render_context[key] = "[QR-Fehler]"
                _log_handler(f"QR-Code für '{value}' Fehler: {e}", "ERROR", log_callback)

    # 3. Alle Platzhalter (inkl. der vorbereiteten Bilder/QRs) von docxtpl rendern lassen.
    doc.render(render_context, autoescape=True)
    return doc


def _speichere_dokument(doc, vorlage_pfad, eintrag, is_anlage, haupt_export_ordner, categories, log_callback, is_dark_mode=False):
    """
    Speichert ein generiertes Word-Dokument im richtigen Ordner und mit dem korrekten Namen.

    - Kategorisierung: Der Dateiname der Vorlage wird geprüft. Beginnt er mit einem
      der in der GUI definierten Präfixe (z.B. 'B_'), wird das Dokument in den
      zugehörigen Kategorie-Ordner (z.B. 'Beschilderung') gespeichert.
    - Benennung:
        - Anlagenspezifische Dokumente: '{seriennummer}_{vorlagenname}.docx'
        - Allgemeine Dokumente: '{vorlagenname}.docx'
    """
    vorlage_name = os.path.basename(vorlage_pfad)
    cleaned_name = vorlage_name
    export_kategorie_ordner = haupt_export_ordner
    for prefix, cat_name in categories.items():
        if vorlage_name.lower().startswith(prefix.lower()):
            cleaned_name = vorlage_name[len(prefix):]
            export_kategorie_ordner = os.path.join(haupt_export_ordner, cat_name)
            break
    os.makedirs(export_kategorie_ordner, exist_ok=True)
    seriennummer = str(eintrag.get('anlage_seriennummer', 'Unbekannt')).strip()
    ziel_name = f"{seriennummer}_{cleaned_name}" if is_anlage else cleaned_name

    # Speichere das Dokument
    ziel_pfad = os.path.join(export_kategorie_ordner, ziel_name)
    doc.save(ziel_pfad)

    if is_anlage:
        _log_handler(
            f"-> Speichere anlagenspez. Dokument: '{ziel_name}' für Anlage: '{seriennummer}' in Ordner: '{os.path.basename(export_kategorie_ordner)}'",
            "SUCCESS",
            log_callback,
            is_dark_mode
        )
    else:
        _log_handler(
            f"-> Speichere allgemeines Dokument: '{ziel_name}' in Ordner: '{os.path.basename(export_kategorie_ordner)}'",
            "SUCCESS",
            log_callback,
            is_dark_mode
        )


def _convert_docx_to_pdf_in_folder(ordner, log_callback):
    """Konvertiert alle .docx-Dateien im Ordner (rekursiv) in PDFs. Nutzt docx2pdf (Windows/Word)."""
    try:
        from docx2pdf import convert
    except ImportError:
        log_callback("PDF-Export: Modul 'docx2pdf' nicht installiert. Bitte 'pip install docx2pdf' ausführen (Windows).", "WARN")
        return
    count = 0
    for root, _, files in os.walk(ordner):
        for name in files:
            if not name.lower().endswith('.docx') or name.startswith('~'):
                continue
            docx_path = os.path.join(root, name)
            try:
                convert(docx_path)
                count += 1
                log_callback(f"  -> PDF erstellt: {name}", "SUCCESS")
            except Exception as e:
                log_callback(f"  -> PDF fehlgeschlagen für '{name}': {e}", "ERROR")
    if count:
        log_callback(f"PDF-Erstellung abgeschlossen: {count} Datei(en).", "SUCCESS")
    else:
        log_callback("Keine Word-Dokumente zur PDF-Konvertierung gefunden.", "INFO")


def verarbeite_vorlagen_trockenlauf(vorlagen_ordner, excel_path, log_callback, header_row, bilder_ordner=None):
    """
    Führt eine "Trockenlauf"-Validierung durch, ohne Dokumente zu erstellen.
    Prüft, ob alle referenzierten Bilder und Platzhalter vorhanden sind.
    """
    def _log(msg, level="INFO"):
        _log_handler(msg, level, log_callback)

    _log("=" * 60 + "\n", "SEP")
    _log("Starte Trockenlauf (Konfigurationsprüfung)...", "INFO")
    _log("=" * 60 + "\n", "SEP")

    is_ok = True

    try:
        # 1. Excel-Daten laden
        datensaetze, fallback_marken = lade_excel_daten(excel_path, header_row, log_callback)
        if not datensaetze:
            _log("Keine Datensätze in Excel gefunden. Abbruch.", "WARN")
            return False

        excel_columns = set(datensaetze[0].keys()) | set(fallback_marken.keys())
        _log(f"Gefundene Spalten/Marken in Excel: {len(excel_columns)}", "INFO")
        _log(f"Gefundene Fallback-Textmarken (Blatt 2): {sorted(list(fallback_marken.keys()))}", "INFO")

        # 2. Vorlagen sammeln
        all_templates = []
        for root, _, files in os.walk(vorlagen_ordner):
            for file in files:
                if file.endswith('.docx') and not file.startswith('~'):
                    all_templates.append(os.path.join(root, file))
        _log(f"Gefundene Word-Vorlagen: {len(all_templates)}", "INFO")

        # 3. Vorlagen und Bilder validieren
        _log("\n--- Validierung der Vorlagen ---\n", "SEP")
        all_template_vars = set()
        # 'datetime_utc' wird automatisch hinzugefügt und ist daher kein Fehler, wenn es in Excel fehlt.
        valid_placeholders = excel_columns | {'datetime_utc'}

        for template_path in all_templates:
            try:
                doc = DocxTemplate(template_path)
                template_vars = doc.get_undeclared_template_variables()
                all_template_vars.update(template_vars)

                missing_vars = template_vars - valid_placeholders
                if missing_vars:
                    is_ok = False
                    _log(
                        f"Vorlage '{os.path.basename(template_path)}': Fehlende Platzhalter in Excel: {missing_vars}",
                        "ERROR"
                    )
                else:
                    _log(f"Vorlage '{os.path.basename(template_path)}': OK", "SUCCESS")

                # NEU: Für _img, _qr, _link Platzhalter prüfe nur in Excel/Fallback auf _size
                img_qr_keys = {k for k in template_vars if k.endswith(('_img', '_qr', '_link'))}
                if img_qr_keys:
                    _log(f"Platzhalter-Details für '{os.path.basename(template_path)}':", "INFO")
                for key in img_qr_keys:
                    size_key_lower = f"{key}_size"
                    size_key_camel = f"{key}_Size"
                    # Wert und Größe suchen (Excel/Fallback)
                    value = None
                    size = None
                    # Suche Wert in allen Datensätzen und Fallback
                    for eintrag in datensaetze + [fallback_marken]:
                        if key in eintrag and eintrag[key]:
                            value = eintrag[key]
                        if size_key_lower in eintrag and eintrag[size_key_lower]:
                            size = eintrag[size_key_lower]
                        elif size_key_camel in eintrag and eintrag[size_key_camel]:
                            size = eintrag[size_key_camel]
                    # Standardgrößen
                    if key.endswith('_img'):
                        default_size = 15
                        typ = 'Bild'
                    else:
                        default_size = 4
                        typ = 'QR/Link'
                    used_size = size if size is not None else default_size
                    value_str = value if value is not None else '[kein Wert in Excel/Fallback]'
                    _log(f"  - {typ}: '{key}' → Wert: {value_str}, Größe: {used_size} cm", "INFO")
                    if size is None:
                        _log(
                            f"Hinweis: Für Platzhalter '{key}' wird die Standardgröße verwendet, da kein '{size_key_lower}' oder '{size_key_camel}' in Excel/Fallback gefunden wurde.",
                            "INFO"
                        )

            except Exception as e:
                is_ok = False
                _log(
                    f"Vorlage '{os.path.basename(template_path)}': Konnte nicht gelesen werden. Fehler: {e}",
                    "FATAL"
                )

        _log("\n--- Validierung der Bilder (aus Excel-Daten) ---\n", "SEP")
        if bilder_ordner:
            image_references = set()
            for eintrag in datensaetze + [fallback_marken]:
                for key, value in eintrag.items():
                    if isinstance(value, str) and (value.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.svg'))):
                        image_references.add(value)

            _log(f"Gefundene Bild-Referenzen in Excel: {len(image_references)}", "INFO")
            for img_name in image_references:
                img_path = os.path.join(bilder_ordner, img_name)
                if not os.path.exists(img_path):
                    is_ok = False
                    _log(f"Bild '<strong>{img_name}</strong>' nicht gefunden", "WARN")
            if is_ok and image_references:
                _log("Alle referenzierten Bilder wurden gefunden.", "SUCCESS")

        else:
            _log("Kein Bilder-Ordner angegeben, Prüfung übersprungen.", "WARN")

        _log("\n" + "=" * 60, "SEP")
        if is_ok:
            _log("Trockenlauf erfolgreich! Konfiguration scheint gültig.", "SUCCESS")
        else:
            _log("Trockenlauf mit Fehlern beendet. Bitte prüfen Sie die obigen Meldungen.", "ERROR")

        return is_ok

    except Exception as e:
        _log_handler(f"FATALER FEHLER im Trockenlauf: {e}\n{traceback.format_exc()}", "FATAL", log_callback)
        return False


def verarbeite_vorlagen(
    vorlagen_ordner,
    export_ordner,
    excel_path,
    log_callback,
    progress_callback,
    file_callback,
    header_row,
    svg_scale,
    png_compression,
    categories,
    datetime_utc_format,
    bilder_ordner=None,
    worker_thread=None,
    dry_run=False,
    is_dark_mode=False,
    projekt_name_override=None,
    export_as_pdf=False
):
    """
    Die zentrale Orchestrierungsfunktion, die den gesamten Prozess steuert.

    Ablauf:
    1.  Lädt die Daten aus der Excel-Datei (Hauptdaten und Fallback-Marken).
    2.  Sucht nach allen SVG-Bildern, die in den Daten referenziert werden.
    3.  Konvertiert alle gefundenen, einzigartigen SVGs in PNGs und speichert sie
        in einem temporären Ordner, um die Performance zu verbessern (jedes SVG
        wird nur einmal konvertiert).
    4.  Durchsucht den Vorlagen-Ordner und teilt die gefundenen .docx-Dateien in
        'anlagen'-spezifische und 'allgemein'e Vorlagen auf.
    5.  Erstellt einen eindeutigen Export-Ordner für den aktuellen Durchlauf,
        benannt mit Zeitstempel und Projektname.
    6.  Verarbeitet die 'Anlagen'-Vorlagen: Iteriert über jeden Datensatz aus der
        Excel-Datei und füllt für jeden Datensatz jede Vorlage aus.
    7.  Verarbeitet die 'Allgemein'-Vorlagen: Füllt diese Vorlagen einmal aus,
        wobei der erste Datensatz und die Fallback-Marken als Datenquelle dienen.
    8.  Verwendet bei der Erstellung eine Fallback-Logik: Wenn ein Wert in einem
        Haupt-Datensatz fehlt, wird der entsprechende Wert aus den Fallback-Marken
        (zweites Excel-Blatt) verwendet.
    9.  Löscht nach Abschluss das temporäre Verzeichnis mit den PNGs.
    """
    if dry_run:
        return verarbeite_vorlagen_trockenlauf(vorlagen_ordner, excel_path, log_callback, header_row, bilder_ordner)

    def _log(msg, level="INFO"):
        _log_handler(msg, level, log_callback)

    _log("=" * 60 + "\n", "SEP")
    _log("Starte Dokumentenerstellung...", "INFO")
    _log("=" * 60 + "\n", "SEP")
    temp_dir = None
    try:
        datensaetze, fallback_marken = lade_excel_daten(excel_path, header_row, log_callback)
        if not datensaetze:
            _log("Keine Datensätze gefunden. Abbruch.", "WARN")
            return

        temp_dir = tempfile.mkdtemp(prefix="svg_cache_")
        unique_svg_paths = set()
        if bilder_ordner:
            for eintrag in datensaetze + [fallback_marken]:
                for key, value in eintrag.items():
                    if isinstance(value, str) and value.lower().endswith('.svg'):
                        full_path = os.path.join(bilder_ordner, value)
                        if os.path.exists(full_path):
                            unique_svg_paths.add(full_path)
        svg_png_map = {}
        if unique_svg_paths:
            _log(f"Konvertiere {len(unique_svg_paths)} einzigartige SVGs...", "INFO")
            for svg_path in unique_svg_paths:
                png_filename = os.path.splitext(os.path.basename(svg_path))[0] + ".png"
                temp_png_path = os.path.join(temp_dir, png_filename)
                if svg_to_png_file_pyside(svg_path, temp_png_path, _log, svg_scale, png_compression):
                    svg_png_map[svg_path] = temp_png_path
            _log("SVG-Konvertierung abgeschlossen.", "SUCCESS")

        anlagen_templates, allgemein_templates = [], []
        for root, _, files in os.walk(vorlagen_ordner):
            for file in files:
                if file.endswith('.docx') and not file.startswith('~'):
                    path_parts = os.path.relpath(root, vorlagen_ordner).lower().split(os.sep)
                    if 'anlagen' in path_parts:
                        anlagen_templates.append(os.path.join(root, file))
                    elif 'allgemein' in path_parts:
                        allgemein_templates.append(os.path.join(root, file))

        if projekt_name_override is not None and str(projekt_name_override).strip():
            raw = str(projekt_name_override).strip()
            projekt_name = "".join(c if c not in r'\/:*?"<>|' else "_" for c in raw) or "Unbenanntes_Projekt"
        else:
            projekt_name = str(
                datensaetze[0].get('projekt_name', fallback_marken.get('projekt_name', 'Unbenanntes_Projekt'))
            ).strip()
        haupt_export_ordner = os.path.join(
            export_ordner,
            datetime.now().strftime('%Y-%m-%d_%H%M%S') + "_" + projekt_name
        )

        total_docs = (len(datensaetze) * len(anlagen_templates)) + len(allgemein_templates)
        current_doc = 0

        def process_and_save(template, data, is_anlage):
            nonlocal current_doc
            current_doc += 1
            if worker_thread and worker_thread.isInterruptionRequested():
                return

            progress_callback(current_doc, total_docs)
            file_callback(os.path.basename(template))
            try:
                context = data.copy()
                for key, fallback_value in fallback_marken.items():
                    primary_value = context.get(key)
                    is_empty = pd.isna(primary_value) or (isinstance(primary_value, str) and not primary_value.strip())
                    if is_empty:
                        context[key] = fallback_value

                # datetime_utc dynamisch hinzufügen/überschreiben
                try:
                    # Nur aktuelles Datum im Format JJJJ-MM-DD
                    context['datetime_utc'] = datetime.now().strftime('%Y-%m-%d')
                except Exception as e:
                    context['datetime_utc'] = f"[Format-Fehler: {e}]"
                    _log(f"Ungültiges Zeitstempel-Format: '{datetime_utc_format}'. Fehler: {e}", "WARN")

                context['_bilder_ordner'] = bilder_ordner

                doc = ersetze_platzhalter_mit_docxtpl(template, context, svg_png_map, _log)
                _speichere_dokument(
                    doc,
                    template,
                    data,
                    is_anlage,
                    haupt_export_ordner,
                    categories,
                    _log,
                    is_dark_mode
                )
            except TemplateSyntaxError as e:
                _log(
                    f"FEHLER in Vorlage '{os.path.basename(template)}': Ungültige Syntax. Bitte prüfen Sie die Platzhalter.",
                    "FATAL"
                )
                _log(f"Details: {e}", "ERROR")
            except Exception as e:
                _log(f"FEHLER bei '{os.path.basename(template)}': {e}", "ERROR")

        _log("\n--- Verarbeitung 'Anlagen' ---\n", "SEP")
        for eintrag in datensaetze:
            if worker_thread and worker_thread.isInterruptionRequested():
                break
            for vorlage_pfad in anlagen_templates:
                process_and_save(vorlage_pfad, eintrag, True)

        if worker_thread and not worker_thread.isInterruptionRequested():
            _log("\n--- Verarbeitung 'Allgemein' ---\n", "SEP")
            if allgemein_templates:
                for vorlage_pfad in allgemein_templates:
                    process_and_save(vorlage_pfad, datensaetze[0], False)

        if worker_thread and worker_thread.isInterruptionRequested():
            _log("\n" + "=" * 60, "SEP")
            _log("Vorgang vom Benutzer abgebrochen.", "WARN")
        else:
            _log("\n" + "=" * 60, "SEP")
            _log("Alle Aufgaben abgeschlossen.", "SUCCESS")
            if export_as_pdf and haupt_export_ordner and os.path.isdir(haupt_export_ordner):
                _log("\n--- PDF-Erstellung ---\n", "SEP")
                _convert_docx_to_pdf_in_folder(haupt_export_ordner, _log)

        return haupt_export_ordner
    except Exception as e:
        _log_handler(f"FATALER FEHLER: {e}\n{traceback.format_exc()}", "FATAL", log_callback)
        return None
    finally:
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
                _log_handler("Temporäres Cache-Verzeichnis gelöscht.", "INFO", log_callback)
            except Exception as e:
                _log_handler(f"Cache konnte nicht gelöscht werden: {e}", "ERROR", log_callback)
