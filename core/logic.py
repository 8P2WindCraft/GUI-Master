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
from core.rules import (
    count_anlagen_emits_with_rules,
    evaluate_anlage_template_decision,
)
from core.utils import svg_to_png_file_pyside


def _norm_template_path(p):
    """Absoluter, plattformnormalisierter Pfad für Vergleiche (Windows: normcase)."""
    return os.path.normcase(os.path.normpath(os.path.abspath(p)))


def _format_rule_decision_log(decision: dict, eintrag: dict, template_path: str) -> str:
    seriennummer = str(eintrag.get('anlage_seriennummer', 'Unbekannt')).strip()
    if not decision.get("has_rule"):
        return (
            f"[Regel] Anlage '{seriennummer}' | Vorlage '{os.path.basename(template_path)}' | "
            f"keine aktive Regel ({decision.get('reason')})"
        )
    rule_name = decision.get("rule_name") or "(ohne Name)"
    configured_branch = (decision.get("configured_branch") or "-").upper()
    applied_branch = (decision.get("applied_branch") or "-").upper()
    emit_label = "ERZEUGEN" if decision.get("emit") else "ÜBERSPRINGEN"
    return (
        f"[Regel] Anlage '{seriennummer}' | Vorlage '{os.path.basename(template_path)}' | "
        f"Regel '{rule_name}' | Vorlage-Zweig={configured_branch} | "
        f"Entscheid={applied_branch} | Ergebnis={emit_label} | Grund: {decision.get('reason')}"
    )


def sammle_vorlagen_pfade(vorlagen_ordner):
    """
    Sammelt .docx-Vorlagen unter anlagen/ bzw. allgemein/ (gleiche Regeln wie die Generierung).

    Returns:
        (anlagen_templates, allgemein_templates): Listen absoluter Pfade.
    """
    anlagen_templates, allgemein_templates = [], []
    if not vorlagen_ordner or not os.path.isdir(vorlagen_ordner):
        return anlagen_templates, allgemein_templates
    for root, _, files in os.walk(vorlagen_ordner):
        for file in files:
            if file.endswith('.docx') and not file.startswith('~'):
                path_parts = os.path.relpath(root, vorlagen_ordner).lower().split(os.sep)
                if 'anlagen' in path_parts:
                    anlagen_templates.append(os.path.join(root, file))
                elif 'allgemein' in path_parts:
                    allgemein_templates.append(os.path.join(root, file))
    return anlagen_templates, allgemein_templates


def liste_textmarken_aus_docx(doc_path: str, log_callback=None):
    """
    Liefert sortierte Text-Platzhalter aus einer .docx-Vorlage (ohne _img / _qr / _link),
    wie bei der Render-Reihenfolge in ersetze_platzhalter_mit_docxtpl.
    """
    if not doc_path or not os.path.isfile(doc_path):
        return []
    try:
        doc = DocxTemplate(doc_path)
        variables = doc.get_undeclared_template_variables()
    except Exception as e:
        _log_handler(f"Textmarken aus Vorlage: {e}", "WARN", log_callback)
        return []
    image_keys = {k for k in variables if k.endswith("_img")}
    qr_keys = {k for k in variables if k.endswith(("_qr", "_link"))}
    text_keys = sorted(k for k in variables if k not in image_keys and k not in qr_keys)
    return text_keys


def filter_vorlagen_nach_auswahl(anlagen_templates, allgemein_templates, selected_template_paths):
    """
    Schränkt die Listen auf gewählte Pfade ein.

    - selected_template_paths None: keine Einschränkung (alle übergebenen).
    - sonst: Schnittmenge nach normalisiertem Pfad.
    """
    if selected_template_paths is None:
        return list(anlagen_templates), list(allgemein_templates)
    sel = {_norm_template_path(p) for p in selected_template_paths}
    anlagen_f = [p for p in anlagen_templates if _norm_template_path(p) in sel]
    allg_f = [p for p in allgemein_templates if _norm_template_path(p) in sel]
    return anlagen_f, allg_f


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


def _lageplan_ziel_pfad(vorlage_pfad, eintrag, haupt_export_ordner, categories):
    """Gleiche Ordner- und Dateinamen-Logik wie _speichere_dokument (nur Anlagen)."""
    vorlage_name = os.path.basename(vorlage_pfad)
    cleaned_name = vorlage_name
    export_kategorie_ordner = haupt_export_ordner
    for prefix, cat_name in categories.items():
        if vorlage_name.lower().startswith(prefix.lower()):
            cleaned_name = vorlage_name[len(prefix):]
            export_kategorie_ordner = os.path.join(haupt_export_ordner, cat_name)
            break
    seriennummer = str(eintrag.get('anlage_seriennummer', 'Unbekannt')).strip()
    ziel_name = f"{seriennummer}_{cleaned_name}"
    return export_kategorie_ordner, ziel_name


def _lageplan_quell_unterordner_liste(categories):
    """Suchreihenfolge: Kategorie zu rl_, dann Lageplan, Plan."""
    seen = set()
    out = []
    for pref, name in categories.items():
        if str(pref).lower().startswith('rl') and name:
            if name not in seen:
                out.append(name)
                seen.add(name)
    for fb in ('Lageplan', 'Plan'):
        if fb not in seen:
            out.append(fb)
            seen.add(fb)
    return out


def find_lageplan_source_file(previous_export_root, categories, vorlage_pfad, eintrag):
    """Pfad zur .docx im letzten Export, falls vorhanden; sonst None."""
    if not previous_export_root or not os.path.isdir(previous_export_root):
        return None
    vorlage_name = os.path.basename(vorlage_pfad)
    if not vorlage_name.lower().startswith('rl_'):
        return None
    _, ziel_name = _lageplan_ziel_pfad(vorlage_pfad, eintrag, previous_export_root, categories)
    for sub in _lageplan_quell_unterordner_liste(categories):
        candidate = os.path.normpath(os.path.join(previous_export_root, sub, ziel_name))
        if os.path.isfile(candidate):
            return candidate
    return None


def _normalize_windows_path_for_word_com(path):
    """Normalisiert Pfade für Word-COM (Backslashes, absolut, ohne URL-Encoding)."""
    normalized = os.path.normpath(os.path.abspath(path))
    return normalized.replace("/", "\\")


def _docx_to_pdf_word_com_single(docx_path):
    """
    Konvertiert eine .docx zu .pdf im gleichen Ordner (Word-COM, Windows).
    Rückgabe: (erfolg: bool, fehler oder None).
    """
    try:
        import pythoncom
        import pywintypes
        import win32com.client
    except ImportError as e:
        return False, e

    wd_export_format_pdf = 17
    wd_export_optimize_for_print = 0
    wd_export_all_document = 0
    wd_export_document_content = 0
    wd_export_create_no_bookmarks = 0
    transient_hresults = {-2147418111, -2147417848}

    docx_path = _normalize_windows_path_for_word_com(docx_path)
    pdf_path = _normalize_windows_path_for_word_com(os.path.splitext(docx_path)[0] + ".pdf")

    if not os.path.isfile(docx_path):
        return False, FileNotFoundError(f"Datei nicht gefunden: {docx_path}")

    last_error = None
    for attempt in range(1, 4):
        word = None
        doc = None
        try:
            pythoncom.CoInitialize()
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            doc = word.Documents.Open(docx_path, ReadOnly=True, AddToRecentFiles=False)
            doc.ExportAsFixedFormat(
                OutputFileName=pdf_path,
                ExportFormat=wd_export_format_pdf,
                OpenAfterExport=False,
                OptimizeFor=wd_export_optimize_for_print,
                Range=wd_export_all_document,
                From=1,
                To=1,
                Item=wd_export_document_content,
                IncludeDocProps=True,
                KeepIRM=True,
                CreateBookmarks=wd_export_create_no_bookmarks,
                DocStructureTags=True,
                BitmapMissingFonts=True,
                UseISO19005_1=False
            )
            return True, None
        except pywintypes.com_error as e:
            last_error = e
            hresult = e.args[0] if e.args else None
            if hresult in transient_hresults and attempt < 3:
                import time
                time.sleep(0.8 * attempt)
                continue
            return False, e
        except Exception as e:
            last_error = e
            return False, e
        finally:
            if doc is not None:
                try:
                    doc.Close(False)
                except Exception:
                    pass
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    return False, last_error


def konvertiere_ein_docx_zu_pdf(docx_path, log_callback=None):
    """Eine .docx-Datei per Word-COM in PDF umwandeln; loggt Ergebnis."""
    ok, err = _docx_to_pdf_word_com_single(docx_path)
    if isinstance(err, ImportError):
        _log_handler(
            "PDF-Export: pywin32 nicht verfügbar. Bitte in der aktiven Umgebung installieren: "
            "pip install pywin32 && pywin32_postinstall.py -install",
            "WARN",
            log_callback,
        )
        return False
    if ok:
        _log_handler(
            f"PDF erstellt: {os.path.basename(os.path.splitext(docx_path)[0] + '.pdf')}",
            "SUCCESS",
            log_callback,
        )
        return True
    _log_handler(f"PDF fehlgeschlagen für '{os.path.basename(docx_path)}': {err}", "ERROR", log_callback)
    return False


def _convert_docx_to_pdf_in_folder(ordner, log_callback, progress_callback=None, file_callback=None):
    """Konvertiert alle .docx-Dateien im Ordner (rekursiv) in PDFs über Word-COM (Windows)."""
    result = {
        "docx_count": 0,
        "success_count": 0,
        "failure_count": 0,
        "error": None,
    }
    try:
        import pywintypes  # noqa: F401 — Prüfung, ob pywin32 nutzbar ist
        import win32com.client  # noqa: F401
    except ImportError:
        log_callback(
            "PDF-Export: pywin32 nicht verfügbar. Bitte in der aktiven Umgebung installieren: "
            "pip install pywin32 && pywin32_postinstall.py -install",
            "WARN"
        )
        result["error"] = "missing_pywin32"
        return result

    docx_paths = []
    for root, _, files in os.walk(ordner):
        for name in files:
            if not name.lower().endswith('.docx') or name.startswith('~'):
                continue
            docx_paths.append(os.path.join(root, name))

    result["docx_count"] = len(docx_paths)
    if not docx_paths:
        if progress_callback:
            progress_callback(0, 0)
        log_callback("Keine Word-Dokumente zur PDF-Konvertierung gefunden.", "INFO")
        return result

    for index, docx_path in enumerate(docx_paths, start=1):
        name = os.path.basename(docx_path)
        if file_callback:
            file_callback(name)
        try:
            ok, error = _docx_to_pdf_word_com_single(docx_path)
            if ok:
                result["success_count"] += 1
                log_callback(f"  -> PDF erstellt: {name}", "SUCCESS")
            else:
                result["failure_count"] += 1
                log_callback(f"  -> PDF fehlgeschlagen für '{name}': {error}", "ERROR")
        except Exception as e:
            result["failure_count"] += 1
            log_callback(f"  -> PDF fehlgeschlagen für '{name}': {e}", "ERROR")
        if progress_callback:
            progress_callback(index, result["docx_count"])

    if result["success_count"]:
        log_callback(f"PDF-Erstellung abgeschlossen: {result['success_count']} Datei(en).", "SUCCESS")
    if result["failure_count"]:
        log_callback(
            f"PDF-Erstellung mit Fehlern: {result['failure_count']} von {result['docx_count']} Datei(en) fehlgeschlagen.",
            "WARN"
        )
    return result


def verarbeite_vorlagen_trockenlauf(
    vorlagen_ordner,
    excel_path,
    log_callback,
    header_row,
    bilder_ordner=None,
    selected_template_paths=None,
    rules_enabled=False,
    signage_rules=None,
    categories=None,
    reuse_lageplan_from_last_export=False,
    previous_export_root=None,
    export_as_pdf=False,
):
    """
    Führt eine "Trockenlauf"-Validierung durch, ohne Dokumente zu erstellen.
    Prüft, ob alle referenzierten Bilder und Platzhalter vorhanden sind.

    Nur Vorlagen unter anlagen/ und allgemein/ (wie bei der Generierung).
    Mit selected_template_paths wird auf eine Teilmenge eingeschränkt.
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

        # 2. Vorlagen sammeln (nur Anlagen/Allgemein, optional gefiltert)
        anlagen_raw, allg_raw = sammle_vorlagen_pfade(vorlagen_ordner)
        anlagen_t, allg_t = filter_vorlagen_nach_auswahl(anlagen_raw, allg_raw, selected_template_paths)
        all_templates = anlagen_t + allg_t
        _log(
            f"Gefundene Word-Vorlagen (Anlagen/Allgemein): {len(all_templates)} "
            f"(Anlagen: {len(anlagen_t)}, Allgemein: {len(allg_t)})",
            "INFO",
        )
        if export_as_pdf:
            _log(
                "PDF-Erzeugung ist aktiviert: Nach der DOCX-Erstellung würden die erzeugten Dokumente "
                "zusätzlich in PDF umgewandelt.",
                "INFO",
            )
        if not all_templates:
            _log("Keine Vorlagen für den Trockenlauf ausgewählt oder vorhanden. Abbruch.", "WARN")
            return False

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

        if rules_enabled and signage_rules:
            naive_anlagen = len(datensaetze) * len(anlagen_t)
            eff_anlagen = count_anlagen_emits_with_rules(
                datensaetze,
                anlagen_t,
                vorlagen_ordner,
                fallback_marken,
                bilder_ordner,
                True,
                signage_rules,
            )
            _log("\n--- Regeln (optionale Schilder) ---\n", "SEP")
            _log(
                f"Aktiv: {eff_anlagen} von {naive_anlagen} möglichen Anlagen-Dokumenten würden erzeugt "
                f"({naive_anlagen - eff_anlagen} übersprungen).",
                "INFO",
            )
            _log("\n--- Regel-Entscheidungen (Vorschau) ---\n", "SEP")
            for eintrag in datensaetze:
                for tp in anlagen_t:
                    decision = evaluate_anlage_template_decision(
                        vorlagen_ordner,
                        tp,
                        eintrag,
                        fallback_marken,
                        bilder_ordner,
                        True,
                        signage_rules,
                    )
                    if not decision.get("has_rule"):
                        continue
                    lvl = "INFO" if decision.get("emit") else "WARN"
                    _log(_format_rule_decision_log(decision, eintrag, tp), lvl)

        if reuse_lageplan_from_last_export:
            cats = categories or {}
            copy_n = 0
            gen_n = 0
            prev_ok = bool(previous_export_root and os.path.isdir(previous_export_root))
            for eintrag in datensaetze:
                for tp in anlagen_t:
                    decision = evaluate_anlage_template_decision(
                        vorlagen_ordner,
                        tp,
                        eintrag,
                        fallback_marken,
                        bilder_ordner,
                        rules_enabled,
                        signage_rules,
                    )
                    if not decision.get("emit"):
                        continue
                    if not os.path.basename(tp).lower().startswith('rl_'):
                        continue
                    if prev_ok and find_lageplan_source_file(previous_export_root, cats, tp, eintrag):
                        copy_n += 1
                    else:
                        gen_n += 1
            _log("\n--- Lageplan-Übernahme (Vorschau) ---\n", "SEP")
            if not prev_ok:
                _log(
                    "Kein gültiger letzter Export-Ordner: alle rl_-Lagepläne würden wie gewohnt neu erstellt.",
                    "INFO",
                )
            _log(
                f"rl_-Vorlagen: {copy_n} würden aus dem letzten Export kopiert, "
                f"{gen_n} neu aus Vorlage erzeugt.",
                "INFO",
            )

        _log("\n" + "=" * 60, "SEP")
        if is_ok:
            _log("Trockenlauf erfolgreich! Konfiguration scheint gültig.", "SUCCESS")
        else:
            _log("Trockenlauf mit Fehlern beendet. Bitte prüfen Sie die obigen Meldungen.", "ERROR")

        return is_ok

    except Exception as e:
        _log_handler(f"FATALER FEHLER im Trockenlauf: {e}\n{traceback.format_exc()}", "FATAL", log_callback)
        return False


def verarbeite_vorlagen_preview(
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
    is_dark_mode=False,
    projekt_name_override=None,
    export_as_pdf=False,
    selected_template_paths=None,
    rules_enabled=False,
    signage_rules=None,
    reuse_lageplan_from_last_export=False,
    previous_export_root=None,
    preview_template_abs=None,
):
    """
    Erzeugt eine Stichprobe wie im echten Lauf, aber nur mit der ersten Excel-Datenzeile
    für Anlagen-Vorlagen; Allgemein wie in Produktion (erster Datensatz + Fallback).
    """
    def _log(msg, level="INFO"):
        _log_handler(msg, level, log_callback)

    _log("=" * 60 + "\n", "SEP")
    _log("Starte Dokument-Vorschau (erste Excel-Datenzeile)...", "INFO")
    _log(
        "Hinweis: Vorschau ≠ rechtlicher Prüfstand; Layout kann bei anderen Daten abweichen.",
        "INFO",
    )
    _log("=" * 60 + "\n", "SEP")
    temp_dir = None
    try:
        datensaetze, fallback_marken = lade_excel_daten(excel_path, header_row, log_callback)
        if not datensaetze:
            _log("Keine Datensätze gefunden. Abbruch.", "WARN")
            return None

        eintrag_preview = datensaetze[0]
        sn = str(eintrag_preview.get("anlage_seriennummer", "")).strip()
        _log(f"Vorschau-Datensatz: erste Zeile (anlage_seriennummer: {sn!r}).", "INFO")

        anlagen_raw, allg_raw = sammle_vorlagen_pfade(vorlagen_ordner)
        anlagen_templates, allgemein_templates = filter_vorlagen_nach_auswahl(
            anlagen_raw, allg_raw, selected_template_paths
        )
        if preview_template_abs:
            target = _norm_template_path(preview_template_abs)
            anlagen_templates = [p for p in anlagen_templates if _norm_template_path(p) == target]
            allgemein_templates = [p for p in allgemein_templates if _norm_template_path(p) == target]
            if not anlagen_templates and not allgemein_templates:
                _log(
                    f"Vorschau: Vorlage nicht in der aktuellen Auswahl oder nicht unter anlagen/ bzw. allgemein/: "
                    f"{os.path.basename(preview_template_abs)}",
                    "WARN",
                )
                return None

        anlagen_emits = count_anlagen_emits_with_rules(
            datensaetze[:1],
            anlagen_templates,
            vorlagen_ordner,
            fallback_marken,
            bilder_ordner,
            rules_enabled,
            signage_rules,
        )
        generation_docs_total = anlagen_emits + len(allgemein_templates)
        if generation_docs_total == 0:
            _log(
                "Vorschau: Keine Dokumente zu erzeugen (0 erwartete Ausgaben nach Regeln / ohne Allgemein). "
                "Es werden trotzdem Entscheidungen protokolliert.",
                "WARN",
            )
        pdf_docs_total = generation_docs_total if export_as_pdf else 0
        overall_total_steps = generation_docs_total + pdf_docs_total
        if overall_total_steps == 0:
            overall_total_steps = 1

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

        if projekt_name_override is not None and str(projekt_name_override).strip():
            raw = str(projekt_name_override).strip()
            projekt_name = "".join(c if c not in r'\/:*?"<>|' else "_" for c in raw) or "Unbenanntes_Projekt"
        else:
            projekt_name = str(
                datensaetze[0].get('projekt_name', fallback_marken.get('projekt_name', 'Unbenanntes_Projekt'))
            ).strip()
        haupt_export_ordner = os.path.join(
            export_ordner,
            datetime.now().strftime('%Y-%m-%d_%H%M%S') + "_" + projekt_name + "_Vorschau",
        )

        current_doc = 0
        warned_lageplan_no_export = False

        def process_and_save(template, data, is_anlage):
            nonlocal current_doc
            current_doc += 1
            if worker_thread and worker_thread.isInterruptionRequested():
                return

            progress_callback(current_doc, overall_total_steps)
            file_callback(os.path.basename(template))
            try:
                context = data.copy()
                for key, fallback_value in fallback_marken.items():
                    primary_value = context.get(key)
                    is_empty = pd.isna(primary_value) or (isinstance(primary_value, str) and not primary_value.strip())
                    if is_empty:
                        context[key] = fallback_value

                try:
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
                    is_dark_mode,
                )
            except TemplateSyntaxError as e:
                _log(
                    f"FEHLER in Vorlage '{os.path.basename(template)}': Ungültige Syntax. Bitte prüfen Sie die Platzhalter.",
                    "FATAL",
                )
                _log(f"Details: {e}", "ERROR")
            except Exception as e:
                _log(f"FEHLER bei '{os.path.basename(template)}': {e}", "ERROR")

        _log("\n--- Vorschau 'Anlagen' (nur erste Zeile) ---\n", "SEP")
        for vorlage_pfad in anlagen_templates:
            if worker_thread and worker_thread.isInterruptionRequested():
                break
            decision = evaluate_anlage_template_decision(
                vorlagen_ordner,
                vorlage_pfad,
                eintrag_preview,
                fallback_marken,
                bilder_ordner,
                rules_enabled,
                signage_rules,
            )
            if decision.get("has_rule"):
                detail = _format_rule_decision_log(decision, eintrag_preview, vorlage_pfad)
                if decision.get("emit"):
                    _log(detail, "INFO")
                else:
                    _log("Vorschau übersprungen: " + detail, "WARN")
            elif not decision.get("emit"):
                _log(
                    f"Vorschau übersprungen für '{os.path.basename(vorlage_pfad)}' (keine Ausgabe).",
                    "WARN",
                )
            if not decision.get("emit"):
                continue
            vn = os.path.basename(vorlage_pfad)
            if reuse_lageplan_from_last_export and vn.lower().startswith('rl_'):
                if not previous_export_root or not os.path.isdir(previous_export_root):
                    if not warned_lageplan_no_export:
                        _log(
                            "Lageplan-Übernahme: Kein gültiger letzter Export-Ordner – "
                            "rl_-Dateien werden wie gewohnt neu erstellt.",
                            "INFO",
                        )
                        warned_lageplan_no_export = True
                    process_and_save(vorlage_pfad, eintrag_preview, True)
                    continue
                src = find_lageplan_source_file(
                    previous_export_root, categories, vorlage_pfad, eintrag_preview
                )
                export_kat, ziel_name = _lageplan_ziel_pfad(
                    vorlage_pfad, eintrag_preview, haupt_export_ordner, categories
                )
                if src:
                    try:
                        os.makedirs(export_kat, exist_ok=True)
                        shutil.copy2(src, os.path.join(export_kat, ziel_name))
                    except OSError as e:
                        _log(
                            f"Lageplan kopieren fehlgeschlagen ({e}), neu aus Vorlage.",
                            "WARN",
                        )
                        process_and_save(vorlage_pfad, eintrag_preview, True)
                        continue
                    current_doc += 1
                    if worker_thread and worker_thread.isInterruptionRequested():
                        break
                    progress_callback(current_doc, overall_total_steps)
                    file_callback(vn)
                    _log(
                        f"Lageplan übernommen (letzter Export): '{ziel_name}'",
                        "SUCCESS",
                    )
                    continue
                _log(
                    f"Keine vorherige Datei für '{ziel_name}' im letzten Export – "
                    "Erstellung wie gewohnt aus Vorlage.",
                    "INFO",
                )
                process_and_save(vorlage_pfad, eintrag_preview, True)
                continue

            process_and_save(vorlage_pfad, eintrag_preview, True)

        if worker_thread and not worker_thread.isInterruptionRequested():
            _log("\n--- Vorschau 'Allgemein' ---\n", "SEP")
            if allgemein_templates:
                for vorlage_pfad in allgemein_templates:
                    process_and_save(vorlage_pfad, datensaetze[0], False)

        docx_erzeugt = False
        if os.path.isdir(haupt_export_ordner):
            for root, _, files in os.walk(haupt_export_ordner):
                for name in files:
                    if name.lower().endswith('.docx') and not name.startswith('~'):
                        docx_erzeugt = True
                        break
                if docx_erzeugt:
                    break

        if worker_thread and worker_thread.isInterruptionRequested():
            _log("\n" + "=" * 60, "SEP")
            _log("Vorschau vom Benutzer abgebrochen.", "WARN")
            return None

        if not docx_erzeugt:
            _log(
                "Vorschau: Keine DOCX-Dateien erzeugt (alle Anlagen-Vorlagen übersprungen oder fehlerhaft).",
                "WARN",
            )
            return None

        _log("\n" + "=" * 60, "SEP")
        _log("Vorschau-Dokumente (DOCX) erstellt.", "SUCCESS")
        if export_as_pdf and haupt_export_ordner and os.path.isdir(haupt_export_ordner):
            _log("\n--- PDF-Erstellung (Vorschau) ---\n", "SEP")

            def _pdf_progress(pdf_current, pdf_total):
                progress_callback(generation_docs_total + pdf_current, overall_total_steps)

            def _pdf_file(filename):
                file_callback(f"PDF: {filename}")

            pdf_result = _convert_docx_to_pdf_in_folder(
                haupt_export_ordner,
                _log,
                progress_callback=_pdf_progress,
                file_callback=_pdf_file,
            )
            if pdf_result.get("error") == "missing_pywin32":
                _log(
                    "PDF-Vorschau nicht möglich: Microsoft Word und pywin32 werden für den Export benötigt.",
                    "WARN",
                )
            elif pdf_result.get("failure_count"):
                _log(
                    "Vorschau-PDF: Mindestens eine Konvertierung ist fehlgeschlagen (Word installiert / Datei nicht gesperrt?).",
                    "WARN",
                )

        return haupt_export_ordner
    except Exception as e:
        _log_handler(f"FATALER FEHLER (Vorschau): {e}\n{traceback.format_exc()}", "FATAL", log_callback)
        return None
    finally:
        if temp_dir and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
                _log_handler("Temporäres Cache-Verzeichnis gelöscht.", "INFO", log_callback)
            except Exception as e:
                _log_handler(f"Cache konnte nicht gelöscht werden: {e}", "ERROR", log_callback)


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
    export_as_pdf=False,
    selected_template_paths=None,
    rules_enabled=False,
    signage_rules=None,
    reuse_lageplan_from_last_export=False,
    previous_export_root=None,
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
        return verarbeite_vorlagen_trockenlauf(
            vorlagen_ordner,
            excel_path,
            log_callback,
            header_row,
            bilder_ordner,
            selected_template_paths=selected_template_paths,
            rules_enabled=rules_enabled,
            signage_rules=signage_rules,
            categories=categories,
            reuse_lageplan_from_last_export=reuse_lageplan_from_last_export,
            previous_export_root=previous_export_root,
            export_as_pdf=export_as_pdf,
        )

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

        anlagen_raw, allg_raw = sammle_vorlagen_pfade(vorlagen_ordner)
        anlagen_templates, allgemein_templates = filter_vorlagen_nach_auswahl(
            anlagen_raw, allg_raw, selected_template_paths
        )
        anlagen_emits = count_anlagen_emits_with_rules(
            datensaetze,
            anlagen_templates,
            vorlagen_ordner,
            fallback_marken,
            bilder_ordner,
            rules_enabled,
            signage_rules,
        )
        generation_docs_total = anlagen_emits + len(allgemein_templates)
        if generation_docs_total == 0:
            _log(
                "Keine Vorlagen zum Erzeugen oder keine Auswahl: Abbruch (0 Dokumente).",
                "WARN",
            )
            return None
        pdf_docs_total = generation_docs_total if export_as_pdf else 0
        overall_total_steps = generation_docs_total + pdf_docs_total

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

        current_doc = 0

        def process_and_save(template, data, is_anlage):
            nonlocal current_doc
            current_doc += 1
            if worker_thread and worker_thread.isInterruptionRequested():
                return

            progress_callback(current_doc, overall_total_steps)
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

        warned_lageplan_no_export = False

        _log("\n--- Verarbeitung 'Anlagen' ---\n", "SEP")
        for eintrag in datensaetze:
            if worker_thread and worker_thread.isInterruptionRequested():
                break
            for vorlage_pfad in anlagen_templates:
                decision = evaluate_anlage_template_decision(
                    vorlagen_ordner,
                    vorlage_pfad,
                    eintrag,
                    fallback_marken,
                    bilder_ordner,
                    rules_enabled,
                    signage_rules,
                )
                if decision.get("has_rule"):
                    lvl = "INFO" if decision.get("emit") else "WARN"
                    _log(_format_rule_decision_log(decision, eintrag, vorlage_pfad), lvl)
                if not decision.get("emit"):
                    continue
                vn = os.path.basename(vorlage_pfad)
                if reuse_lageplan_from_last_export and vn.lower().startswith('rl_'):
                    if not previous_export_root or not os.path.isdir(previous_export_root):
                        if not warned_lageplan_no_export:
                            _log(
                                "Lageplan-Übernahme: Kein gültiger letzter Export-Ordner – "
                                "rl_-Dateien werden wie gewohnt neu erstellt.",
                                "INFO",
                            )
                            warned_lageplan_no_export = True
                        process_and_save(vorlage_pfad, eintrag, True)
                        continue
                    src = find_lageplan_source_file(
                        previous_export_root, categories, vorlage_pfad, eintrag
                    )
                    export_kat, ziel_name = _lageplan_ziel_pfad(
                        vorlage_pfad, eintrag, haupt_export_ordner, categories
                    )
                    if src:
                        try:
                            os.makedirs(export_kat, exist_ok=True)
                            shutil.copy2(src, os.path.join(export_kat, ziel_name))
                        except OSError as e:
                            _log(
                                f"Lageplan kopieren fehlgeschlagen ({e}), neu aus Vorlage.",
                                "WARN",
                            )
                            process_and_save(vorlage_pfad, eintrag, True)
                            continue
                        current_doc += 1
                        if worker_thread and worker_thread.isInterruptionRequested():
                            break
                        progress_callback(current_doc, overall_total_steps)
                        file_callback(vn)
                        _log(
                            f"Lageplan übernommen (letzter Export): '{ziel_name}'",
                            "SUCCESS",
                        )
                        continue
                    _log(
                        f"Keine vorherige Datei für '{ziel_name}' im letzten Export – "
                        "Erstellung wie gewohnt aus Vorlage.",
                        "INFO",
                    )
                    process_and_save(vorlage_pfad, eintrag, True)
                    continue

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
                def _pdf_progress(pdf_current, pdf_total):
                    # PDF-Schritte direkt hinter die Dokumentenerzeugung hängen.
                    progress_callback(generation_docs_total + pdf_current, overall_total_steps)

                def _pdf_file(filename):
                    file_callback(f"PDF: {filename}")

                _convert_docx_to_pdf_in_folder(
                    haupt_export_ordner,
                    _log,
                    progress_callback=_pdf_progress,
                    file_callback=_pdf_file,
                )

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
