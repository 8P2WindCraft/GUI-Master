# -*- coding: utf-8 -*-
"""
Lese-Zugriff auf Excel-Dateien über Windows COM, wenn die Datei in Excel geöffnet ist.
Nur unter Windows und nur wenn Excel die Datei bereits geöffnet hat.
"""
import os
import sys

import pandas as pd


def _excel_app():
    """Verbinde mit der laufenden Excel-Instanz (oder starte keine neue)."""
    if sys.platform != "win32":
        return None
    try:
        import win32com.client
        return win32com.client.GetActiveObject("Excel.Application")
    except Exception:
        return None


def _find_open_workbook(app, pfad):
    """Findet die geöffnete Arbeitsmappe mit dem angegebenen Pfad."""
    if app is None or not pfad:
        return None
    pfad_norm = os.path.normpath(os.path.abspath(pfad)).lower()
    try:
        for wb in app.Workbooks:
            try:
                full = wb.FullName
                if full:
                    full_norm = os.path.normpath(os.path.abspath(full)).lower()
                    if full_norm == pfad_norm:
                        return wb
            except Exception:
                continue
    except Exception:
        pass
    return None


def lade_excel_daten_via_com(pfad, header_row=3, log_callback=None):
    """
    Liest Excel-Daten über COM aus der bereits in Excel geöffneten Arbeitsmappe.
    Gibt dasselbe Format zurück wie core.logic.lade_excel_daten:
    (datensaetze, fallback_marken).

    Returns:
        tuple: (datensaetze, fallback_marken) oder None bei Fehler.
    """
    def _log(msg, level="INFO"):
        if log_callback:
            try:
                from core.logging import _log_handler
                _log_handler(msg, level, log_callback)
            except Exception:
                pass

    if sys.platform != "win32":
        return None
    app = _excel_app()
    wb = _find_open_workbook(app, pfad)
    if wb is None:
        _log("Excel-Datei nicht in geöffneter Excel-Instanz gefunden.", "WARN")
        return None
    try:
        # Blatt 1: Header in header_row, Daten darunter
        sh1 = wb.Sheets(1)
        used = sh1.UsedRange
        if used is None:
            return [], {}
        max_row = used.Rows.Count
        max_col = used.Columns.Count
        header_idx = header_row  # 1-basiert in Excel
        if max_row < header_idx or max_col < 1:
            return [], {}

        # Spaltenüberschriften aus header_row
        headers = []
        for c in range(1, max_col + 1):
            v = sh1.Cells(header_idx, c).Value
            headers.append(str(v).strip() if v is not None else "")

        # Datensätze: nur Zeilen mit anlage_seriennummer
        datensaetze = []
        for r in range(header_idx + 1, max_row + 1):
            row_dict = {}
            for c, col_name in enumerate(headers):
                if not col_name:
                    continue
                v = sh1.Cells(r, c + 1).Value
                if isinstance(v, str):
                    row_dict[col_name] = v.strip()
                elif v is not None:
                    row_dict[col_name] = v
                else:
                    row_dict[col_name] = ""
            if "anlage_seriennummer" in row_dict and row_dict.get("anlage_seriennummer"):
                datensaetze.append(row_dict)

        _log(f"Via COM: {len(datensaetze)} Datensätze aus geöffneter Arbeitsmappe geladen.", "SUCCESS")

        # Blatt 2: Fallback-Marken (Spalte B = Schlüssel, Spalte C = Wert)
        fallback_marken = {}
        if wb.Sheets.Count >= 2:
            sh2 = wb.Sheets(2)
            last_row = sh2.UsedRange.Rows.Count if sh2.UsedRange else 0
            for r in range(2, last_row + 1):
                key_cell = sh2.Cells(r, 2).Value
                if key_cell is None:
                    continue
                key = str(key_cell).strip()
                if not key:
                    continue
                val_cell = sh2.Cells(r, 3).Value
                value = str(val_cell).strip() if val_cell is not None else ""
                fallback_marken[key] = value
            _log(f"Via COM: {len(fallback_marken)} Fallback-Textmarken aus Blatt 2 geladen.", "SUCCESS")

        return datensaetze, fallback_marken
    except Exception as e:
        _log(f"COM-Lesen fehlgeschlagen: {e}", "WARN")
        return None


def read_excel_sheet1_via_com(pfad, header_row=3):
    """
    Liest das erste Blatt vollständig als DataFrame (für GUI-Vorschau, alle Zeilen).
    Returns:
        pandas.DataFrame oder None bei Fehler.
    """
    if sys.platform != "win32":
        return None
    app = _excel_app()
    wb = _find_open_workbook(app, pfad)
    if wb is None:
        return None
    try:
        sh1 = wb.Sheets(1)
        used = sh1.UsedRange
        if used is None:
            return pd.DataFrame()
        max_row = used.Rows.Count
        max_col = used.Columns.Count
        hi = header_row
        if max_row < hi or max_col < 1:
            return pd.DataFrame()
        headers = []
        for c in range(1, max_col + 1):
            v = sh1.Cells(hi, c).Value
            headers.append(str(v).strip() if v is not None else "")
        rows = []
        for r in range(hi + 1, max_row + 1):
            row = []
            for c in range(1, max_col + 1):
                v = sh1.Cells(r, c).Value
                if isinstance(v, str):
                    row.append(v.strip())
                else:
                    row.append(v)
            rows.append(row)
        df = pd.DataFrame(rows, columns=headers).dropna(how="all")
        return df
    except Exception:
        return None
