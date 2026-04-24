"""

Optionale Beschilderungs-Regeln: pro Excel-Zeile und Vorlage entscheiden, ob ein Dokument erzeugt wird.

Regeln gruppieren typischerweise zwei sich ausschließende Vorlagen (z. B. mit / ohne Notausgangszeichen).

"""

from __future__ import annotations

import os

import re

from typing import Any

import pandas as pd

def normalize_template_rel(vorlagen_ordner: str, template_abs: str) -> str:

    """Relativer Pfad zur Vorlage unter vorlagen_ordner (wie in der GUI-Auswahl)."""

    return os.path.normpath(os.path.relpath(template_abs, vorlagen_ordner))

def merge_row_context(eintrag: dict, fallback_marken: dict) -> dict:

    """Wie in verarbeite_vorlagen: Zeile + Fallback für leere Felder."""

    ctx = dict(eintrag)

    for key, fallback_value in fallback_marken.items():

        primary_value = ctx.get(key)

        is_empty = pd.isna(primary_value) or (

            isinstance(primary_value, str) and not str(primary_value).strip()

        )

        if is_empty:

            ctx[key] = fallback_value

    return ctx

def text_fields_any_length_gt(ctx: dict, columns: list[str], min_len: int) -> bool:

    """True, wenn mindestens eine der Spalten einen String mit Länge > min_len hat."""

    if not columns:

        return False

    for col in columns:

        col = (col or "").strip()

        if not col:

            continue

        val = ctx.get(col)

        if val is None or (isinstance(val, float) and pd.isna(val)):

            continue

        s = str(val).strip()

        if len(s) > min_len:

            return True

    return False

def branch_image_file_exists(ctx: dict, bilder_ordner: str | None, column: str) -> bool:

    """True, wenn Spalte gesetzt ist und die Datei unter bilder_ordner existiert."""

    if not bilder_ordner or not column or not str(column).strip():

        return False

    val = ctx.get(column.strip())

    if val is None or (isinstance(val, float) and pd.isna(val)):

        return False

    s = str(val).strip()

    if not s:

        return False

    path = os.path.join(bilder_ordner, s)

    return os.path.isfile(path)

def _norm_rel_list(paths: list[str]) -> set[str]:

    return {os.path.normpath(p) for p in paths if p}

def _resolve_ctx_key(column: str, textmarke: str) -> str:

    """Lookup-Schlüssel im Kontext: bevorzugt explizite Textmarke, sonst Excel-Spaltenname."""

    tm = (textmarke or "").strip()

    if tm:

        return tm

    return (column or "").strip()

def _ctx_string_value(ctx: dict, column: str, textmarke: str) -> str | None:

    key = _resolve_ctx_key(column, textmarke)

    if not key:

        return None

    val = ctx.get(key)

    if val is None or (isinstance(val, float) and pd.isna(val)):

        return None

    return str(val)

def _field_satisfies_mode(

    s: str | None,

    mode: str,

    min_len: int,

    needle: str,

    equals_val: str,

    regex_pattern: str,

) -> bool:

    if s is None:

        s = ""

    st = s.strip()

    mode = (mode or "not_empty").strip().lower()

    if mode == "not_empty":

        return bool(st)

    if mode == "length_gt":

        return len(st) > int(min_len)

    if mode == "length_gte":

        return len(st) >= int(min_len)

    if mode == "contains":

        if not needle:

            return False

        return needle.lower() in st.lower()

    if mode == "equals":

        return st == (equals_val or "").strip()

    if mode == "equals_ignorecase":

        return st.lower() == (equals_val or "").strip().lower()

    if mode == "regex":

        if not regex_pattern:

            return False

        try:

            return re.search(regex_pattern, st) is not None

        except re.error:

            return False

    # Unbekannter Modus: konservativ wie „nicht leer“

    return bool(st)

def evaluate_text_fields_condition(ctx: dict, when: dict[str, Any]) -> tuple[bool, str, dict[str, Any]]:

    """

    Wertet when.text_condition aus; fällt auf Legacy text_fields_any_length_gt zurück.

    Rückgabe: (erfüllt, kurzer Grund, Debug-Dict)

    """

    meta: dict[str, Any] = {"mode": None, "fields": [], "legacy": False}

    if not when:

        return False, "Keine Bedingung konfiguriert", meta

    if "text_condition" in when:

        tc = when.get("text_condition") or {}

        mode = str(tc.get("mode") or "not_empty").strip().lower()

        min_len = int(tc.get("min", tc.get("min_len", 3)))

        needle = str(tc.get("needle", tc.get("contains", "")))

        equals_val = str(tc.get("equals_value", tc.get("equals", "")))

        regex_pattern = str(tc.get("regex_pattern", tc.get("regex", "")))

        raw_fields = tc.get("fields")

        if not raw_fields:

            fields = [{"column": c, "textmarke": ""} for c in (tc.get("columns") or [])]

        else:

            fields = list(raw_fields)

        meta["mode"] = mode

        meta["fields"] = fields

        if not fields:

            return False, "Keine Textfelder (Spalte/Textmarke) angegeben", meta

        for f in fields:

            if isinstance(f, str):

                col, tm = f.strip(), ""

            else:

                col = str((f or {}).get("column", "")).strip()

                tm = str((f or {}).get("textmarke", "")).strip()

            s = _ctx_string_value(ctx, col, tm)

            if _field_satisfies_mode(s, mode, min_len, needle, equals_val, regex_pattern):

                key = _resolve_ctx_key(col, tm)

                return True, f"Bedingung '{mode}' erfüllt (Schlüssel '{key}')", meta

        return False, f"Bedingung '{mode}' für kein konfiguriertes Feld erfüllt", meta

    # Legacy

    tcfg = when.get("text_fields_any_length_gt") or {}

    cols = tcfg.get("columns") or []

    if isinstance(cols, str):

        cols = [c.strip() for c in cols.split(",") if c.strip()]

    min_len = int(tcfg.get("min", 3))

    meta["legacy"] = True

    meta["mode"] = "length_gt"

    meta["fields"] = [{"column": c, "textmarke": ""} for c in cols]

    ok = text_fields_any_length_gt(ctx, cols, min_len)

    if ok:

        return True, f"Textbedingung erfüllt (mindestens eine Spalte > {min_len})", meta

    return False, f"Textbedingung nicht erfüllt (keine Spalte > {min_len})", meta

def find_rule_for_template(

    rules_list: list[dict[str, Any]] | None, template_rel: str

) -> tuple[dict[str, Any] | None, str | None]:

    """

    Findet die erste aktive Regel, die diese Vorlage (relativer Pfad) in templates_if oder templates_else führt.

    Rückgabe: (rule, 'if'|'else'|None)

    """

    if not rules_list:

        return None, None

    tr = os.path.normpath(template_rel)

    for rule in rules_list:

        if not rule.get("enabled", True):

            continue

        tif = _norm_rel_list(rule.get("templates_if") or [])

        tel = _norm_rel_list(rule.get("templates_else") or [])

        if tr in tif:

            return rule, "if"

        if tr in tel:

            return rule, "else"

    return None, None

def should_emit_anlage_template(

    vorlagen_ordner: str,

    template_abs: str,

    eintrag: dict,

    fallback_marken: dict,

    bilder_ordner: str | None,

    rules_enabled: bool,

    rules_list: list[dict[str, Any]] | None,

) -> bool:

    """

    Entscheidet, ob für diese Anlagen-Zeile diese Vorlage erzeugt wird.

    Ohne Regeln oder wenn Regeln aus: immer True.

    Zwei Modi:

    - Ohne ``branch_on_image.column``: Verzweigung nur über Text (mind. eine Spalte mit Länge > min

      → ``templates_if``, sonst ``templates_else``).

    - Mit Bild-Spalte: zuerst Textbedingung (sonst kein Dokument), dann Bild vorhanden → IF, sonst ELSE.

    """

    return evaluate_anlage_template_decision(

        vorlagen_ordner,

        template_abs,

        eintrag,

        fallback_marken,

        bilder_ordner,

        rules_enabled,

        rules_list,

    )["emit"]

def evaluate_anlage_template_decision(

    vorlagen_ordner: str,

    template_abs: str,

    eintrag: dict,

    fallback_marken: dict,

    bilder_ordner: str | None,

    rules_enabled: bool,

    rules_list: list[dict[str, Any]] | None,

) -> dict[str, Any]:

    """

    Liefert die Regel-Entscheidung inkl. Metadaten für Logging/Debug.

    """

    rel = normalize_template_rel(vorlagen_ordner, template_abs)

    result: dict[str, Any] = {

        "emit": True,

        "template_rel": rel,

        "has_rule": False,

        "rule_name": None,

        "configured_branch": None,

        "applied_branch": None,

        "reason": "Regeln deaktiviert oder nicht konfiguriert",

        "text_ok": None,

        "text_columns": [],

        "text_min_len": None,

        "text_condition_mode": None,

        "image_column": "",

        "image_exists": None,

    }

    if not rules_enabled or not rules_list:

        return result

    rule, configured_branch = find_rule_for_template(rules_list, rel)

    if rule is None:

        result["reason"] = "Keine passende Regel für Vorlage"

        return result

    result["has_rule"] = True

    result["configured_branch"] = configured_branch

    result["rule_name"] = str(rule.get("name") or rule.get("id") or "(ohne Name)")

    ctx = merge_row_context(eintrag, fallback_marken)

    when = rule.get("when") or {}

    text_ok, text_reason, tc_meta = evaluate_text_fields_condition(ctx, when)

    result["text_ok"] = text_ok

    result["text_condition_mode"] = tc_meta.get("mode")

    # Abwärtskompatibel befüllen

    if tc_meta.get("legacy"):

        tcfg = when.get("text_fields_any_length_gt") or {}

        cols = tcfg.get("columns") or []

        if isinstance(cols, str):

            cols = [c.strip() for c in cols.split(",") if c.strip()]

        result["text_columns"] = cols

        result["text_min_len"] = int(tcfg.get("min", 3))

    else:

        fields = tc_meta.get("fields") or []

        cols_disp: list[str] = []

        for f in fields:

            if isinstance(f, str):

                cols_disp.append(_resolve_ctx_key(f, ""))

            else:

                fd = f or {}

                cols_disp.append(

                    _resolve_ctx_key(str(fd.get("column", "")), str(fd.get("textmarke", "")))

                )

        result["text_columns"] = cols_disp

        if "text_condition" in when:

            tc = when.get("text_condition") or {}

            result["text_min_len"] = int(tc.get("min", tc.get("min_len", 3)))

    tif = _norm_rel_list(rule.get("templates_if") or [])

    tel = _norm_rel_list(rule.get("templates_else") or [])

    bio = rule.get("branch_on_image") or {}

    img_col = (bio.get("column") or "").strip()

    result["image_column"] = img_col

    # Ohne Bild-Spalte: Verzweigung nur über Textbedingung.

    if not img_col:

        if text_ok:

            result["applied_branch"] = "if"

            result["emit"] = rel in tif

            result["reason"] = text_reason

        else:

            result["applied_branch"] = "else"

            result["emit"] = rel in tel

            result["reason"] = text_reason

        return result

    # Mit Bild-Spalte: zuerst Textbedingung, dann Bild vorhanden => if/else.

    if not text_ok:

        result["emit"] = False

        result["applied_branch"] = None

        result["reason"] = text_reason

        result["image_exists"] = False

        return result

    has_img = branch_image_file_exists(ctx, bilder_ordner, img_col)

    result["image_exists"] = has_img

    if has_img:

        result["applied_branch"] = "if"

        result["emit"] = rel in tif

        result["reason"] = f"Bild in Spalte '{img_col}' vorhanden"

    else:

        result["applied_branch"] = "else"

        result["emit"] = rel in tel

        result["reason"] = f"Kein Bild in Spalte '{img_col}' gefunden"

    return result

def count_anlagen_emits_with_rules(

    datensaetze: list[dict],

    anlagen_templates: list[str],

    vorlagen_ordner: str,

    fallback_marken: dict,

    bilder_ordner: str | None,

    rules_enabled: bool,

    rules_list: list[dict[str, Any]] | None,

) -> int:

    """Anzahl der tatsächlich zu erzeugenden Anlagen-Dokumente unter Berücksichtigung der Regeln."""

    n = 0

    for eintrag in datensaetze:

        for tp in anlagen_templates:

            if should_emit_anlage_template(

                vorlagen_ordner,

                tp,

                eintrag,

                fallback_marken,

                bilder_ordner,

                rules_enabled,

                rules_list,

            ):

                n += 1

    return n

