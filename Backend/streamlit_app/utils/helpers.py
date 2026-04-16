# -*- coding: utf-8 -*-
"""Hilfsfunktionen für das Backend."""
from typing import Any


def row_to_dict(row: dict[str, Any] | None) -> dict[str, Any]:
    """Konvertiert Supabase-Row zu dict (für einfache Handhabung)."""
    if row is None:
        return {}
    return dict(row) if hasattr(row, "keys") else {}


def rows_to_list(rows: list) -> list[dict]:
    """Konvertiert Liste von Supabase-Rows zu Liste von dicts."""
    return [row_to_dict(r) for r in (rows or [])]
