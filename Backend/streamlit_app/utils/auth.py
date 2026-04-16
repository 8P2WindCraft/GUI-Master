# -*- coding: utf-8 -*-
"""Auth-Hilfen für Admin-Bereich (vereinfacht für Prototyp)."""
import streamlit as st


def check_admin_session() -> bool:
    """Prüft, ob Admin-Session aktiv ist. Für Prototyp: immer True."""
    if "admin_logged_in" not in st.session_state:
        st.session_state["admin_logged_in"] = True  # Prototyp: automatisch eingeloggt
    return st.session_state.get("admin_logged_in", False)


def require_admin():
    """Stellt sicher, dass Admin eingeloggt ist. Sonst Abbruch."""
    if not check_admin_session():
        st.error("Bitte melden Sie sich an.")
        st.stop()
