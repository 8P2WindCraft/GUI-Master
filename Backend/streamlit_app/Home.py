# -*- coding: utf-8 -*-
"""Hauptseite der Admin-App - DocxTpl Supabase Backend."""
import streamlit as st
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from utils.auth import check_admin_session
from utils.supabase_client import get_supabase, is_configured

st.set_page_config(
    page_title="DocxTpl Backend",
    page_icon="📋",
    layout="wide",
)

st.title("DocxTpl Supabase Backend")
st.caption("Migration von Excel zu Supabase – Admin-Frontend")

# Konfigurationsprüfung
if not is_configured():
    st.error(
        "Supabase ist nicht konfiguriert. Bitte erstellen Sie eine `.env`-Datei "
        "mit SUPABASE_URL und SUPABASE_SERVICE_ROLE_KEY. Siehe `.env.example`."
    )
    st.stop()

try:
    sb = get_supabase()
    # Quick-Check: Organisationen laden
    orgs = sb.table("organisations").select("id, name").execute()
    st.success("Verbindung zu Supabase erfolgreich.")
except Exception as e:
    st.error(f"Verbindungsfehler: {e}")
    st.stop()

# Dashboard-Inhalt
check_admin_session()

col1, col2, col3 = st.columns(3)
with col1:
    st.metric("Organisationen", len(orgs.data) if orgs.data else 0)
with col2:
    try:
        projekte = sb.table("projekte").select("id", count="exact").execute()
        st.metric("Projekte", projekte.count or 0)
    except Exception:
        st.metric("Projekte", "—")
with col3:
    try:
        anlagen = sb.table("anlagen").select("id", count="exact").execute()
        st.metric("Anlagen", anlagen.count or 0)
    except Exception:
        st.metric("Anlagen", "—")

st.divider()
st.subheader("Schnellnavigation")
st.markdown("""
- **Projekte** – Projekte anlegen und verwalten
- **Anlagen verwalten** – Anlagen und deren Daten bearbeiten
- **Fallback-Marken** – Projektweite Key-Value-Daten (Excel Blatt 2)
- **Dateien** – Bilder und Dokumente hochladen
- **Kunden-Links** – Token für Kunden-Dateneingabe erstellen
- **Excel-Import** – Bestehende Excel-Daten migrieren
- **Einstellungen** – Kategorien und Optionen
""")
