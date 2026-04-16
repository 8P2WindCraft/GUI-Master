# -*- coding: utf-8 -*-
"""Seite: Kunden-Links erstellen."""
import streamlit as st
import sys
import os
import uuid
from datetime import datetime, timedelta

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.auth import require_admin
from utils.supabase_client import get_supabase
from utils.helpers import rows_to_list

require_admin()
sb = get_supabase()

st.title("Kunden-Links")
st.caption("Token-Links erstellen, damit Kunden gezielt Daten eingeben und Dateien hochladen können.")

projekte = sb.table("projekte").select("id, name").order("name").execute()
proj_list = rows_to_list(projekte.data)
if not proj_list:
    st.warning("Bitte zuerst ein Projekt anlegen.")
    st.stop()

proj_options = {p["name"]: p["id"] for p in proj_list}
projekt_name = st.selectbox("Projekt", options=list(proj_options.keys()))
projekt_id = proj_options[projekt_name]

anlagen = sb.table("anlagen").select("id, seriennummer").eq("projekt_id", projekt_id).order("seriennummer").execute()
anlagen_list = rows_to_list(anlagen.data)

with st.expander("Neuen Kunden-Link erstellen"):
    with st.form("neu_link"):
        beschreibung = st.text_input("Beschreibung", placeholder="z.B. Anlage 1-5 Daten ergänzen")
        anlage_ids = st.multiselect(
            "Anlagen (leer = alle)",
            options=[a["id"] for a in anlagen_list],
            format_func=lambda x: next((a["seriennummer"] for a in anlagen_list if a["id"] == x), str(x)),
        )
        felder = st.text_input("Erlaubte Felder (kommagetrennt, leer = alle)", placeholder="z.B. standort, foto_img, anmeldung_link")
        upload_erlaubt = st.checkbox("Datei-Upload erlauben", value=True)
        tage = st.number_input("Gültigkeit in Tagen", min_value=1, value=30)
        if st.form_submit_button("Link erstellen"):
            token = str(uuid.uuid4())
            ablauf = datetime.utcnow() + timedelta(days=int(tage))
            try:
                sb.table("kunden_anfragen").insert({
                    "projekt_id": projekt_id,
                    "token": token,
                    "beschreibung": beschreibung.strip() or None,
                    "anlage_ids": anlage_ids if anlage_ids else None,
                    "felder": [f.strip() for f in felder.split(",") if f.strip()] if felder else None,
                    "upload_erlaubt": upload_erlaubt,
                    "ablaufdatum": ablauf.isoformat(),
                }).execute()
                base_url = os.getenv("KUNDEN_APP_URL", "http://localhost:8502")
                link = f"{base_url}?token={token}"
                st.success("Link erstellt!")
                st.code(link, language=None)
                st.caption("Diesen Link an den Kunden senden.")
            except Exception as e:
                st.error(str(e))

st.subheader("Bestehende Links")
links = sb.table("kunden_anfragen").select("*").eq("projekt_id", projekt_id).order("created_at", desc=True).execute()
links_list = rows_to_list(links.data)
if not links_list:
    st.info("Noch keine Kunden-Links erstellt.")
else:
    base_url = os.getenv("KUNDEN_APP_URL", "http://localhost:8502")
    for l in links_list:
        with st.container():
            st.write(f"**{l.get('beschreibung') or 'Ohne Beschreibung'}**")
            st.caption(f"Token: {l.get('token', '')[:8]}... | Ablauf: {l.get('ablaufdatum', '—')[:10]}")
            st.code(f"{base_url}?token={l['token']}", language=None)
            if st.button("Löschen", key=f"del_{l['id']}"):
                try:
                    sb.table("kunden_anfragen").delete().eq("id", l["id"]).execute()
                    st.rerun()
                except Exception as e:
                    st.error(str(e))
