# -*- coding: utf-8 -*-
"""Seite: Einstellungen."""
import streamlit as st
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.auth import require_admin
from utils.supabase_client import get_supabase
from utils.helpers import rows_to_list

require_admin()
sb = get_supabase()

st.title("Einstellungen")
st.caption("Kategorien und Optionen.")

# Kategorien pro Projekt
projekte = sb.table("projekte").select("id, name").order("name").execute()
proj_list = rows_to_list(projekte.data)
if not proj_list:
    st.warning("Bitte zuerst ein Projekt anlegen.")
    st.stop()

proj_options = {p["name"]: p["id"] for p in proj_list}
projekt_name = st.selectbox("Projekt", options=list(proj_options.keys()))
projekt_id = proj_options[projekt_name]

kategorien = sb.table("projekt_kategorien").select("*").eq("projekt_id", projekt_id).execute()
kat_list = rows_to_list(kategorien.data)

st.subheader("Dokumentkategorien")
st.caption("Präfix → Ordnername (z.B. b_ → Beschilderung)")

with st.expander("Neue Kategorie"):
    with st.form("neu_kat"):
        praefix = st.text_input("Präfix", placeholder="z.B. b_")
        ordner = st.text_input("Ordnername", placeholder="z.B. Beschilderung")
        if st.form_submit_button("Hinzufügen"):
            if praefix and ordner:
                try:
                    sb.table("projekt_kategorien").upsert({
                        "projekt_id": projekt_id,
                        "praefix": praefix.strip(),
                        "ordner_name": ordner.strip(),
                    }, on_conflict="projekt_id,praefix").execute()
                    st.success("Kategorie hinzugefügt.")
                    st.rerun()
                except Exception as e:
                    st.error(str(e))

for k in kat_list:
    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        st.text_input("Präfix", value=k.get("praefix"), key=f"p_{k['id']}", disabled=True)
    with col2:
        st.text_input("Ordner", value=k.get("ordner_name"), key=f"o_{k['id']}", disabled=True)
    with col3:
        if st.button("Löschen", key=f"del_{k['id']}"):
            sb.table("projekt_kategorien").delete().eq("id", k["id"]).execute()
            st.rerun()
