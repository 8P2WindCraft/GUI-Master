# -*- coding: utf-8 -*-
"""Seite: Projekte verwalten."""
import streamlit as st
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.auth import require_admin
from utils.supabase_client import get_supabase
from utils.helpers import rows_to_list

require_admin()
sb = get_supabase()

st.title("Projekte")
st.caption("Projekte anlegen, bearbeiten und löschen.")

# Organisationen für Dropdown
orgs = sb.table("organisations").select("id, name").execute()
org_list = rows_to_list(orgs.data)
org_map = {str(o["id"]): o["name"] for o in org_list}

# Bestehende Projekte laden
projekte = sb.table("projekte").select("*").order("name").execute()
proj_list = rows_to_list(projekte.data)

with st.expander("Neues Projekt anlegen"):
    with st.form("neu_projekt"):
        org_id = st.selectbox("Organisation", options=list(org_map.keys()), format_func=lambda x: org_map.get(x, x))
        name = st.text_input("Projektname *", placeholder="z.B. Windpark Göhlenkamp")
        export_ordner = st.text_input("Export-Ordner Name", placeholder="z.B. 246778_GWS_WP")
        vorlagen = st.text_input("Vorlagen-Ordner Pfad", placeholder="C:/Pfad/zu/templates")
        bilder = st.text_input("Bilder-Ordner Pfad", placeholder="C:/Pfad/zu/pictures")
        header_row = st.number_input("Header-Zeile (Excel)", min_value=1, value=3)
        if st.form_submit_button("Anlegen"):
            if name:
                sb.table("projekte").insert({
                    "organisation_id": org_id,
                    "name": name.strip(),
                    "export_ordner_name": export_ordner.strip() or None,
                    "vorlagen_ordner_path": vorlagen.strip() or None,
                    "bilder_ordner_path": bilder.strip() or None,
                    "header_row": int(header_row),
                }).execute()
                st.success("Projekt angelegt.")
                st.rerun()
            else:
                st.warning("Projektname ist erforderlich.")

st.subheader("Bestehende Projekte")
if not proj_list:
    st.info("Noch keine Projekte vorhanden.")
else:
    for p in proj_list:
        with st.container():
            col1, col2 = st.columns([3, 1])
            with col1:
                st.write(f"**{p.get('name', '—')}**")
                st.caption(f"Export: {p.get('export_ordner_name') or '—'} | ID: {p.get('id', '')[:8]}...")
            with col2:
                if st.button("Löschen", key=f"del_{p['id']}"):
                    try:
                        sb.table("projekte").delete().eq("id", p["id"]).execute()
                        st.success("Gelöscht.")
                        st.rerun()
                    except Exception as e:
                        st.error(str(e))
