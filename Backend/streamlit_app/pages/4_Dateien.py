# -*- coding: utf-8 -*-
"""Seite: Dateien hochladen und verwalten."""
import streamlit as st
import sys
import os
import uuid
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.auth import require_admin
from utils.supabase_client import get_supabase
from utils.helpers import rows_to_list

require_admin()
sb = get_supabase()

st.title("Dateien")
st.caption("Bilder und Dokumente pro Projekt/Anlage hochladen.")

projekte = sb.table("projekte").select("id, name").order("name").execute()
proj_list = rows_to_list(projekte.data)
if not proj_list:
    st.warning("Bitte zuerst ein Projekt anlegen.")
    st.stop()

proj_options = {p["name"]: p["id"] for p in proj_list}
projekt_name = st.selectbox("Projekt", options=list(proj_options.keys()))
projekt_id = proj_options[projekt_name]

# Anlagen für Dropdown
anlagen = sb.table("anlagen").select("id, seriennummer").eq("projekt_id", projekt_id).order("seriennummer").execute()
anlagen_list = rows_to_list(anlagen.data)
anlage_options = {"(Allgemein)": None}
anlage_options.update({a["seriennummer"]: a["id"] for a in anlagen_list})

# Upload
st.subheader("Datei hochladen")
col1, col2 = st.columns(2)
with col1:
    anlage_sel = st.selectbox("Anlage", options=list(anlage_options.keys()))
    anlage_id = anlage_options.get(anlage_sel)

uploaded = st.file_uploader("Datei wählen", type=["png", "jpg", "jpeg", "gif", "svg", "pdf", "docx"])
if uploaded:
    if st.button("Hochladen"):
        try:
            bucket = "projekt-dateien"
            ext = os.path.splitext(uploaded.name)[1]
            storage_path = f"{projekt_id}/{anlage_id or 'allgemein'}/{uuid.uuid4()}{ext}"
            sb.storage.from_(bucket).upload(
                storage_path,
                uploaded.getvalue(),
                {"content-type": uploaded.type or "application/octet-stream"},
            )
            sb.table("dateien").insert({
                "projekt_id": projekt_id,
                "anlage_id": anlage_id,
                "dateiname": uploaded.name,
                "storage_path": storage_path,
                "mime_type": uploaded.type,
                "hochgeladen_von": "admin",
            }).execute()
            st.success(f"Datei '{uploaded.name}' hochgeladen.")
            st.rerun()
        except Exception as e:
            st.error(f"Fehler: {e}")

# Bucket ggf. anlegen (einmalig)
try:
    buckets = sb.storage.list_buckets()
    if not any(b.name == "projekt-dateien" for b in buckets):
        st.info("Bucket 'projekt-dateien' existiert noch nicht. Bitte in Supabase Dashboard anlegen.")
except Exception as e:
    st.warning(f"Storage-Zugriff: {e}")

# Bestehende Dateien
st.subheader("Bereits hochgeladene Dateien")
dateien = sb.table("dateien").select("*").eq("projekt_id", projekt_id).order("created_at", desc=True).execute()
dateien_list = rows_to_list(dateien.data)
if not dateien_list:
    st.info("Noch keine Dateien hochgeladen.")
else:
    for d in dateien_list:
        col1, col2, col3 = st.columns([3, 1, 1])
        with col1:
            st.write(f"**{d.get('dateiname', '—')}**")
            st.caption(f"Anlage: {d.get('anlage_id') or 'Allgemein'} | {d.get('created_at', '')[:19]}")
        with col2:
            if st.button("Löschen", key=f"del_{d['id']}"):
                try:
                    sb.storage.from_("projekt-dateien").remove([d["storage_path"]])
                    sb.table("dateien").delete().eq("id", d["id"]).execute()
                    st.rerun()
                except Exception as e:
                    st.error(str(e))
