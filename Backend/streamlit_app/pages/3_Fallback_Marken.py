# -*- coding: utf-8 -*-
"""Seite: Fallback-Marken (Excel Blatt 2)."""
import streamlit as st
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.auth import require_admin
from utils.supabase_client import get_supabase
from utils.helpers import rows_to_list

require_admin()
sb = get_supabase()

st.title("Fallback-Marken")
st.caption("Projektweite Key-Value-Daten (entspricht Excel Blatt 2).")

projekte = sb.table("projekte").select("id, name").order("name").execute()
proj_list = rows_to_list(projekte.data)
if not proj_list:
    st.warning("Bitte zuerst ein Projekt anlegen.")
    st.stop()

proj_options = {p["name"]: p["id"] for p in proj_list}
projekt_name = st.selectbox("Projekt", options=list(proj_options.keys()))
projekt_id = proj_options[projekt_name]

# Fallback-Marken laden
marken = sb.table("fallback_marken").select("*").eq("projekt_id", projekt_id).order("schluessel").execute()
marken_list = rows_to_list(marken.data)

with st.expander("Neue Marke hinzufügen"):
    with st.form("neu_marke"):
        schluessel = st.text_input("Schlüssel *", placeholder="z.B. projekt_name")
        wert = st.text_area("Wert", placeholder="z.B. Windpark Göhlenkamp")
        if st.form_submit_button("Hinzufügen"):
            if schluessel.strip():
                try:
                    sb.table("fallback_marken").upsert({
                        "projekt_id": projekt_id,
                        "schluessel": schluessel.strip(),
                        "wert": (wert or "").strip(),
                    }, on_conflict="projekt_id,schluessel").execute()
                    st.success("Marke gespeichert.")
                    st.rerun()
                except Exception as e:
                    st.error(str(e))
            else:
                st.warning("Schlüssel erforderlich.")

st.subheader("Bestehende Marken")
if not marken_list:
    st.info("Keine Fallback-Marken für dieses Projekt.")
else:
    for m in marken_list:
        col1, col2, col3 = st.columns([2, 3, 1])
        with col1:
            st.text_input("Schlüssel", value=m.get("schluessel", ""), key=f"k_{m['id']}", disabled=True)
        with col2:
            new_val = st.text_input("Wert", value=m.get("wert", ""), key=f"v_{m['id']}")
        with col3:
            if st.button("Aktualisieren", key=f"upd_{m['id']}"):
                try:
                    sb.table("fallback_marken").update({"wert": new_val}).eq("id", m["id"]).execute()
                    st.success("Gespeichert.")
                    st.rerun()
                except Exception as e:
                    st.error(str(e))
            if st.button("Löschen", key=f"del_{m['id']}"):
                try:
                    sb.table("fallback_marken").delete().eq("id", m["id"]).execute()
                    st.rerun()
                except Exception as e:
                    st.error(str(e))
