# -*- coding: utf-8 -*-
"""Seite: Anlagen verwalten."""
import streamlit as st
import pandas as pd
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from utils.auth import require_admin
from utils.supabase_client import get_supabase
from utils.helpers import rows_to_list

require_admin()
sb = get_supabase()

st.title("Anlagen verwalten")
st.caption("Anlagen und deren Daten bearbeiten.")

# Projekt auswählen
projekte = sb.table("projekte").select("id, name").order("name").execute()
proj_list = rows_to_list(projekte.data)
if not proj_list:
    st.warning("Bitte zuerst ein Projekt anlegen (Seite Projekte).")
    st.stop()

proj_options = {p["name"]: p["id"] for p in proj_list}
projekt_name = st.selectbox("Projekt", options=list(proj_options.keys()))
projekt_id = proj_options[projekt_name]

# Anlagen laden
anlagen = sb.table("anlagen").select("*").eq("projekt_id", projekt_id).order("seriennummer").execute()
anlagen_list = rows_to_list(anlagen.data)

# Anlagen-Daten laden
anlage_ids = [a["id"] for a in anlagen_list]
daten_list = []
if anlage_ids:
    daten = sb.table("anlagen_daten").select("*").in_("anlage_id", anlage_ids).execute()
    daten_list = rows_to_list(daten.data)

# Zu dict pro Anlage gruppieren
anlage_daten_map = {}
for d in daten_list:
    aid = d.get("anlage_id")
    if aid not in anlage_daten_map:
        anlage_daten_map[aid] = {}
    anlage_daten_map[aid][d["schluessel"]] = d.get("wert", "")

# Alle Schlüssel sammeln
all_keys = set()
for m in anlage_daten_map.values():
    all_keys.update(m.keys())
all_keys = sorted(all_keys) if all_keys else ["anlage_seriennummer"]

# Tabelle als DataFrame
rows = []
for a in anlagen_list:
    row = {"id": a["id"], "seriennummer": a.get("seriennummer", "")}
    for k in all_keys:
        row[k] = anlage_daten_map.get(a["id"], {}).get(k, "")
    rows.append(row)

if rows:
    df = pd.DataFrame(rows)
    st.dataframe(df.drop(columns=["id"]), use_container_width=True, hide_index=True)
else:
    st.info("Keine Anlagen. Neue anlegen:")

with st.expander("Neue Anlage hinzufügen"):
    with st.form("neu_anlage"):
        sn = st.text_input("Seriennummer *", placeholder="z.B. ANL-001")
        if st.form_submit_button("Anlegen"):
            if sn.strip():
                try:
                    sb.table("anlagen").insert({
                        "projekt_id": projekt_id,
                        "seriennummer": sn.strip(),
                    }).execute()
                    st.success("Anlage angelegt.")
                    st.rerun()
                except Exception as e:
                    st.error(str(e))
            else:
                st.warning("Seriennummer erforderlich.")
