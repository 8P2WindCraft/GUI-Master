# -*- coding: utf-8 -*-
"""Kunden-Frontend: Token-basierte Dateneingabe und Datei-Upload."""
import streamlit as st
import sys
import os
import uuid
from datetime import datetime, timezone

backend_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(backend_root, "streamlit_app"))
sys.path.insert(0, backend_root)

from utils.supabase_client import get_supabase
from utils.helpers import rows_to_list

st.set_page_config(page_title="Dateneingabe", page_icon="📝", layout="wide")

# Token aus URL
query_params = st.query_params
token = query_params.get("token", "")

if not token:
    st.title("Dateneingabe für Kunden")
    token_input = st.text_input("Bitte geben Sie Ihren Zugangstoken ein (oder nutzen Sie den Link):", placeholder="Token hier einfügen...")
    if token_input:
        st.query_params["token"] = token_input
        st.rerun()
    st.info("Sie sollten einen Link mit Token von Ihrem Ansprechpartner erhalten haben.")
    st.stop()

# Supabase
try:
    sb = get_supabase()
except Exception as e:
    st.error("Backend nicht erreichbar. Bitte später erneut versuchen.")
    st.stop()

# Kunden-Anfrage laden
anfragen = sb.table("kunden_anfragen").select("*").eq("token", token).execute()
if not anfragen.data:
    st.error("Ungültiger oder abgelaufener Token.")
    st.stop()

anfrage = anfragen.data[0]
projekt_id = anfrage["projekt_id"]
anlage_ids = anfrage.get("anlage_ids") or []
erlaubte_felder = anfrage.get("felder")  # None = alle
upload_erlaubt = anfrage.get("upload_erlaubt", True)
ablauf = anfrage.get("ablaufdatum")
if ablauf and datetime.fromisoformat(ablauf.replace("Z", "+00:00")) < datetime.now(timezone.utc):
    st.error("Dieser Link ist abgelaufen.")
    st.stop()

st.title(anfrage.get("beschreibung") or "Dateneingabe")
st.caption("Bitte füllen Sie die folgenden Felder aus.")

# Anlagen laden
anlagen_query = sb.table("anlagen").select("id, seriennummer").eq("projekt_id", projekt_id)
if anlage_ids:
    anlagen_query = anlagen_query.in_("id", anlage_ids)
anlagen = anlagen_query.order("seriennummer").execute()
anlagen_list = rows_to_list(anlagen.data)

if not anlagen_list:
    st.warning("Keine Anlagen zur Bearbeitung freigegeben.")
    st.stop()

# Daten pro Anlage
for anlage in anlagen_list:
    aid = anlage["id"]
    sn = anlage.get("seriennummer", "")
    st.subheader(f"Anlage: {sn}")

    daten = sb.table("anlagen_daten").select("schluessel, wert").eq("anlage_id", aid).execute()
    daten_list = rows_to_list(daten.data)
    daten_map = {d["schluessel"]: d.get("wert", "") for d in daten_list}

    keys = sorted(daten_map.keys())
    if erlaubte_felder:
        keys = [k for k in keys if k in erlaubte_felder]
    if not keys and erlaubte_felder:
        keys = list(erlaubte_felder)

    for key in keys:
        val = daten_map.get(key, "")
        new_val = st.text_input(key, value=val, key=f"{aid}_{key}")
        if new_val != val:
            try:
                sb.table("anlagen_daten").upsert({
                    "anlage_id": aid,
                    "schluessel": key,
                    "wert": new_val,
                }, on_conflict="anlage_id,schluessel").execute()
                st.success(f"Gespeichert: {key}")
            except Exception as e:
                st.error(str(e))

    if upload_erlaubt:
        st.write("**Datei hochladen**")
        up = st.file_uploader("Datei", key=f"up_{aid}", type=["png", "jpg", "jpeg", "pdf", "docx"])
        if up and st.button("Hochladen", key=f"btn_{aid}"):
            try:
                bucket = "kunden-uploads"
                ext = os.path.splitext(up.name)[1]
                storage_path = f"{anfrage['id']}/{aid}/{uuid.uuid4()}{ext}"
                sb.storage.from_(bucket).upload(storage_path, up.getvalue(), {"content-type": up.type or "application/octet-stream"})
                file_ins = sb.table("dateien").insert({
                    "projekt_id": projekt_id,
                    "anlage_id": aid,
                    "dateiname": up.name,
                    "storage_path": storage_path,
                    "mime_type": up.type,
                    "hochgeladen_von": "kunde",
                }).execute()
                sb.table("kunden_uploads").insert({
                    "kunden_anfrage_id": anfrage["id"],
                    "anlage_id": aid,
                    "datei_id": file_ins.data[0]["id"],
                }).execute()
                st.success(f"Datei '{up.name}' hochgeladen.")
                st.rerun()
            except Exception as e:
                st.error(f"Fehler: {e}")

    st.divider()

st.success("Vielen Dank für Ihre Eingaben!")
