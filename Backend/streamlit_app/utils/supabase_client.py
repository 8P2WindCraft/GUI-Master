# -*- coding: utf-8 -*-
"""Supabase-Client für das Backend."""
import os
from supabase import create_client, Client
from dotenv import load_dotenv

load_dotenv()

_url = os.getenv("SUPABASE_URL")
_key = os.getenv("SUPABASE_SERVICE_ROLE_KEY")

_supabase: Client | None = None


def get_supabase() -> Client:
    """Singleton Supabase-Client."""
    global _supabase
    if _supabase is None:
        if not _url or not _key:
            raise RuntimeError(
                "SUPABASE_URL und SUPABASE_SERVICE_ROLE_KEY müssen in .env gesetzt sein."
            )
        _supabase = create_client(_url, _key)
    return _supabase


def is_configured() -> bool:
    """Prüft, ob Supabase konfiguriert ist."""
    return bool(_url and _key)
