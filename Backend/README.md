# DocxTpl Supabase Backend

Migration von Excel zu Supabase mit Admin-Frontend und Kunden-Frontend.

## Voraussetzungen

- Python 3.10+
- Supabase-Account (https://supabase.com)

## Einrichtung

### 1. Supabase-Projekt anlegen

1. Projekt bei Supabase erstellen
2. Im SQL Editor das Skript `migrations/001_initial.sql` ausführen
3. Storage-Buckets anlegen:
   - `projekt-dateien` (für Admin-Uploads)
   - `kunden-uploads` (für Kunden-Uploads)

### 2. Umgebung konfigurieren

```powershell
cd Backend
copy .env.example .env
```

`.env` bearbeiten und eintragen:
- `SUPABASE_URL` – aus Supabase Projekt → Settings → API
- `SUPABASE_SERVICE_ROLE_KEY` – aus Supabase Projekt → Settings → API

### 3. Abhängigkeiten installieren

```powershell
pip install -r requirements.txt
```

### 4. Anwendung starten

**Admin-Frontend (Port 8501):**
```powershell
.\run_admin.ps1
```
oder
```powershell
streamlit run streamlit_app/Home.py --server.port 8501
```

**Kunden-Frontend (Port 8502):**
```powershell
.\run_kunden.ps1
```
oder
```powershell
streamlit run kunden_frontend/app.py --server.port 8502
```

## Struktur

```
Backend/
├── migrations/          # SQL-Skripte
├── streamlit_app/       # Admin-Frontend
│   ├── Home.py          # Startseite
│   ├── pages/           # Unterseiten
│   └── utils/           # Supabase-Client, Hilfen
├── kunden_frontend/     # Kunden-Frontend (token-basiert)
├── run_admin.ps1
├── run_kunden.ps1
└── requirements.txt
```

## Funktionen

- **Projekte** – Projekte anlegen und verwalten
- **Anlagen** – Anlagen und flexible Key-Value-Daten
- **Fallback-Marken** – Projektweite Daten (Excel Blatt 2)
- **Dateien** – Upload von Bildern/Dokumenten
- **Kunden-Links** – Token erstellen für gezielte Kunden-Dateneingabe
- **Excel-Import** – Bestehende Excel-Daten migrieren
- **Einstellungen** – Kategorien pro Projekt

## Kunden-Link

Nach Erstellung eines Kunden-Links (Seite "Kunden-Links") erhält der Kunde z.B.:
```
http://localhost:8502?token=abc-123-uuid
```
Der Kunde kann damit nur die freigegebenen Anlagen und Felder bearbeiten sowie ggf. Dateien hochladen.
