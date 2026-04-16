-- Supabase Backend - Initiales Schema
-- Ausführen in Supabase SQL Editor

-- Organisationen (Multi-Tenancy)
CREATE TABLE organisations (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  name TEXT NOT NULL,
  created_at TIMESTAMPTZ DEFAULT now()
);

-- Projekte
CREATE TABLE projekte (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  organisation_id UUID REFERENCES organisations(id),
  name TEXT NOT NULL,
  export_ordner_name TEXT,
  vorlagen_ordner_path TEXT,
  bilder_ordner_path TEXT,
  header_row INT DEFAULT 3,
  created_at TIMESTAMPTZ DEFAULT now(),
  updated_at TIMESTAMPTZ DEFAULT now()
);

-- Anlagen (eine Zeile pro anlage_seriennummer)
CREATE TABLE anlagen (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  projekt_id UUID REFERENCES projekte(id) ON DELETE CASCADE,
  seriennummer TEXT NOT NULL,
  created_at TIMESTAMPTZ DEFAULT now(),
  UNIQUE(projekt_id, seriennummer)
);

-- Anlagen-Daten: flexible Key-Value wie Excel-Spalten
CREATE TABLE anlagen_daten (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  anlage_id UUID REFERENCES anlagen(id) ON DELETE CASCADE,
  schluessel TEXT NOT NULL,
  wert TEXT,
  UNIQUE(anlage_id, schluessel)
);

-- Fallback-Marken (Excel Blatt 2)
CREATE TABLE fallback_marken (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  projekt_id UUID REFERENCES projekte(id) ON DELETE CASCADE,
  schluessel TEXT NOT NULL,
  wert TEXT,
  UNIQUE(projekt_id, schluessel)
);

-- Dateien (Bilder, PDFs etc., Storage-Pfad in Supabase)
CREATE TABLE dateien (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  projekt_id UUID REFERENCES projekte(id),
  anlage_id UUID REFERENCES anlagen(id),
  dateiname TEXT NOT NULL,
  storage_path TEXT NOT NULL,
  mime_type TEXT,
  hochgeladen_von TEXT,
  created_at TIMESTAMPTZ DEFAULT now()
);

-- Kunden-Zugangslinks
CREATE TABLE kunden_anfragen (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  projekt_id UUID REFERENCES projekte(id),
  token TEXT UNIQUE NOT NULL,
  beschreibung TEXT,
  anlage_ids UUID[],
  felder TEXT[],
  upload_erlaubt BOOLEAN DEFAULT true,
  ablaufdatum TIMESTAMPTZ,
  created_at TIMESTAMPTZ DEFAULT now()
);

-- Kunden-Uploads
CREATE TABLE kunden_uploads (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  kunden_anfrage_id UUID REFERENCES kunden_anfragen(id),
  anlage_id UUID REFERENCES anlagen(id),
  datei_id UUID REFERENCES dateien(id),
  feld_schluessel TEXT,
  kommentar TEXT,
  created_at TIMESTAMPTZ DEFAULT now()
);

-- Projekt-Kategorien
CREATE TABLE projekt_kategorien (
  id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  projekt_id UUID REFERENCES projekte(id),
  praefix TEXT NOT NULL,
  ordner_name TEXT NOT NULL,
  UNIQUE(projekt_id, praefix)
);

-- Standard-Organisation anlegen
INSERT INTO organisations (name) VALUES ('Standard');
