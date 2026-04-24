"""
Zentrale Konstanten der GUI-App (Version, Pfade, Regel-UI, Ressourcen-Filter).
"""

# Version / Build
VERSION = "7.0.1"
BUILD_DATE = "2025-01-27"

# Unterstützte Bildformate für den Bilder-Vorschau-Tab
BILDER_VORSCHAU_EXTENSIONS = {'.png', '.jpg', '.jpeg', '.gif', '.svg', '.bmp'}

# (value, label) für Textbedingungen im Regeleditor
RULE_CONDITION_CHOICES = (
    ("not_empty", "Nicht leer (nach Trim)"),
    ("length_gt", "Text länger als … Zeichen (>)"),
    ("length_gte", "Mindestlänge … Zeichen (≥)"),
    ("contains", "Enthält Text …"),
    ("equals", "Exakt gleich …"),
    ("equals_ignorecase", "Gleich … (Groß/Klein egal)"),
    ("regex", "Erfüllt Regex …"),
)

# Maximale Anzahl Einträge im Untermenü "Letzte Projekte"
MAX_RECENT_PROJECTS = 6

# Pfade: feste Keys für self.paths, load_settings, Projekt-JSON (neu + flaches Alt-Format)
# In dieser Reihenfolge auch an UI-Eingaben (…_edit) angebunden.
PROJECT_FILE_PATH_KEYS = (
    'excel_path', 'vorlagen_ordner', 'bilder_ordner', 'export_ordner',
)
