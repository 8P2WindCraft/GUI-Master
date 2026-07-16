"""Tab „Datenansicht“: Excel-Daten (Blatt 1), Fortschritt und Status wie zuvor in der Hauptsteuerung."""
import os

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QProgressBar,
    QPushButton,
    QTableView,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from ..constants import PROJECT_FILE_PATH_KEYS, VERSION

_PATH_LABELS = {
    'excel_path': 'Excel-Datei',
    'vorlagen_ordner': 'Vorlagen-Ordner',
    'bilder_ordner': 'Bilder-Ordner',
    'export_ordner': 'Export-Ordner',
}


class ProjectDataViewMixin:
    """Excel-Vorschau (Blatt 1) und kompakte Projektübersicht im zweiten Tab."""

    def setup_datenansicht_tab(self, tab):
        layout = QVBoxLayout(tab)

        # --- Wie früher Reiter 1: Daten-Ansicht Blatt 1 + Tabelle ---
        ansicht_row = QHBoxLayout()
        ansicht_row.addWidget(QLabel("Daten-Ansicht (Blatt 1)"))
        ansicht_row.addStretch()
        self.excel_ansicht_writeback_btn = QPushButton("Zurückschreiben")
        self.excel_ansicht_writeback_btn.setToolTip(
            "Liest die Excel-Datei erneut und aktualisiert die Tabelle in der Ansicht – "
            "z. B. nach Speichern in Excel."
        )
        self.excel_ansicht_writeback_btn.clicked.connect(self._on_excel_ansicht_zurueckschreiben)
        ansicht_row.addWidget(self.excel_ansicht_writeback_btn)
        layout.addLayout(ansicht_row)

        self.table = QTableView()
        self.table.setAlternatingRowColors(True)
        self.table.setShowGrid(True)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setHighlightSections(False)
        layout.addWidget(self.table, 1)

        self.progress = QProgressBar()
        layout.addWidget(self.progress)
        self.current_file_label = QLabel("Bereit zum Starten...")
        self.current_file_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.current_file_label)

        self.open_export_folder_btn = QPushButton("Export-Ordner öffnen")
        self.open_export_folder_btn.clicked.connect(self.open_export_folder)
        self.open_export_folder_btn.setVisible(False)
        layout.addWidget(self.open_export_folder_btn)

        layout.addWidget(QFrame(frameShape=QFrame.HLine))

        # --- Kompakte Projektübersicht (Text) ---
        meta_title = QLabel("Projektübersicht")
        meta_title.setStyleSheet("font-weight: 600;")
        layout.addWidget(meta_title)
        self.datenansicht_text = QTextEdit()
        self.datenansicht_text.setReadOnly(True)
        self.datenansicht_text.setPlaceholderText("Keine Metadaten geladen.")
        self.datenansicht_text.setMaximumHeight(220)
        layout.addWidget(self.datenansicht_text)

        meta_btn_row = QHBoxLayout()
        refresh_meta_btn = QPushButton("Projektübersicht aktualisieren")
        refresh_meta_btn.setToolTip("Pfade und Projektdaten aus dem aktuellen Zustand neu aufbauen")
        refresh_meta_btn.clicked.connect(self.refresh_datenansicht_meta)
        meta_btn_row.addWidget(refresh_meta_btn)
        meta_btn_row.addStretch()
        layout.addLayout(meta_btn_row)

        self.refresh_datenansicht_meta()

    def refresh_datenansicht(self):
        """Tab gewechselt: Excel neu einlesen und Metadaten aktualisieren."""
        if hasattr(self, 'show_excel_data'):
            self.show_excel_data()
        self.refresh_datenansicht_meta()

    def refresh_datenansicht_meta(self):
        """Nur Text-Übersicht (Projektpfade, Kategorien, …)."""
        if not hasattr(self, 'datenansicht_text'):
            return

        lines = []
        lines.append("=== Projektdaten ===")
        lines.append("")
        proj = getattr(self, 'current_project_path', None) or ""
        if proj:
            lines.append(f"Aktuelle Projektdatei (.dta.json):\n  {proj}")
        else:
            lines.append("Aktuelle Projektdatei: (keine gespeichert / nur Einstellungen)")
        lines.append("")

        root = (self.settings.get('project_root_dir', '') or '').strip()
        if hasattr(self, 'project_root_dir_edit'):
            root = self.project_root_dir_edit.text().strip() or root
        lines.append(f"Zentrales Projektverzeichnis:\n  {root or '—'}")
        lines.append("")

        lines.append("=== Pfade ===")
        for key in PROJECT_FILE_PATH_KEYS:
            label = _PATH_LABELS.get(key, key)
            val = (self.paths.get(key, '') or '').strip()
            lines.append(f"{label}:\n  {val or '—'}")
        lines.append("")

        lines.append("=== Export / Name ===")
        override = self.settings.get('projekt_name_override', '')
        if hasattr(self, 'projekt_name_override_edit'):
            override = self.projekt_name_override_edit.text().strip()
        lines.append(f"Projektname (Export-Ordner, optional):\n  {override or '— (aus Excel)'}")
        lines.append("")

        lines.append("=== Kategorien (Präfix → Anzeigename) ===")
        cats = getattr(self, 'categories', {}) or {}
        if cats:
            for pref, name in sorted(cats.items(), key=lambda x: str(x[0]).lower()):
                lines.append(f"  {pref!r} → {name}")
        else:
            lines.append("  —")
        lines.append("")

        lines.append("=== Regeln (Beschilderung) ===")
        lines.append(f"  Regeln aktiv: {'Ja' if getattr(self, 'rules_enabled', False) else 'Nein'}")
        lines.append(
            f"  Lageplan aus letztem Export: "
            f"{'Ja' if getattr(self, 'reuse_lageplan_from_last_export', False) else 'Nein'}"
        )
        rules = getattr(self, 'signage_rules', None) or []
        lines.append(f"  Anzahl Regel-Einträge: {len(rules)}")
        lines.append("")

        lines.append("=== Vorlagen-Auswahl ===")
        sel = getattr(self, 'selected_template_rel_paths', None)
        if sel is None:
            lines.append("  Modus: alle Vorlagen (keine eingeschränkte Teilmenge)")
        elif not sel:
            lines.append("  Modus: explizite Auswahl — derzeit keine gewählt")
        else:
            lines.append(f"  Modus: Teilmenge ({len(sel)} Vorlage(n))")
            for rel in sorted(sel)[:20]:
                lines.append(f"    • {rel}")
            if len(sel) > 20:
                lines.append(f"    … und {len(sel) - 20} weitere")
        lines.append("")

        lines.append("=== Letzte Projekte (Kurz) ===")
        recent = getattr(self, 'recent_projects', []) or []
        if recent:
            for i, p in enumerate(recent[:6], 1):
                if os.path.isfile(p):
                    lines.append(f"  {i}. {os.path.basename(p)}")
                    lines.append(f"     {p}")
                else:
                    lines.append(f"  {i}. (fehlt) {p}")
        else:
            lines.append("  —")
        lines.append("")

        lines.append(f"=== App ===\n  Version: {VERSION}")

        self.datenansicht_text.setPlainText("\n".join(lines))
