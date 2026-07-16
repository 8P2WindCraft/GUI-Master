"""Hauptsteuerung: Pfade, Excel-Vorschau-Tabelle, Fortschritt, Start-Buttons."""
import os

from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QCheckBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from core.utils import resource_path

from ..constants import BUILD_DATE, VERSION


class MainControlsMixin:
    """Pfadeingaben, Tabelle, Progress, Trockenlauf/Vorschau/Start."""

    def setup_controls_tab(self, tab):
        """Erstellt den Inhalt des 'Hauptsteuerung'-Tabs."""
        layout = QVBoxLayout(tab)
        # Oben: ausgewählter Excel-Dateiname (breit, Tooltip = voller Pfad)
        name_row = QHBoxLayout()
        self.excel_basename_line = QLineEdit()
        self.excel_basename_line.setReadOnly(True)
        self.excel_basename_line.setFrame(False)
        self.excel_basename_line.setFocusPolicy(Qt.NoFocus)
        self.excel_basename_line.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed
        )
        self.excel_basename_line.setPlaceholderText("Kein Projektname verfügbar")
        name_font = self.excel_basename_line.font()
        name_font.setBold(True)
        self.excel_basename_line.setFont(name_font)
        name_row.addWidget(self.excel_basename_line, 1)
        layout.addLayout(name_row)

        path_map = {
            'excel_path': 'Excel-Datei',
            'vorlagen_ordner': 'Vorlagen-Ordner',
            'bilder_ordner': 'Bilder-Ordner (optional)',
            'export_ordner': 'Export-Ordner'
        }
        path_placeholders = {
            'excel_path': 'Pfad zur Master-Excel (.xlsx / .xls)',
            'vorlagen_ordner': 'Ordner mit Unterordnern …\\anlagen\\ und …\\allgemein\\',
            'bilder_ordner': 'Ordner mit Bildern für *_img-Platzhalter (optional)',
            'export_ordner': 'Zielordner für erzeugte Dokumente',
        }
        for key, name in path_map.items():
            row = QHBoxLayout()
            row.addWidget(QLabel(f"{name}:"))
            path_edit = QLineEdit(self.paths.get(key, ''))
            path_edit.setPlaceholderText(path_placeholders.get(key, ''))
            path_edit.setClearButtonEnabled(True)
            path_edit.editingFinished.connect(lambda k=key, p=path_edit: self.on_path_change(k, p.text()))
            path_edit.textChanged.connect(lambda text, widget=path_edit: self.validate_path(widget))
            setattr(self, f"{key}_edit", path_edit)
            row.addWidget(path_edit)
            browse_btn = QPushButton("...")
            browse_btn.clicked.connect(lambda c, k=key: self.browse(k))
            row.addWidget(browse_btn)
            if key == 'export_ordner':
                layout.addLayout(row)
                proj_row = QHBoxLayout()
                proj_row.addWidget(QLabel("Projektname (Export-Ordner):"))
                self.projekt_name_override_edit = QLineEdit(self.settings.get('projekt_name_override', ''))
                self.projekt_name_override_edit.setPlaceholderText("Leer = aus Excel")
                self.projekt_name_override_edit.editingFinished.connect(self.save_all_settings)
                self.projekt_name_override_edit.textChanged.connect(lambda _t: self._update_excel_basename_display())
                proj_row.addWidget(self.projekt_name_override_edit)
                layout.addLayout(proj_row)
                continue
            # Excel-Icon-Button nur beim Excel-Pfad
            if key == 'excel_path':
                path_edit.textChanged.connect(lambda _t: self._update_excel_basename_display())
                excel_icon_btn = QPushButton()
                excel_icon_btn.setIcon(QIcon(resource_path('Pictures/excel_icon.png')))
                excel_icon_btn.setToolTip("Excel-Datei öffnen")
                excel_icon_btn.setFixedWidth(32)
                excel_icon_btn.clicked.connect(self.open_excel_file)
                row.addWidget(excel_icon_btn)
            layout.addLayout(row)

        settings_layout = QHBoxLayout()
        setting_map = {
            'header_row': ('Header-Zeile', (1, 100), 3),
            'svg_scale': ('SVG-Skala', (1, 10), 3),
            'png_compression': ('PNG-Komp. (0=Max)', (-1, 100), -1)
        }
        spin_tooltips = {
            'header_row': "Excel-Zeile mit den Spaltennamen (1 = erste Zeile der Datei).",
            'svg_scale': "Skalierung bei der SVG→PNG-Konvertierung für die Word-Ausgabe.",
            'png_compression': "0 = maximale PNG-Kompression, höher = weniger Kompression; -1 = Standard.",
        }
        for key, (name, r, default) in setting_map.items():
            settings_layout.addWidget(QLabel(name))
            spin = QSpinBox()
            spin.setRange(*r)
            spin.setValue(self.settings.get(key, default))
            spin.setToolTip(spin_tooltips.get(key, ""))
            spin.valueChanged.connect(self.save_all_settings)
            setattr(self, f"{key}_spin", spin)
            settings_layout.addWidget(spin)
        layout.addLayout(settings_layout)

        # Layout für UTC-Format
        utc_format_layout = QHBoxLayout()
        utc_format_layout.addWidget(QLabel("Zeitstempel-Format (datetime_utc):"))
        self.datetime_utc_format_edit = QLineEdit(self.settings.get('datetime_utc_format'))
        self.datetime_utc_format_edit.editingFinished.connect(self.save_all_settings)
        utc_format_layout.addWidget(self.datetime_utc_format_edit)
        layout.addLayout(utc_format_layout)

        project_root_layout = QHBoxLayout()
        project_root_layout.addWidget(QLabel("Projektverzeichnis (zentral):"))
        self.project_root_dir_edit = QLineEdit(self.settings.get('project_root_dir', ''))
        self.project_root_dir_edit.setPlaceholderText("Optional: Standardordner für *.dta.json")
        self.project_root_dir_edit.editingFinished.connect(self.on_project_root_dir_change)
        self.project_root_dir_edit.textChanged.connect(lambda _text: self.validate_path(self.project_root_dir_edit))
        project_root_layout.addWidget(self.project_root_dir_edit)
        project_root_browse_btn = QPushButton("...")
        project_root_browse_btn.setToolTip("Zentrales Projektverzeichnis auswählen")
        project_root_browse_btn.clicked.connect(self.browse_project_root_dir)
        project_root_layout.addWidget(project_root_browse_btn)
        layout.addLayout(project_root_layout)
        self.validate_path(self.project_root_dir_edit)

        # Versions-Info und Pfad-Normalisierung
        version_layout = QHBoxLayout()
        version_layout.addWidget(QLabel(f"Version: {VERSION}"))
        version_layout.addWidget(QLabel(f"Build: {BUILD_DATE}"))
        version_layout.addStretch()
        # Button: Pfade normalisieren für Team-Portabilität
        normalize_btn = QPushButton("🔄 Pfade normalisieren")
        normalize_btn.setToolTip(
            f"Wandelt den lokalen Benutzerordner unter …\\Users\\… in '{{username}}' um,\n"
            "damit Einstellungen/Projekte auf allen Rechnern geteilt werden können."
        )
        normalize_btn.clicked.connect(self.normalize_all_paths)
        version_layout.addWidget(normalize_btn)
        denorm_btn = QPushButton("Lokalen Benutzer einsetzen")
        denorm_btn.setToolTip(
            f"Ersetzt '{{username}}' in allen Pfaden durch '{self.current_username}' "
            "(nach Import oder „Pfade normalisieren“)."
        )
        denorm_btn.clicked.connect(self.denormalize_all_paths)
        version_layout.addWidget(denorm_btn)
        layout.addLayout(version_layout)

        layout.addWidget(QFrame(frameShape=QFrame.HLine))

        self.export_as_pdf_check = QCheckBox("Word-Dokumente nach Generierung in PDF umwandeln")
        self.export_as_pdf_check.setChecked(self.settings.get('export_as_pdf', False))
        self.export_as_pdf_check.stateChanged.connect(self.save_all_settings)
        layout.addWidget(self.export_as_pdf_check)

        self.parallel_doc_generation_check = QCheckBox("Word-Erzeugung parallelisieren (mehrere Threads, experimentell)")
        self.parallel_doc_generation_check.setChecked(
            self.settings.get('parallel_doc_generation', False)
        )
        self.parallel_doc_generation_check.setToolTip(
            "Erhöht die Parallelität bei docxtpl und Lageplan-Kopien. "
            "Kann schneller sein, setzt aber mehr Last auf Word/CPU. Bei Problemen deaktivieren."
        )
        self.parallel_doc_generation_check.stateChanged.connect(self.save_all_settings)
        layout.addWidget(self.parallel_doc_generation_check)

        workflow_frame = QFrame()
        workflow_frame.setFrameShape(QFrame.StyledPanel)
        workflow_layout = QVBoxLayout(workflow_frame)
        workflow_title = QLabel("Workflow-Status")
        workflow_title.setStyleSheet("font-weight: 600;")
        workflow_layout.addWidget(workflow_title)
        self.workflow_traffic_label = QLabel("Status: Nicht bereit")
        self.workflow_traffic_label.setStyleSheet(
            "padding: 4px 8px; border-radius: 4px; background: #f8d7da; color: #721c24; font-weight: 600;"
        )
        workflow_layout.addWidget(self.workflow_traffic_label)
        self.workflow_status_label = QLabel("")
        self.workflow_status_label.setWordWrap(True)
        workflow_layout.addWidget(self.workflow_status_label)
        layout.addWidget(workflow_frame)

        quick_actions_row = QHBoxLayout()
        quick_actions_row.addWidget(QLabel("Projekt-Quick-Actions:"))
        quick_new_btn = QPushButton("Neu")
        quick_new_btn.clicked.connect(self.project_new)
        quick_actions_row.addWidget(quick_new_btn)
        quick_open_btn = QPushButton("Öffnen...")
        quick_open_btn.clicked.connect(self.project_open)
        quick_actions_row.addWidget(quick_open_btn)
        quick_save_btn = QPushButton("Speichern")
        quick_save_btn.clicked.connect(self.project_save)
        quick_actions_row.addWidget(quick_save_btn)
        quick_actions_row.addStretch()
        layout.addLayout(quick_actions_row)

        btn_row = QHBoxLayout()
        self.dry_run_btn = QPushButton("Konfiguration prüfen")
        self.preview_btn = QPushButton("Vorschau")
        self.preview_btn.setToolTip(
            "Erzeugt pro ausgewählter Vorlage ein Beispieldokument mit der ersten Excel-Datenzeile "
            "(Ordner …_Vorschau unter dem Export-Pfad)."
        )
        self.start_btn = QPushButton("Start")
        self.close_btn = QPushButton("Schließen")

        self.dry_run_btn.clicked.connect(self.start_dry_run)
        self.preview_btn.clicked.connect(self.start_preview_batch)
        self.start_btn.clicked.connect(self.start)
        self.close_btn.clicked.connect(self.handle_close_or_cancel)

        btn_row.addWidget(self.dry_run_btn)
        btn_row.addWidget(self.preview_btn)
        btn_row.addWidget(self.start_btn)
        btn_row.addWidget(self.close_btn)
        layout.addLayout(btn_row)

        self._update_excel_basename_display()
        self.update_workflow_status()

    def update_workflow_status(self):
        """Zeigt den aktuellen Stand der Pflichtschritte bis zum Start an."""
        if not hasattr(self, "workflow_status_label"):
            return

        missing_for_start = [
            k for k in ["excel_path", "vorlagen_ordner", "export_ordner"]
            if not (self.paths.get(k) or "").strip()
        ]
        missing_for_dry_run = [
            k for k in ["excel_path", "vorlagen_ordner"]
            if not (self.paths.get(k) or "").strip()
        ]

        templates_total = 0
        templates_selected = 0
        if hasattr(self, "_template_checkbox_by_rel") and self._template_checkbox_by_rel:
            templates_total = len(self._template_checkbox_by_rel)
            templates_selected = sum(1 for cb in self._template_checkbox_by_rel.values() if cb.isChecked())
        has_template_selection = templates_total > 0 and templates_selected > 0

        is_running = hasattr(self, "worker") and self.worker is not None and self.worker.isRunning()
        can_start = (not missing_for_start) and has_template_selection
        can_dry_run = (not missing_for_dry_run) and has_template_selection
        can_preview = can_start

        status_parts = []
        if missing_for_start:
            status_parts.append(f"1) Fehlende Pflichtpfade: {len(missing_for_start)}")
        else:
            status_parts.append("1) Pflichtpfade: OK")

        if templates_total == 0:
            status_parts.append("2) Vorlagen: keine gefunden")
        elif templates_selected == 0:
            status_parts.append(f"2) Vorlagen: 0/{templates_total} gewählt")
        else:
            status_parts.append(f"2) Vorlagen: {templates_selected}/{templates_total} gewählt")

        status_parts.append("3) Start bereit: Ja" if can_start else "3) Start bereit: Nein")

        self.workflow_status_label.setText(" | ".join(status_parts))
        if hasattr(self, "workflow_traffic_label"):
            if is_running:
                self.workflow_traffic_label.setText("Status: Lauf aktiv")
                self.workflow_traffic_label.setStyleSheet(
                    "padding: 4px 8px; border-radius: 4px; background: #fff3cd; color: #856404; font-weight: 600;"
                )
            elif can_start:
                self.workflow_traffic_label.setText("Status: Bereit")
                self.workflow_traffic_label.setStyleSheet(
                    "padding: 4px 8px; border-radius: 4px; background: #d4edda; color: #155724; font-weight: 600;"
                )
            else:
                self.workflow_traffic_label.setText("Status: Nicht bereit")
                self.workflow_traffic_label.setStyleSheet(
                    "padding: 4px 8px; border-radius: 4px; background: #f8d7da; color: #721c24; font-weight: 600;"
                )

        # Buttons nur freigeben, wenn Voraussetzungen erfüllt sind und kein Lauf aktiv ist.
        if hasattr(self, "start_btn"):
            self.start_btn.setEnabled((not is_running) and can_start)
        if hasattr(self, "preview_btn"):
            self.preview_btn.setEnabled((not is_running) and can_preview)
        if hasattr(self, "dry_run_btn"):
            self.dry_run_btn.setEnabled((not is_running) and can_dry_run)

    def _update_excel_basename_display(self):
        """Setzt die obere Zeile primär auf den Projektnamen (Fallback: Excel-Dateiname)."""
        if not hasattr(self, "excel_basename_line"):
            return
        p = (self.excel_path_edit.text().strip() if hasattr(self, "excel_path_edit") else "")
        if not p and hasattr(self, "paths"):
            p = (self.paths.get("excel_path") or "").strip()
        project_name = ""
        if hasattr(self, "projekt_name_override_edit"):
            project_name = self.projekt_name_override_edit.text().strip()
        if not project_name and hasattr(self, "settings"):
            project_name = (self.settings.get("projekt_name_override") or "").strip()
        if not project_name and hasattr(self, "_read_name_from_excel"):
            project_name = (self._read_name_from_excel() or "").strip()
        if project_name and hasattr(self, "_base_name_for_project_dta_file"):
            project_name = self._base_name_for_project_dta_file(project_name) or project_name

        if project_name:
            self.excel_basename_line.setText(project_name)
            tip = "Projektname"
            if p:
                tip += f"\nExcel: {p}"
            self.excel_basename_line.setToolTip(tip)
            return

        if p:
            self.excel_basename_line.setText(os.path.basename(p))
            self.excel_basename_line.setToolTip(f"Kein Projektname gefunden.\nExcel: {p}")
            return

        self.excel_basename_line.clear()
        self.excel_basename_line.setToolTip("Kein Projektname gefunden.")

    def _on_excel_ansicht_zurueckschreiben(self):
        """Aktualisiert die Tabelle aus der Datei (nach Bearbeitung in Excel)."""
        p = (self.excel_path_edit.text().strip() if hasattr(self, "excel_path_edit") else "")
        if not p or not os.path.isfile(p):
            QMessageBox.warning(
                self,
                "Excel-Ansicht",
                "Bitte zuerst eine vorhandene Excel-Datei unter „Excel-Datei“ wählen.",
            )
            return
        self.show_excel_data()
        if hasattr(self, "current_file_label"):
            self.current_file_label.setText("Ansicht aus Excel wurde aktualisiert.")
