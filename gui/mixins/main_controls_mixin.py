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
    QTableView,
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
        self.excel_basename_line.setPlaceholderText("Keine Excel-Datei gewählt")
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
        for key, name in path_map.items():
            row = QHBoxLayout()
            row.addWidget(QLabel(f"{name}:"))
            path_edit = QLineEdit(self.paths.get(key, ''))
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
        for key, (name, r, default) in setting_map.items():
            settings_layout.addWidget(QLabel(name))
            spin = QSpinBox()
            spin.setRange(*r)
            spin.setValue(self.settings.get(key, default))
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
        # Bereich Tabelle: Ansicht-Überschrift + Inhalt aus Datei in die Tabelle „zurücklesen“
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
        layout.addWidget(self.table)
        self.progress = QProgressBar()
        layout.addWidget(self.progress)
        self.current_file_label = QLabel("Bereit zum Starten...")
        self.current_file_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.current_file_label)

        self.open_export_folder_btn = QPushButton("Export-Ordner öffnen")
        self.open_export_folder_btn.clicked.connect(self.open_export_folder)
        self.open_export_folder_btn.setVisible(False)
        layout.addWidget(self.open_export_folder_btn)

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

    def _update_excel_basename_display(self):
        """Setzt die obere Zeile auf den Dateinamen der gewählten Excel-Datei (Tooltip = voller Pfad)."""
        if not hasattr(self, "excel_basename_line"):
            return
        p = (self.excel_path_edit.text().strip() if hasattr(self, "excel_path_edit") else "")
        if not p and hasattr(self, "paths"):
            p = (self.paths.get("excel_path") or "").strip()
        if p:
            self.excel_basename_line.setText(os.path.basename(p))
            self.excel_basename_line.setToolTip(p)
        else:
            self.excel_basename_line.clear()
            self.excel_basename_line.setToolTip("")

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
