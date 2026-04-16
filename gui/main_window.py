import json
import os
import sys
import traceback

import pandas as pd
from datetime import datetime

from PySide6.QtCore import Qt, QAbstractTableModel, QThread, Signal
from PySide6.QtGui import QIcon, QAction, QActionGroup, QTextCursor, QPixmap
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton,
    QFileDialog, QProgressBar, QTableView, QMessageBox, QFrame,
    QSpinBox, QTextEdit, QMainWindow, QTabWidget, QGroupBox, QScrollArea, QCheckBox,
    QMenu, QComboBox, QTableWidget, QTableWidgetItem, QSplitter
)

from core.logic import verarbeite_vorlagen
from core.logging import _log_handler
from core.utils import resource_path


# ==============================================================================
# VERSION
# ==============================================================================
VERSION = "7.0.1"
BUILD_DATE = "2025-01-27"

# Unterstützte Bildformate für den Bilder-Vorschau-Tab
BILDER_VORSCHAU_EXTENSIONS = {'.png', '.jpg', '.jpeg', '.gif', '.svg', '.bmp'}

# Maximale Anzahl Einträge im Untermenü "Letzte Projekte"
MAX_RECENT_PROJECTS = 6


class PandasModel(QAbstractTableModel):
    """
    Ein benutzerdefiniertes Tabellenmodell für PySide6, das einen
    Pandas DataFrame als Datenquelle für eine QTableView verwenden kann.
    Dies ermöglicht die Anzeige der Excel-Daten direkt in der GUI.
    """
    def __init__(self, df):
        super().__init__()
        self._df = df

    def rowCount(self, parent=None):
        return self._df.shape[0]

    def columnCount(self, parent=None):
        return self._df.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid() and role == Qt.DisplayRole:
            return str(self._df.iloc[index.row(), index.column()])

    def headerData(self, col, orient, role=Qt.DisplayRole):
        if orient == Qt.Horizontal and role == Qt.DisplayRole:
            return self._df.columns[col]


class Worker(QThread):
    """
    Ein QThread-Worker, der die zeitintensive Dokumentenerstellung in einem
    separaten Thread ausführt. Dies verhindert, dass die GUI während des
    Prozesses "einfriert" und nicht mehr reagiert.

    Signale:
    - log: Sendet Log-Nachrichten an die GUI.
    - progress: Sendet den Fortschritt (aktueller Wert, Maximalwert).
    - finished: Signalisiert das Ende der Verarbeitung.
    - current_file: Sendet den Namen der aktuell bearbeiteten Datei.
    """
    log = Signal(str)  # html_message
    progress = Signal(int, int)
    finished = Signal(bool, str)  # success, export_path
    current_file = Signal(str)

    def __init__(self, **kwargs):
        super().__init__()
        self.params = kwargs
        self.dry_run = kwargs.get('dry_run', False)

    def run(self):
        """Startet die Verarbeitung durch Aufruf der `verarbeite_vorlagen` Funktion."""
        export_path = None
        try:
            thread_safe_params = self.params.copy()
            thread_safe_params['log_callback'] = self.log.emit
            thread_safe_params['progress_callback'] = self.progress.emit
            thread_safe_params['file_callback'] = self.current_file.emit
            thread_safe_params['worker_thread'] = self
            thread_safe_params.pop('theme', None)

            result = verarbeite_vorlagen(**thread_safe_params)

            if self.dry_run:
                self.finished.emit(result, None)  # Bei Trockenlauf ist das Ergebnis ein Boolean
            elif not self.isInterruptionRequested():
                self.finished.emit(bool(result), result)  # Bei echtem Lauf ist es der Pfad
            else:
                self.finished.emit(False, None)

        except Exception as e:
            error_msg = f"FATALER FEHLER im Worker-Thread: {e}\n{traceback.format_exc()}"
            _log_handler(error_msg, "FATAL", self.log.emit)
            self.finished.emit(False, None)


class MainWindow(QMainWindow):
    """
    Das Hauptfenster der Anwendung.
    Es initialisiert die Benutzeroberfläche, verwaltet die Benutzereingaben,
    lädt und speichert Einstellungen und startet den Worker-Thread.
    """
    def __init__(self):
        # #region agent log
        log_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), '.cursor', 'debug.log')
        try:
            with open(log_path, 'a', encoding='utf-8') as f:
                import json
                f.write(json.dumps({"sessionId":"debug-session","runId":"run1","hypothesisId":"C","location":"gui/main_window.py:109","message":"MainWindow.__init__ entry","data":{},"timestamp":int(__import__('time').time()*1000)}) + '\n')
        except: pass
        # #endregion
        super().__init__()
        self.setWindowTitle(f'DocxTpl Automatisierung v{VERSION} - Build {BUILD_DATE}')
        self.category_widgets = {}
        self.theme_action_group = None
        self.last_export_path = None
        # Aktuellen Windows-Benutzernamen ermitteln (für portable Pfade)
        self.current_username = os.environ.get('USERNAME', os.environ.get('USER', 'UnknownUser'))
        # #region agent log
        try:
            with open(log_path, 'a', encoding='utf-8') as f:
                import json
                f.write(json.dumps({"sessionId":"debug-session","runId":"run1","hypothesisId":"C","location":"gui/main_window.py:118","message":"Before load_settings_and_init_vars","data":{"username":self.current_username},"timestamp":int(__import__('time').time()*1000)}) + '\n')
        except: pass
        # #endregion
        self.load_settings_and_init_vars()
        # #region agent log
        try:
            with open(log_path, 'a', encoding='utf-8') as f:
                import json
                f.write(json.dumps({"sessionId":"debug-session","runId":"run1","hypothesisId":"B","location":"gui/main_window.py:120","message":"Before setup_ui","data":{"paths_keys":list(self.paths.keys())},"timestamp":int(__import__('time').time()*1000)}) + '\n')
        except: pass
        # #endregion
        self.setup_ui()
        # #region agent log
        try:
            with open(log_path, 'a', encoding='utf-8') as f:
                import json
                f.write(json.dumps({"sessionId":"debug-session","runId":"run1","hypothesisId":"B","location":"gui/main_window.py:122","message":"After setup_ui","data":{"has_bilder_ordner_edit":hasattr(self, 'bilder_ordner_edit')},"timestamp":int(__import__('time').time()*1000)}) + '\n')
        except: pass
        # #endregion

    def load_settings_and_init_vars(self):
        """Lädt die zuletzt verwendeten Einstellungen (Pfade, Optionen) aus einer JSON-Datei."""
        last = self.load_settings()
        self.paths = {
            'excel_path': last.get('excel_path', ''),
            'vorlagen_ordner': last.get('vorlagen_ordner', ''),
            'bilder_ordner': last.get('bilder_ordner', ''),
            'export_ordner': last.get('export_ordner', '')
        }
        self.settings = {
            'header_row': last.get('header_row', 3),
            'svg_scale': last.get('svg_scale', 3),
            'png_compression': last.get('png_compression', -1),
            'theme': last.get('theme', 'Light'),
            'datetime_utc_format': last.get('datetime_utc_format', '%Y-%m-%d %H:%M:%S UTC'),
            'projekt_name_override': last.get('projekt_name_override', ''),
            'export_as_pdf': last.get('export_as_pdf', False)
        }
        self.categories = last.get('categories', {'b_': 'Beschilderung', 'ba_': 'Betriebsanweisung'})
        # Projektverwaltung: zuletzt verwendete Projekte und aktiver Projektpfad
        self.recent_projects = last.get('recent_projects', [])
        self.current_project_path = last.get('last_project_path', None)

    def setup_ui(self):
        """Erstellt und arrangiert alle UI-Elemente im Hauptfenster."""
        self.resize(800, 700)
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        tab_widget = QTabWidget()
        main_layout.addWidget(tab_widget)
        self.tab_widget = tab_widget

        # Tab 1: Hauptsteuerung
        controls_tab = QWidget()
        tab_widget.addTab(controls_tab, "Hauptsteuerung")
        self.setup_controls_tab(controls_tab)

        # Tab 2: Kategorien
        categories_tab = QWidget()
        tab_widget.addTab(categories_tab, "Kategorien")
        self.setup_categories_tab(categories_tab)

        # Tab 3: Bilder-Vorschau
        bilder_vorschau_tab = QWidget()
        tab_widget.addTab(bilder_vorschau_tab, "Bilder-Vorschau")
        self.bilder_vorschau_tab_index = tab_widget.indexOf(bilder_vorschau_tab)
        self.setup_bilder_vorschau_tab(bilder_vorschau_tab)

        # Tab 4: Logs
        logs_tab = QWidget()
        tab_widget.addTab(logs_tab, "Logs")
        tab_widget.currentChanged.connect(self._on_tab_changed)
        self.setup_logs_tab(logs_tab)

        self.setup_menu()
        self.apply_theme(self.settings.get('theme', 'Light'))

        for key in self.paths:
            self.validate_path(getattr(self, f"{key}_edit"))

        self.show_excel_data()

        # Zeige Versions-Info in den Logs
        try:
            import PySide6
            pyside_version = PySide6.__version__
        except ImportError:
            pyside_version = "Nicht verfügbar"

        self.log_text.append(
            f'<div style="color: #666; text-align: center; padding: 10px; border: 1px solid #ccc; margin: 10px 0; background-color: #f9f9f9;">'
            f'<strong>DocxTpl Automatisierung v{VERSION}</strong><br>'
            f'Build: {BUILD_DATE}<br>'
            f'Python: {sys.version.split()[0]} | PySide6: {pyside_version}</div>'
        )

    def setup_controls_tab(self, tab):
        """Erstellt den Inhalt des 'Hauptsteuerung'-Tabs."""
        layout = QVBoxLayout(tab)
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
            f"Wandelt Benutzername '{self.current_username}' in '{{username}}' um,\n"
            "damit Projekte auf allen Rechnern funktionieren."
        )
        normalize_btn.clicked.connect(self.normalize_all_paths)
        version_layout.addWidget(normalize_btn)
        layout.addLayout(version_layout)

        layout.addWidget(QFrame(frameShape=QFrame.HLine))
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

        btn_row = QHBoxLayout()
        self.dry_run_btn = QPushButton("Konfiguration prüfen")
        self.start_btn = QPushButton("Start")
        self.close_btn = QPushButton("Schließen")

        self.dry_run_btn.clicked.connect(self.start_dry_run)
        self.start_btn.clicked.connect(self.start)
        self.close_btn.clicked.connect(self.handle_close_or_cancel)

        btn_row.addWidget(self.dry_run_btn)
        btn_row.addWidget(self.start_btn)
        btn_row.addWidget(self.close_btn)
        layout.addLayout(btn_row)

    def setup_categories_tab(self, tab):
        """Erstellt den Inhalt des 'Kategorien'-Tabs für die Verwaltung der Ausgabeordner."""
        layout = QVBoxLayout(tab)

        # Versions-Info
        version_info = QLabel(f"Version {VERSION} - Build {BUILD_DATE}")
        version_info.setStyleSheet("color: #666; font-size: 10px; padding: 5px;")
        version_info.setAlignment(Qt.AlignCenter)
        layout.addWidget(version_info)

        group_box = QGroupBox("Dokumentkategorien")
        layout.addWidget(group_box)
        group_layout = QVBoxLayout(group_box)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_content = QWidget()
        self.categories_ui_layout = QVBoxLayout(scroll_content)
        scroll.setWidget(scroll_content)

        for prefix, name in self.categories.items():
            self.add_category_widget(prefix, name)

        btn_layout = QHBoxLayout()
        add_btn = QPushButton("+")
        add_btn.setFixedWidth(30)
        add_btn.setToolTip("Neue Kategorie hinzufügen")
        add_btn.clicked.connect(self.add_category)
        remove_btn = QPushButton("-")
        remove_btn.setFixedWidth(30)
        remove_btn.setToolTip("Ausgewählte Kategorie entfernen")
        remove_btn.clicked.connect(self.remove_category)
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(remove_btn)
        btn_layout.addStretch()

        group_layout.addLayout(btn_layout)
        group_layout.addWidget(scroll)

        save_btn = QPushButton("Kategorien speichern")
        save_btn.clicked.connect(self.save_categories)
        layout.addWidget(save_btn)
        layout.addStretch()

    def setup_bilder_vorschau_tab(self, tab):
        """Erstellt den Inhalt des 'Bilder-Vorschau'-Tabs: Dateiliste des Bilder-Ordners und Bildvorschau."""
        layout = QVBoxLayout(tab)

        self.bilder_vorschau_hint_label = QLabel("Bitte in der Hauptsteuerung einen Bilder-Ordner wählen.")
        self.bilder_vorschau_hint_label.setStyleSheet("color: #666; padding: 8px;")
        layout.addWidget(self.bilder_vorschau_hint_label)

        splitter = QSplitter(Qt.Horizontal)

        self.bilder_vorschau_table = QTableWidget()
        self.bilder_vorschau_table.setColumnCount(2)
        self.bilder_vorschau_table.setHorizontalHeaderLabels(["Dateiname", "Dateityp"])
        self.bilder_vorschau_table.horizontalHeader().setStretchLastSection(True)
        self.bilder_vorschau_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.bilder_vorschau_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.bilder_vorschau_table.itemSelectionChanged.connect(self._on_bilder_vorschau_selection_changed)
        splitter.addWidget(self.bilder_vorschau_table)

        preview_frame = QFrame()
        preview_frame.setFrameShape(QFrame.StyledPanel)
        preview_layout = QVBoxLayout(preview_frame)
        preview_layout.addWidget(QLabel("Vorschau:"))
        self.bilder_vorschau_preview_label = QLabel("Keine Vorschau")
        self.bilder_vorschau_preview_label.setAlignment(Qt.AlignCenter)
        self.bilder_vorschau_preview_label.setMinimumSize(200, 200)
        self.bilder_vorschau_preview_label.setMaximumSize(400, 400)
        self.bilder_vorschau_preview_label.setStyleSheet("border: 1px solid #ccc; background: #f5f5f5; padding: 8px;")
        self.bilder_vorschau_preview_label.setScaledContents(False)
        preview_layout.addWidget(self.bilder_vorschau_preview_label)
        preview_layout.addStretch()
        splitter.addWidget(preview_frame)

        splitter.setSizes([400, 400])
        layout.addWidget(splitter)

    def _on_tab_changed(self, index):
        """Lädt die Bilder-Vorschau-Liste neu, wenn der Bilder-Vorschau-Tab aktiv wird."""
        if hasattr(self, 'bilder_vorschau_tab_index') and index == self.bilder_vorschau_tab_index:
            self.refresh_bilder_vorschau()

    def refresh_bilder_vorschau(self):
        """Liest den Bilder-Ordner ein und füllt die Tabelle; setzt die Vorschau zurück."""
        if not hasattr(self, 'bilder_vorschau_table'):
            return
        path = (self.bilder_ordner_edit.text() if hasattr(self, 'bilder_ordner_edit') else "").strip()
        self.bilder_vorschau_table.setRowCount(0)
        self.bilder_vorschau_preview_label.clear()
        self.bilder_vorschau_preview_label.setText("Keine Vorschau")
        if not path:
            self.bilder_vorschau_hint_label.setVisible(True)
            self.bilder_vorschau_hint_label.setText("Bitte in der Hauptsteuerung einen Bilder-Ordner wählen.")
            return
        if not os.path.isdir(path):
            self.bilder_vorschau_hint_label.setVisible(True)
            self.bilder_vorschau_hint_label.setText(f"Ordner existiert nicht: {path}")
            return
        self.bilder_vorschau_hint_label.setVisible(False)
        try:
            names = sorted(os.listdir(path))
        except OSError:
            self.bilder_vorschau_hint_label.setVisible(True)
            self.bilder_vorschau_hint_label.setText("Ordner konnte nicht gelesen werden.")
            return
        for name in names:
            full = os.path.join(path, name)
            if os.path.isfile(full):
                ext = os.path.splitext(name)[1].lower()
                if ext in BILDER_VORSCHAU_EXTENSIONS:
                    row = self.bilder_vorschau_table.rowCount()
                    self.bilder_vorschau_table.insertRow(row)
                    self.bilder_vorschau_table.setItem(row, 0, QTableWidgetItem(name))
                    self.bilder_vorschau_table.setItem(row, 1, QTableWidgetItem(ext or "—"))
                    self.bilder_vorschau_table.item(row, 0).setData(Qt.UserRole, full)
                    self.bilder_vorschau_table.item(row, 1).setData(Qt.UserRole, full)

    def _on_bilder_vorschau_selection_changed(self):
        """Zeigt bei Auswahl einer Zeile die Bildvorschau an."""
        if not hasattr(self, 'bilder_vorschau_preview_label'):
            return
        row = self.bilder_vorschau_table.currentRow()
        if row < 0:
            self.bilder_vorschau_preview_label.clear()
            self.bilder_vorschau_preview_label.setText("Keine Vorschau")
            return
        item = self.bilder_vorschau_table.item(row, 0)
        file_path = item.data(Qt.UserRole) if item else None
        if not file_path or not os.path.isfile(file_path):
            self.bilder_vorschau_preview_label.clear()
            self.bilder_vorschau_preview_label.setText("Keine Vorschau")
            return
        ext = os.path.splitext(file_path)[1].lower()
        if ext == '.svg':
            try:
                from core.utils import svg_to_png_file_pyside
                import tempfile
                fd, tmp_png = tempfile.mkstemp(suffix='.png')
                os.close(fd)
                try:
                    if svg_to_png_file_pyside(file_path, tmp_png, scale=2, compression=-1):
                        pix = QPixmap(tmp_png)
                        if not pix.isNull():
                            scaled = pix.scaled(400, 400, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                            self.bilder_vorschau_preview_label.setPixmap(scaled)
                            self.bilder_vorschau_preview_label.setText("")
                        else:
                            self.bilder_vorschau_preview_label.clear()
                            self.bilder_vorschau_preview_label.setText("SVG konnte nicht geladen werden.")
                    else:
                        self.bilder_vorschau_preview_label.clear()
                        self.bilder_vorschau_preview_label.setText("Keine Vorschau")
                finally:
                    try:
                        os.unlink(tmp_png)
                    except OSError:
                        pass
            except Exception:
                self.bilder_vorschau_preview_label.clear()
                self.bilder_vorschau_preview_label.setText("Keine Vorschau")
        else:
            pix = QPixmap(file_path)
            if not pix.isNull():
                scaled = pix.scaled(400, 400, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                self.bilder_vorschau_preview_label.setPixmap(scaled)
                self.bilder_vorschau_preview_label.setText("")
            else:
                self.bilder_vorschau_preview_label.clear()
                self.bilder_vorschau_preview_label.setText("Keine Vorschau")

    def setup_logs_tab(self, tab):
        """Erstellt den Inhalt des 'Logs'-Tabs mit Filter-Dropdown."""
        layout = QVBoxLayout(tab)

        # Versions-Info (immer sichtbar)
        version_info = QLabel(f"DocxTpl Automatisierung v{VERSION} - Build {BUILD_DATE}")
        version_info.setStyleSheet("color: #666; font-size: 10px; padding: 5px;")
        version_info.setAlignment(Qt.AlignCenter)
        layout.addWidget(version_info)

        # Filter-Dropdown
        self.log_filter_combo = QComboBox()
        self.log_filter_combo.addItems([
            "ALLE", "INFO", "WARN", "ERROR", "SUCCESS", "FATAL", "SEP"
        ])
        self.log_filter_combo.currentTextChanged.connect(self.apply_log_filter)
        layout.addWidget(self.log_filter_combo)
        # Log-Textfeld
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setContextMenuPolicy(Qt.CustomContextMenu)
        self.log_text.customContextMenuRequested.connect(self.show_log_context_menu)
        self.log_lines = []  # Speichert alle Log-Zeilen als (level, html) Tupel
        layout.addWidget(self.log_text)

    def setup_menu(self):
        """Erstellt die Menüleiste für die Anwendung, inkl. Projekt- und Theme-Menüs."""
        menu_bar = self.menuBar()
        # Projekt-Menü
        project_menu = menu_bar.addMenu("Projekt")
        act_new = QAction("Neu", self)
        act_open = QAction("Öffnen...", self)
        act_save = QAction("Speichern", self)
        act_save_as = QAction("Speichern unter...", self)
        project_menu.addAction(act_new)
        project_menu.addAction(act_open)
        self.recent_projects_menu = project_menu.addMenu("Letzte Projekte")
        self.recent_projects_menu.aboutToShow.connect(self._refresh_recent_projects_menu)
        project_menu.addSeparator()
        project_menu.addAction(act_save)
        project_menu.addAction(act_save_as)

        act_new.triggered.connect(self.project_new)
        act_open.triggered.connect(self.project_open)
        act_save.triggered.connect(self.project_save)
        act_save_as.triggered.connect(self.project_save_as)

        self._refresh_recent_projects_menu()

        # Ansicht/Theme
        view_menu = menu_bar.addMenu("Ansicht")
        theme_menu = view_menu.addMenu("Theme")

        self.theme_action_group = QActionGroup(self)
        self.theme_action_group.setExclusive(True)

        actions = {
            "Light": QAction("Hell (Standard)", self, checkable=True),
            "Dark": QAction("Dark Mode", self, checkable=True),
            "Girly": QAction("Girly Mode", self, checkable=True)
        }

        for name, action in actions.items():
            action.setData(name)
            self.theme_action_group.addAction(action)
            theme_menu.addAction(action)

        self.theme_action_group.triggered.connect(self.on_theme_selected)

    # -----------------------------
    # Projektverwaltung - Methoden
    # -----------------------------
    def _normalize_path(self, path):
        """
        Normalisiert einen Pfad für portables Speichern:
        Ersetzt JEDEN Benutzernamen nach C:/Users/ durch {username}.
        Beispiel: C:/Users/BeyzaAygündüz/... → C:/Users/{username}/...
        """
        if not path:
            return path

        import re
        # Beide Slash-Richtungen: C:\Users\<beliebiger_name>\ und C:/Users/<beliebiger_name>/
        # Regex erkennt alles zwischen Users\ und dem nächsten \ oder /
        normalized = re.sub(r'C:\\Users\\[^\\]+', r'C:\\Users\\{username}', path)
        normalized = re.sub(r'C:/Users/[^/]+', r'C:/Users/{username}', normalized)
        return normalized

    def _denormalize_path(self, path):
        """
        Denormalisiert einen Pfad beim Laden:
        Ersetzt {username} durch den aktuellen Benutzernamen.
        Beispiel: C:/Users/{username}/... → C:/Users/Jonas/...
        """
        if not path:
            return path
        # Beide Slash-Richtungen unterstützen
        denormalized = path.replace('C:\\Users\\{username}', f'C:\\Users\\{self.current_username}')
        denormalized = denormalized.replace('C:/Users/{username}', f'C:/Users/{self.current_username}')
        return denormalized

    def normalize_all_paths(self):
        """
        Normalisiert alle aktuellen Pfade in der UI (Button-Callback).
        Zeigt danach eine Bestätigung an.
        """
        changed_count = 0
        for key, value in self.paths.items():
            if value and '{username}' not in value:  # Nur normalisieren wenn noch kein {username} drin steht
                normalized = self._normalize_path(value)
                self.paths[key] = normalized
                if hasattr(self, f"{key}_edit"):
                    getattr(self, f"{key}_edit").setText(normalized)
                changed_count += 1

        if changed_count > 0:
            self.save_all_settings()
            QMessageBox.information(
                self,
                "Pfade normalisiert",
                f"{changed_count} Pfade wurden normalisiert (Benutzername → '{{username}}').\n\n"
                "Die Pfade funktionieren jetzt auf allen Rechnern!"
            )
        else:
            QMessageBox.information(
                self,
                "Bereits normalisiert",
                "Alle Pfade sind bereits normalisiert (enthalten bereits '{username}')."
            )

    def _collect_project_dict(self):
        """Sammelt den aktuellen Zustand für ein Projekt-JSON."""
        # Stelle sicher, dass Settings aus Spins/Textfeldern aktuell sind
        for key in self.settings:
            if hasattr(self, f"{key}_spin"):
                self.settings[key] = getattr(self, f"{key}_spin").value()
        if hasattr(self, 'datetime_utc_format_edit'):
            self.settings['datetime_utc_format'] = self.datetime_utc_format_edit.text()
        if hasattr(self, 'projekt_name_override_edit'):
            self.settings['projekt_name_override'] = self.projekt_name_override_edit.text().strip()
        if hasattr(self, 'export_as_pdf_check'):
            self.settings['export_as_pdf'] = self.export_as_pdf_check.isChecked()

        # Pfade normalisieren (Benutzername → {username}) für Portabilität
        normalized_paths = {k: self._normalize_path(v) for k, v in self.paths.items()}

        return {
            'version': VERSION,
            'paths': normalized_paths,
            'settings': self.settings,
            'categories': self.categories,
            'saved_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }

    def _apply_project_dict(self, data):
        """Wendet ein geladenes Projekt auf UI und Zustand an."""
        # Pfade denormalisieren ({username} → aktueller Benutzername)
        loaded_paths = data.get('paths', self.paths)
        self.paths = {k: self._denormalize_path(v) for k, v in loaded_paths.items()}

        self.settings = {**self.settings, **data.get('settings', {})}
        self.categories = data.get('categories', self.categories)

        # UI aktualisieren: Pfade
        for key, value in self.paths.items():
            if hasattr(self, f"{key}_edit"):
                getattr(self, f"{key}_edit").setText(value)
                self.validate_path(getattr(self, f"{key}_edit"))
        # Settings (Spins)
        for key in ['header_row', 'svg_scale', 'png_compression']:
            if hasattr(self, f"{key}_spin"):
                getattr(self, f"{key}_spin").setValue(
                    self.settings.get(key, getattr(self, f"{key}_spin").value())
                )
        # UTC Format
        if hasattr(self, 'datetime_utc_format_edit'):
            self.datetime_utc_format_edit.setText(
                self.settings.get('datetime_utc_format', self.datetime_utc_format_edit.text())
            )
        # Projektname (Export-Ordner)
        if hasattr(self, 'projekt_name_override_edit'):
            self.projekt_name_override_edit.setText(self.settings.get('projekt_name_override', ''))
        # PDF-Export
        if hasattr(self, 'export_as_pdf_check'):
            self.export_as_pdf_check.setChecked(self.settings.get('export_as_pdf', False))
        # Theme
        self.apply_theme(self.settings.get('theme', 'Light'))
        # Kategorien-UI neu aufbauen
        if hasattr(self, 'categories_ui_layout'):
            # Bestehende Widgets entfernen
            for layout_item, _triplet in list(self.category_widgets.items()):
                while layout_item.count():
                    child = layout_item.takeAt(0)
                    if child.widget():
                        child.widget().deleteLater()
                self.categories_ui_layout.removeItem(layout_item)
                layout_item.deleteLater()
                del self.category_widgets[layout_item]
            # Neu hinzufügen
            for prefix, name in self.categories.items():
                self.add_category_widget(prefix, name)
        # Excel-Vorschau aktualisieren
        self.show_excel_data()
        # Einstellungen persistieren
        self.save_all_settings()

    def _remember_recent_project(self, file_path):
        """Merkt sich ein Projekt in der Liste zuletzt verwendeter Projekte."""
        try:
            if not file_path:
                return
            if file_path in self.recent_projects:
                self.recent_projects.remove(file_path)
            self.recent_projects.insert(0, file_path)
            self.recent_projects = self.recent_projects[:10]
            self.current_project_path = file_path
            # in settings.json speichern
            self.save_all_settings()
        except Exception:
            pass

    def _refresh_recent_projects_menu(self):
        """Aktualisiert das Untermenü 'Letzte Projekte' mit den zuletzt geöffneten Projekten."""
        self.recent_projects_menu.clear()
        recent = getattr(self, 'recent_projects', [])[:MAX_RECENT_PROJECTS]
        for path in recent:
            if not os.path.isfile(path):
                continue
            display_name = os.path.basename(path)
            action = QAction(display_name, self)
            action.setToolTip(path)
            action.triggered.connect(lambda checked, p=path: self.project_open_path(p))
            self.recent_projects_menu.addAction(action)
        if not self.recent_projects_menu.actions():
            none_action = QAction("(Keine)", self)
            none_action.setEnabled(False)
            self.recent_projects_menu.addAction(none_action)

    def project_new(self):
        """Setzt ein leeres Projekt (ohne Pfade), behält aber Kategorien/Settings-Grundwerte bei."""
        self.paths = {k: '' for k in self.paths.keys()}
        for key in ['header_row', 'svg_scale', 'png_compression']:
            if hasattr(self, f"{key}_spin"):
                getattr(self, f"{key}_spin").setValue(
                    self.settings.get(key, getattr(self, f"{key}_spin").minimum())
                )
        if hasattr(self, 'datetime_utc_format_edit'):
            self.datetime_utc_format_edit.setText('%Y-%m-%d %H:%M:%S UTC')
        self.apply_theme('Light')
        for key in self.paths:
            if hasattr(self, f"{key}_edit"):
                getattr(self, f"{key}_edit").setText('')
                self.validate_path(getattr(self, f"{key}_edit"))
        self.current_project_path = None
        self.save_all_settings()

    def project_open_path(self, file_path):
        """Lädt ein Projekt aus dem angegebenen Pfad (ohne Dateidialog)."""
        if not file_path or not os.path.isfile(file_path):
            if file_path:
                QMessageBox.warning(self, "Fehler", f"Projektdatei nicht gefunden:\n{file_path}")
            return
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            # Erlaube beide Formate: altes Flat-Settings.json oder neues Projektformat
            if 'paths' in data and 'settings' in data:
                self._apply_project_dict(data)
            else:
                # Altes Format: direkt anwenden
                self._apply_project_dict({
                    'paths': {k: data.get(k, '') for k in self.paths.keys()},
                    'settings': {k: data.get(k, self.settings.get(k)) for k in self.settings.keys()},
                    'categories': data.get('categories', self.categories)
                })
            self._remember_recent_project(file_path)
        except Exception as e:
            QMessageBox.warning(self, "Fehler", f"Projekt konnte nicht geladen werden: {e}")

    def project_open(self):
        """Öffnet eine Projektdatei (*.dta.json) per Dateidialog."""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Projekt öffnen", os.getcwd(), "Projektdateien (*.dta.json *.json)"
        )
        if file_path:
            self.project_open_path(file_path)

    def project_save(self):
        """Speichert das aktuelle Projekt an den zuletzt genutzten Projektpfad oder fragt nach einem neuen Pfad."""
        if not self.current_project_path:
            return self.project_save_as()
        try:
            data = self._collect_project_dict()
            with open(self.current_project_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
            QMessageBox.information(self, "Gespeichert", f"Projekt gespeichert: {self.current_project_path}")
        except Exception as e:
            QMessageBox.warning(self, "Fehler", f"Projekt konnte nicht gespeichert werden: {e}")

    def project_save_as(self):
        """Speichert das aktuelle Projekt unter einem neuen Dateinamen (*.dta.json)."""
        default_name = 'projekt.dta.json'
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Projekt speichern unter",
            os.path.join(os.getcwd(), default_name),
            "Projektdateien (*.dta.json)"
        )
        if not file_path:
            return
        try:
            data = self._collect_project_dict()
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
            self._remember_recent_project(file_path)
            QMessageBox.information(self, "Gespeichert", f"Projekt gespeichert: {file_path}")
        except Exception as e:
            QMessageBox.warning(self, "Fehler", f"Projekt konnte nicht gespeichert werden: {e}")

    def on_theme_selected(self, action):
        """Wird aufgerufen, wenn ein Theme aus dem Menü ausgewählt wird."""
        theme_name = action.data()
        self.apply_theme(theme_name)
        self.settings['theme'] = theme_name
        self.save_all_settings()

    def apply_theme(self, theme_name):
        """Wendet das ausgewählte Stylesheet an und aktualisiert die Menü-Auswahl. Setzt im Dark Mode den Log-Text auf weiß."""
        stylesheet = self.get_stylesheet(theme_name)
        self.setStyleSheet(stylesheet)  # Apply to main window
        for action in self.theme_action_group.actions():
            if action.data() == theme_name:
                action.setChecked(True)
                break
        # Im Dark Mode: Log-Text maximal hell und kontrastreich
        self.is_dark_mode = (theme_name == 'Dark')
        if self.is_dark_mode:
            self.log_text.setStyleSheet("color: #FFF; background-color: #181818; font-weight: 500;")
        else:
            self.log_text.setStyleSheet("")

    def get_stylesheet(self, theme_name):
        """Gibt den QSS-Stylesheet-String für das angeforderte Theme zurück."""
        theme_map = {
            'Dark': 'dark.qss',
            'Girly': 'girly.qss'
        }
        qss_file = theme_map.get(theme_name)
        if qss_file:
            qss_path = resource_path(qss_file)
            if os.path.exists(qss_path):
                try:
                    with open(qss_path, 'r') as f:
                        return f.read()
                except IOError:
                    return ""  # Fallback to default
        return ""  # Light/Default Theme

    def on_path_change(self, key, value):
        """Wird aufgerufen, wenn ein Pfad geändert wird, und speichert die Einstellung."""
        self.paths[key] = value
        self.save_all_settings()

        # Aktualisiere Excel-Datenanzeige wenn sich der Excel-Pfad ändert
        if key == 'excel_path':
            self.show_excel_data()
        # Bilder-Vorschau-Tab aktualisieren wenn sich der Bilder-Ordner ändert
        if key == 'bilder_ordner' and hasattr(self, 'bilder_vorschau_table'):
            self.refresh_bilder_vorschau()

    def browse(self, key):
        """Öffnet einen Datei- oder Ordnerdialog und aktualisiert das entsprechende Eingabefeld."""
        path_edit = getattr(self, f"{key}_edit")
        dialog_title = "Ordner auswählen" if 'ordner' in key else "Excel-Datei auswählen"
        start_path = path_edit.text()
        if 'ordner' in key:
            path = QFileDialog.getExistingDirectory(self, dialog_title, start_path)
        else:
            path, _ = QFileDialog.getOpenFileName(self, dialog_title, start_path, "*.xlsx *.xls")
        if path:
            path_edit.setText(path)
            self.on_path_change(key, path)
            self.validate_path(path_edit)

    def validate_path(self, line_edit_widget):
        """Prüft, ob der Pfad in einem QLineEdit existiert und färbt es entsprechend."""
        path = line_edit_widget.text()
        is_optional_empty = 'bilder_ordner' in self.paths and hasattr(self, 'bilder_ordner_edit') and line_edit_widget == self.bilder_ordner_edit and not path

        valid_style = "background-color: #d4edda; color: #155724;"  # Greenish
        invalid_style = "background-color: #f8d7da; color: #721c24;"  # Reddish

        if os.path.exists(path) or is_optional_empty:
            line_edit_widget.setStyleSheet(valid_style)
        else:
            line_edit_widget.setStyleSheet(invalid_style)

    def add_category_widget(self, prefix, name):
        """Fügt der UI eine neue Zeile zur Eingabe einer Kategorie hinzu."""
        row_layout = QHBoxLayout()
        checkbox = QCheckBox()
        prefix_edit = QLineEdit(prefix)
        name_edit = QLineEdit(name)

        row_layout.addWidget(checkbox)
        row_layout.addWidget(QLabel("Präfix:"))
        row_layout.addWidget(prefix_edit)
        row_layout.addWidget(QLabel("Ordnername:"))
        row_layout.addWidget(name_edit)

        self.categories_ui_layout.addLayout(row_layout)
        self.category_widgets[row_layout] = (checkbox, prefix_edit, name_edit)

    def add_category(self):
        """Event-Handler, um eine neue, leere Kategorie-Zeile hinzuzufügen."""
        self.add_category_widget("", "")

    def remove_category(self):
        """Event-Handler, um alle ausgewählten Kategorie-Zeilen zu entfernen."""
        to_remove = [l for l, (c, _, _) in self.category_widgets.items() if c.isChecked()]
        for layout in to_remove:
            while layout.count():
                child = layout.takeAt(0)
                if child.widget():
                    child.widget().deleteLater()
            self.categories_ui_layout.removeItem(layout)
            layout.deleteLater()
            del self.category_widgets[layout]

    def save_categories(self):
        """Sammelt die Daten aus den Kategorie-Eingabefeldern und speichert sie."""
        self.categories.clear()
        # Werte bestehen aus (checkbox, prefix_edit, name_edit)
        for checkbox, prefix_edit, name_edit in self.category_widgets.values():
            prefix = prefix_edit.text().strip()
            name = name_edit.text().strip()
            if prefix and name:
                self.categories[prefix] = name
        self.save_all_settings()
        QMessageBox.information(self, "Gespeichert", "Kategorien aktualisiert.")

    def save_all_settings(self):
        """Sammelt alle aktuellen Einstellungen und speichert sie in 'settings.json'."""
        for key in self.settings:
            if hasattr(self, f"{key}_spin"):
                self.settings[key] = getattr(self, f"{key}_spin").value()

        if hasattr(self, 'datetime_utc_format_edit'):
            self.settings['datetime_utc_format'] = self.datetime_utc_format_edit.text()
        if hasattr(self, 'projekt_name_override_edit'):
            self.settings['projekt_name_override'] = self.projekt_name_override_edit.text().strip()
        if hasattr(self, 'export_as_pdf_check'):
            self.settings['export_as_pdf'] = self.export_as_pdf_check.isChecked()

        try:
            payload = {**self.paths, **self.settings, 'categories': self.categories}
            # Projektverwaltung: letzte Projekte und aktueller Projektpfad speichern
            if hasattr(self, 'recent_projects'):
                payload['recent_projects'] = self.recent_projects
            if hasattr(self, 'current_project_path') and self.current_project_path:
                payload['last_project_path'] = self.current_project_path
            with open('settings.json', 'w', encoding='utf-8') as f:
                json.dump(payload, f, indent=4, ensure_ascii=False)
        except IOError as e:
            self.log_text.append(f"Speicherfehler: {e}")
            _log_handler(f"Speicherfehler: {e}", "ERROR", self.append_html_log)

    def load_settings(self):
        """Lädt Einstellungen aus 'settings.json', falls die Datei existiert."""
        try:
            if os.path.exists('settings.json'):
                with open('settings.json', 'r', encoding='utf-8') as f:
                    return json.load(f)
        except (json.JSONDecodeError, IOError):
            pass
        return {}

    def start_dry_run(self):
        """Startet den Trockenlauf-Prozess."""
        if not all(self.paths.get(k) for k in ['excel_path', 'vorlagen_ordner']):
            QMessageBox.warning(self, "Fehlende Pfade", "Bitte Excel- und Vorlagen-Pfad für den Trockenlauf angeben!")
            return
        self.save_all_settings()
        self.log_text.clear()

        worker_params = {
            **self.paths,
            **self.settings,
            'categories': self.categories,
            'dry_run': True,
            'is_dark_mode': self.is_dark_mode
        }

        self.worker = Worker(**worker_params)
        self.worker.finished.connect(self.on_worker_finished)
        self.worker.log.connect(self.append_html_log)

        # Für den Trockenlauf brauchen wir diese Signale nicht, aber der Worker erwartet sie
        self.worker.current_file.connect(lambda: None)
        self.worker.progress.connect(lambda: None)

        self.worker.start()
        self.set_ui_running_state(True)

    def start(self, dry_run=False):
        """
        Startet den Dokumentenerstellungsprozess.
        - Prüft, ob alle notwendigen Pfade angegeben sind.
        - Speichert die aktuellen Einstellungen.
        - Erstellt und startet den Worker-Thread mit allen notwendigen Parametern.
        - Deaktiviert den 'Start'-Knopf, um doppelte Ausführungen zu verhindern.
        """
        if not all(self.paths.get(k) for k in ['excel_path', 'vorlagen_ordner', 'export_ordner']):
            QMessageBox.warning(self, "Fehlende Pfade", "Bitte Excel-, Vorlagen- und Export-Pfad angeben!")
            return
        self.save_all_settings()
        self.log_text.clear()

        worker_params = {
            **self.paths,
            **self.settings,
            'categories': self.categories,
            'dry_run': False,
            'is_dark_mode': self.is_dark_mode
        }

        self.worker = Worker(**worker_params)

        self.worker.finished.connect(self.on_worker_finished)
        self.worker.log.connect(self.append_html_log)
        self.worker.current_file.connect(self.update_current_file_label)
        self.worker.progress.connect(self.update_progress_bar)

        self.worker.start()
        self.set_ui_running_state(True)

    def set_ui_running_state(self, is_running):
        """Aktiviert/Deaktiviert UI-Elemente, während der Worker läuft."""
        self.start_btn.setEnabled(not is_running)
        self.dry_run_btn.setEnabled(not is_running)
        self.close_btn.setText("Abbrechen" if is_running else "Schließen")
        if not is_running:
            self.current_file_label.setText("Bereit zum Starten...")
        self.open_export_folder_btn.setVisible(False)

    def handle_close_or_cancel(self):
        """Schließt die App oder bricht den Worker ab."""
        if hasattr(self, 'worker') and self.worker.isRunning():
            reply = QMessageBox.question(
                self,
                "Abbrechen?",
                "Möchten Sie den aktuellen Vorgang wirklich abbrechen?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            if reply == QMessageBox.Yes:
                self.worker.requestInterruption()
        else:
            self.close()

    def on_worker_finished(self, success, export_path):
        """Wird aufgerufen, wenn der Worker-Thread seine Arbeit beendet hat."""
        self.set_ui_running_state(False)
        self.last_export_path = export_path

        is_dry_run = self.worker.dry_run
        was_cancelled = self.worker.isInterruptionRequested()

        if was_cancelled:
            self.current_file_label.setText("Vorgang abgebrochen.")
            QMessageBox.warning(self, "Abgebrochen", "Der Vorgang wurde vom Benutzer abgebrochen.")
        elif is_dry_run:
            if success:
                self.current_file_label.setText("Trockenlauf erfolgreich!")
                QMessageBox.information(self, "Trockenlauf", "Konfigurationsprüfung erfolgreich abgeschlossen.")
            else:
                self.current_file_label.setText("Trockenlauf fehlgeschlagen.")
                QMessageBox.warning(
                    self,
                    "Trockenlauf",
                    "Konfigurationsprüfung hat Fehler gefunden. Bitte Logs prüfen."
                )
        elif success and self.last_export_path:
            self.current_file_label.setText("Verarbeitung abgeschlossen!")
            self.open_export_folder_btn.setVisible(True)
            QMessageBox.information(self, "Fertig", "Alle Dokumente wurden erfolgreich erstellt.")
        else:
            self.current_file_label.setText("Verarbeitung mit Fehlern abgeschlossen.")
            QMessageBox.warning(
                self,
                "Fehler",
                "Die Verarbeitung wurde mit Fehlern abgeschlossen. Bitte Logs prüfen."
            )

    def update_current_file_label(self, filename):
        """Aktualisiert das Label, das den Namen der aktuell verarbeiteten Datei anzeigt."""
        self.current_file_label.setText(f"Verarbeite: {filename}...")

    def update_progress_bar(self, current, total):
        """Aktualisiert die Fortschrittsanzeige."""
        if total > 0:
            self.progress.setValue(int(current / total * 100))
        else:
            self.progress.setValue(0)

    def append_html_log(self, html):
        """Ein Slot, der HTML-formatierten Text sicher an das Log-Fenster anhängt und für Filter speichert."""
        import re
        # Robust: Suche nach [ LEVEL ] mit beliebigen Leerzeichen, unabhängig von HTML-Tags
        m = re.search(r'\[\s*([A-Z]+)\s*\]', html)
        level = m.group(1) if m else "INFO"
        self.log_lines.append((level, html))
        self.apply_log_filter()

    def apply_log_filter(self):
        """Zeigt nur die Log-Zeilen an, die dem gewählten Filter entsprechen."""
        filter_level = self.log_filter_combo.currentText()
        self.log_text.clear()
        for level, html in self.log_lines:
            if filter_level == "ALLE" or level == filter_level:
                self.log_text.moveCursor(QTextCursor.End)
                self.log_text.insertHtml(html)
        self.log_text.ensureCursorVisible()

    def show_log_context_menu(self, position):
        """Zeigt ein Kontextmenü für Log-Einträge mit Optionen zum Öffnen von Ordnern/Dateien/Textmarken."""
        cursor = self.log_text.cursorForPosition(position)
        cursor.select(QTextCursor.WordUnderCursor)
        selected_text = cursor.selectedText()

        # Suche nach Pfaden in der aktuellen Zeile
        line_cursor = self.log_text.cursorForPosition(position)
        line_cursor.movePosition(QTextCursor.StartOfLine)
        line_cursor.movePosition(QTextCursor.EndOfLine, QTextCursor.KeepAnchor)
        line_text = line_cursor.selectedText()

        menu = QMenu(self)

        # Extrahiere mögliche Pfade aus der Zeile
        paths = self.extract_paths_from_log_line(line_text)

        if paths:
            # Füge Menüpunkte für jeden gefundenen Pfad hinzu
            for path_type, path, display_name in paths:
                if path_type == "file":
                    action = menu.addAction(f"📄 {display_name} öffnen")
                    action.triggered.connect(lambda checked, p=path: self.open_file(p))
                elif path_type == "folder":
                    action = menu.addAction(f"📁 {display_name} öffnen")
                    action.triggered.connect(lambda checked, p=path: self.open_folder(p))
                elif path_type == "template":
                    action = menu.addAction(f"📝 {display_name} öffnen")
                    action.triggered.connect(lambda checked, p=path: self.open_template_file(p))
            menu.addSeparator()

        # Prüfe auf fehlende Platzhalter/Textmarken
        fehlende_marken = self.extract_missing_placeholders_from_log_line(line_text)
        if fehlende_marken:
            action = menu.addAction("➕ Textmarke(n) im Excel hinzufügen")
            action.triggered.connect(lambda checked, marken=fehlende_marken: self.add_placeholders_to_excel(marken))
            menu.addSeparator()

        # Prüfe auf fehlende Size-Textmarken
        fehlende_size_marken = self.extract_missing_size_placeholders_from_log_line(line_text)
        if fehlende_size_marken:
            action = menu.addAction("➕ Size-Textmarke(n) im Excel hinzufügen")
            action.triggered.connect(lambda checked, marken=fehlende_size_marken: self.add_placeholders_to_excel(marken))
            menu.addSeparator()

        # Standard-Menüpunkte
        copy_action = menu.addAction("📋 Ausgewählten Text kopieren")
        copy_action.triggered.connect(self.copy_selected_text)

        clear_action = menu.addAction("🗑️ Log löschen")
        clear_action.triggered.connect(self.clear_log)

        menu.exec_(self.log_text.mapToGlobal(position))

    def extract_paths_from_log_line(self, line_text):
        """Extrahiert mögliche Datei- und Ordnerpfade aus einer Log-Zeile."""
        paths = []
        import re
        # Suche nach Dateinamen in verschiedenen Log-Formaten
        # 1. Export-Log: Dokument: 'Dateiname.docx'
        datei_matches = re.findall(r"Dokument: '([^']+)'", line_text)
        # 2. Fehler/Warnung: Vorlage 'Dateiname.docx'
        vorlage_matches = re.findall(r"Vorlage '([^']+)'", line_text)
        # 3. Export-Log: in Ordner: 'Ordnername'
        ordner_matches = re.findall(r"in Ordner: '([^']+)'", line_text)

        # Exportierte Dateien suchen (wie bisher)
        if hasattr(self, 'last_export_path') and self.last_export_path:
            base_path = self.last_export_path
            # Füge Dateien aus Export-Log hinzu
            for datei in datei_matches:
                for category_name in self.categories.values():
                    potential_path = os.path.join(base_path, category_name, datei)
                    if os.path.exists(potential_path):
                        paths.append(("file", potential_path, datei))
                        break
            # Füge Dateien aus Fehler/Warnung hinzu
            for datei in vorlage_matches:
                for category_name in self.categories.values():
                    potential_path = os.path.join(base_path, category_name, datei)
                    if os.path.exists(potential_path):
                        paths.append(("file", potential_path, datei))
                        break
            # Füge Ordner hinzu
            for ordner in ordner_matches:
                potential_path = os.path.join(base_path, ordner)
                if os.path.exists(potential_path):
                    paths.append(("folder", potential_path, ordner))

        # Vorlagen-Dateien suchen
        # Suche in self.paths['vorlagen_ordner'] und allen Unterordnern nach passenden Dateinamen
        if hasattr(self, 'paths') and 'vorlagen_ordner' in self.paths:
            vorlagen_root = self.paths['vorlagen_ordner']
            alle_vorlagen = set(datei_matches + vorlage_matches)
            for suchname in alle_vorlagen:
                for root, _, files in os.walk(vorlagen_root):
                    for file in files:
                        if file == suchname:
                            full_path = os.path.join(root, file)
                            paths.append(("template", full_path, suchname))
        return paths

    def extract_missing_placeholders_from_log_line(self, line_text):
        """Extrahiert fehlende Platzhalter/Textmarken aus einer Log-Zeile."""
        import re
        # Sucht nach: Fehlende Platzhalter in Excel: {'foo', 'bar'}
        match = re.search(r"Fehlende Platzhalter in Excel: (\{.*?\})", line_text)
        if match:
            try:
                # Sichere Auswertung des Sets
                marken = eval(match.group(1), {"__builtins__": None}, {})
                if isinstance(marken, set):
                    return sorted(marken)
            except Exception:
                pass
        return []

    def extract_missing_size_placeholders_from_log_line(self, line_text):
        """Extrahiert fehlende Size-Textmarken aus einer Log-Zeile."""
        import re
        # NEU: Erkenne auch die neue Info-Log-Zeile
        # Beispiel: Hinweis: Für Platzhalter 'ba_logo_img' wird die Standardgröße verwendet, da kein 'ba_logo_img_size' oder 'ba_logo_img_Size' in Excel/Fallback gefunden wurde.
        match_alt = re.search(r"Keine Größenangabe\s*\(([^)]+)\).*?für Platzhalter", line_text)
        match_neu = re.search(
            r"Hinweis: Für Platzhalter '([^']+)' wird die Standardgröße verwendet, da kein '([^']+)' oder '([^']+)' in Excel/Fallback gefunden wurde",
            line_text
        )
        if match_alt:
            inhalt = match_alt.group(1)
            self.append_html_log(f'<span style="color:#888">[DEBUG] Klammer-Inhalt: {inhalt}</span><br>')
            marken = re.findall(r"'([^']+)'", inhalt)
            if marken:
                self.append_html_log(
                    f'<span style="color:#888">[DEBUG] Erkannte Size-Textmarken: {", ".join(marken)}</span><br>'
                )
            else:
                self.append_html_log(f'<span style="color:#888">[DEBUG] Keine Size-Textmarken erkannt.</span><br>')
            return sorted(marken)
        elif match_neu:
            # Extrahiere die beiden _size-Varianten
            size1 = match_neu.group(2)
            size2 = match_neu.group(3)
            marken = [size1, size2]
            self.append_html_log(
                f'<span style="color:#888">[DEBUG] Erkannte Size-Textmarken (neu): {", ".join(marken)}</span><br>'
            )
            return sorted(marken)
        self.append_html_log(f'<span style="color:#888">[DEBUG] Keine Size-Textmarken erkannt.</span><br>')
        return []

    def add_placeholders_to_excel(self, placeholders):
        """Fügt die angegebenen Platzhalter als neue Zeilen im zweiten Blatt der aktuellen Excel-Datei hinzu (Spalte 2=Textmarke, Spalte 3=leer)."""
        from openpyxl import load_workbook
        import os
        excel_path = self.paths.get('excel_path')
        if not excel_path or not os.path.exists(excel_path):
            QMessageBox.warning(self, "Fehler", "Excel-Datei nicht gefunden!")
            return
        wb = None
        try:
            wb = load_workbook(excel_path)
            if len(wb.sheetnames) < 2:
                QMessageBox.warning(self, "Fehler", "Die Excel-Datei hat kein zweites Blatt!")
                return
            ws = wb[wb.sheetnames[1]]
            # Nur eine Size-Variante pro Basis-Textmarke einfügen
            filtered = {}
            for marke in placeholders:
                if marke.lower().endswith('_size'):
                    base = marke[:-5].lower()
                    # Bevorzuge die Variante mit kleinem 's'
                    if base not in filtered or marke.endswith('_size'):
                        filtered[base] = marke
                else:
                    filtered[marke] = marke
            for marke in filtered.values():
                ws.append([None, marke, ""])
            wb.save(excel_path)
            QMessageBox.information(
                self,
                "Erfolg",
                f"{len(filtered)} Textmarke(n) wurden im zweiten Blatt hinzugefügt."
            )
        except Exception as e:
            QMessageBox.warning(self, "Fehler", f"Konnte Textmarken nicht hinzufügen: {e}")
        finally:
            if wb is not None:
                wb.close()

    def open_file(self, file_path):
        """Öffnet eine Datei mit der Standard-Anwendung."""
        try:
            os.startfile(file_path)
        except Exception as e:
            QMessageBox.warning(self, "Fehler", f"Konnte die Datei nicht öffnen: {e}")

    def open_folder(self, folder_path):
        """Öffnet einen Ordner im Explorer."""
        try:
            os.startfile(folder_path)
        except Exception as e:
            QMessageBox.warning(self, "Fehler", f"Konnte den Ordner nicht öffnen: {e}")

    def open_template_file(self, file_path):
        """Öffnet eine Word-Vorlage mit der Standard-Anwendung."""
        try:
            os.startfile(file_path)
        except Exception as e:
            QMessageBox.warning(self, "Fehler", f"Konnte die Vorlage nicht öffnen: {e}")

    def copy_selected_text(self):
        """Kopiert den ausgewählten Text in die Zwischenablage."""
        cursor = self.log_text.textCursor()
        if cursor.hasSelection():
            text = cursor.selectedText()
            clipboard = QApplication.clipboard()
            clipboard.setText(text)

    def clear_log(self):
        """Löscht den gesamten Log-Inhalt."""
        reply = QMessageBox.question(
            self,
            "Log löschen",
            "Möchten Sie wirklich den gesamten Log löschen?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.log_text.clear()

    def show_excel_data(self):
        """
        Versucht, die Daten aus der angegebenen Excel-Datei zu laden und
        in der Tabelle (QTableView) in der GUI anzuzeigen.
        Bei gesperrter Datei (z. B. in Excel geöffnet) wird unter Windows
        ein COM-Fallback verwendet.
        """
        path = self.paths.get('excel_path')
        header_row = self.settings.get('header_row', 3)
        if path and os.path.exists(path):
            try:
                df = pd.read_excel(path, header=header_row - 1)
                self.table.setModel(PandasModel(df))
            except (PermissionError, OSError) as e:
                if getattr(e, 'errno', None) == 13 or isinstance(e, PermissionError):
                    try:
                        from core.excel_com import read_excel_sheet1_via_com
                        df = read_excel_sheet1_via_com(path, header_row)
                        if df is not None:
                            self.table.setModel(PandasModel(df))
                            return
                    except Exception:
                        pass
                self.table.setModel(None)
                QMessageBox.critical(self, "Fehler beim Lesen der Excel-Datei", str(e))
            except Exception as e:
                self.table.setModel(None)  # Clear table on error
                QMessageBox.critical(self, "Fehler beim Lesen der Excel-Datei", str(e))

    def open_export_folder(self):
        """Öffnet den zuletzt verwendeten Export-Ordner im Datei-Explorer."""
        if self.last_export_path and os.path.isdir(self.last_export_path):
            try:
                os.startfile(self.last_export_path)
            except Exception as e:
                QMessageBox.warning(self, "Fehler", f"Konnte den Ordner nicht öffnen: {e}")

    def open_excel_file(self):
        excel_path = self.paths.get('excel_path')
        if excel_path and os.path.exists(excel_path):
            try:
                os.startfile(excel_path)
            except Exception as e:
                QMessageBox.warning(self, "Fehler", f"Excel konnte nicht geöffnet werden: {e}")
        else:
            QMessageBox.warning(self, "Fehler", "Excel-Datei nicht gefunden!")

    def closeEvent(self, event):
        """Stellt sicher, dass der Worker-Thread sauber beendet wird, wenn das Fenster geschlossen wird."""
        if hasattr(self, 'worker') and self.worker.isRunning():
            self.worker.requestInterruption()
            self.worker.wait()
        event.accept()
