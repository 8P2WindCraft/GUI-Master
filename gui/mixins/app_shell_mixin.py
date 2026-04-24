"""PDF-Tab, Log-Tab, Menüleiste (Projekt, Theme)."""
from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QActionGroup
from PySide6.QtWidgets import (
    QComboBox,
    QHBoxLayout,
    QLabel,
    QProgressBar,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from ..constants import BUILD_DATE, VERSION


class AppShellMixin:
    """Manuelle PDF-Ansicht, Log mit Filter, Menüs Projekte/Ansicht."""

    def setup_pdf_tab(self, tab):
        """Erstellt den separaten PDF-Reiter für die manuelle Konvertierung des letzten Exports."""
        layout = QVBoxLayout(tab)

        info = QLabel("PDF-Umwandlung für den letzten erfolgreichen Export")
        info.setStyleSheet("font-weight: 600;")
        layout.addWidget(info)

        self.pdf_last_export_label = QLabel("Letzter Export: —")
        layout.addWidget(self.pdf_last_export_label)

        self.pdf_progress = QProgressBar()
        self.pdf_progress.setValue(0)
        layout.addWidget(self.pdf_progress)

        self.pdf_current_file_label = QLabel("Bereit.")
        self.pdf_current_file_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.pdf_current_file_label)

        self.pdf_summary_label = QLabel("Noch keine PDF-Konvertierung gestartet.")
        self.pdf_summary_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.pdf_summary_label)

        btn_row = QHBoxLayout()
        self.pdf_convert_btn = QPushButton("Letzten Export in PDF umwandeln")
        self.pdf_convert_btn.clicked.connect(self.convert_last_export_to_pdf)
        btn_row.addWidget(self.pdf_convert_btn)

        self.pdf_open_folder_btn = QPushButton("Letzten Export-Ordner öffnen")
        self.pdf_open_folder_btn.clicked.connect(self.open_export_folder)
        btn_row.addWidget(self.pdf_open_folder_btn)

        self.pdf_select_folder_btn = QPushButton("Ordner auswählen...")
        self.pdf_select_folder_btn.clicked.connect(self.select_pdf_source_folder)
        btn_row.addWidget(self.pdf_select_folder_btn)
        layout.addLayout(btn_row)

        self.pdf_log_text = QTextEdit()
        self.pdf_log_text.setReadOnly(True)
        layout.addWidget(self.pdf_log_text)

        self.refresh_pdf_tab()

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
