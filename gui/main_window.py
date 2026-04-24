"""
Hauptfenster: ``MainWindow`` (Tabs, Menüs). Logik in ``gui.mixins`` und ``gui.workers``/``gui.models``.
"""
import json
import os
import sys

from PySide6.QtWidgets import QMainWindow, QTabWidget, QVBoxLayout, QWidget

from core.logging import _log_handler

from .constants import BUILD_DATE, PROJECT_FILE_PATH_KEYS, VERSION
from .mixins import (
    AppShellMixin,
    BilderVorschauMixin,
    CategoriesTabMixin,
    GenerationWorkflowMixin,
    LogViewMixin,
    MainControlsMixin,
    MediaSizesMixin,
    PdfExportMixin,
    ProjectFileMixin,
    SignageRulesMixin,
    TemplatesMixin,
    ThemePathMixin,
)


class MainWindow(
    ProjectFileMixin,
    ThemePathMixin,
    MainControlsMixin,
    MediaSizesMixin,
    SignageRulesMixin,
    TemplatesMixin,
    CategoriesTabMixin,
    BilderVorschauMixin,
    AppShellMixin,
    GenerationWorkflowMixin,
    PdfExportMixin,
    LogViewMixin,
    QMainWindow,
):
    """
    Hauptfenster: Tabs (Hauptsteuerung, Vorlagen, Medien/QR, Regeln, Kategorien, Vorschau, PDF, Log),
    Menüs (Projekt *.dta.json, Themen) und Hintergrund-Worker für Generierung bzw. PDF.

    Pfade: Dict ``self.paths`` mit festen Keys ``PROJECT_FILE_PATH_KEYS``; persistent in
    ``settings.json``, Projekte (``*.dta.json``) zusätzlich ``categories`` und Regeln.
    """

    def __init__(self):
        super().__init__()
        self.setWindowTitle(f'DocxTpl Automatisierung v{VERSION} - Build {BUILD_DATE}')
        self.category_widgets = {}
        self.theme_action_group = None
        self.last_export_path = None
        # Aktuellen Windows-Benutzernamen ermitteln (für portable Pfade)
        self.current_username = os.environ.get('USERNAME', os.environ.get('USER', 'UnknownUser'))
        self.load_settings_and_init_vars()
        self.setup_ui()

    def load_settings_and_init_vars(self):
        """Lädt die zuletzt verwendeten Einstellungen (Pfade, Optionen) aus einer JSON-Datei."""
        last = self.load_settings()
        self.paths = {k: last.get(k, '') for k in PROJECT_FILE_PATH_KEYS}
        # Gespeicherte portable Pfade ({username}) sofort für dieses Konto auflösen
        self.paths = {k: self._denormalize_path(v) for k, v in self.paths.items()}
        self.settings = {
            'header_row': last.get('header_row', 3),
            'svg_scale': last.get('svg_scale', 3),
            'png_compression': last.get('png_compression', -1),
            'theme': last.get('theme', 'Light'),
            'datetime_utc_format': last.get('datetime_utc_format', '%Y-%m-%d %H:%M:%S UTC'),
            'projekt_name_override': last.get('projekt_name_override', ''),
            'export_as_pdf': last.get('export_as_pdf', False),
            'parallel_doc_generation': last.get('parallel_doc_generation', False),
        }
        self.categories = last.get(
            'categories',
            {'b_': 'Beschilderung', 'ba_': 'Betriebsanweisung', 'p_': 'Pläne'},
        )
        self.categories.setdefault('p_', 'Pläne')
        # None = alle Vorlagen (wie bisher); Liste = explizite Teilauswahl
        if 'selected_template_rel_paths' in last:
            self.selected_template_rel_paths = last.get('selected_template_rel_paths')
        else:
            self.selected_template_rel_paths = None
        self.rules_enabled = last.get('rules_enabled', False)
        self.reuse_lageplan_from_last_export = last.get('reuse_lageplan_from_last_export', False)
        self.signage_rules = last.get('signage_rules', [])
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

        # Tab 2: Vorlagen-Auswahl
        templates_tab = QWidget()
        tab_widget.addTab(templates_tab, "Vorlagen")
        self.vorlagen_tab_index = tab_widget.indexOf(templates_tab)
        self.setup_templates_tab(templates_tab)

        media_sizes_tab = QWidget()
        tab_widget.addTab(media_sizes_tab, "Bilder & QR-Größen")
        self.media_sizes_tab_index = tab_widget.indexOf(media_sizes_tab)
        self.setup_media_sizes_tab(media_sizes_tab)

        # Tab: Regeln (optionale Schilder)
        rules_tab = QWidget()
        tab_widget.addTab(rules_tab, "Regeln")
        self.rules_tab_index = tab_widget.indexOf(rules_tab)
        self.setup_rules_tab(rules_tab)

        # Tab 4: Kategorien
        categories_tab = QWidget()
        tab_widget.addTab(categories_tab, "Kategorien")
        self.setup_categories_tab(categories_tab)

        # Tab 5: Bilder-Vorschau
        bilder_vorschau_tab = QWidget()
        tab_widget.addTab(bilder_vorschau_tab, "Bilder-Vorschau")
        self.bilder_vorschau_tab_index = tab_widget.indexOf(bilder_vorschau_tab)
        self.setup_bilder_vorschau_tab(bilder_vorschau_tab)

        # Tab 6: PDF
        pdf_tab = QWidget()
        tab_widget.addTab(pdf_tab, "PDF")
        self.pdf_tab_index = tab_widget.indexOf(pdf_tab)
        self.setup_pdf_tab(pdf_tab)

        # Tab 7: Logs
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
        if hasattr(self, 'parallel_doc_generation_check'):
            self.settings['parallel_doc_generation'] = self.parallel_doc_generation_check.isChecked()

        self._snapshot_template_selection()
        self._snapshot_signage_rules_from_ui()

        try:
            payload = {**self.paths, **self.settings, 'categories': self.categories}
            payload['selected_template_rel_paths'] = self.selected_template_rel_paths
            payload['rules_enabled'] = self.rules_enabled
            payload['reuse_lageplan_from_last_export'] = self.reuse_lageplan_from_last_export
            payload['signage_rules'] = self.signage_rules
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

    def closeEvent(self, event):
        """Stellt sicher, dass der Worker-Thread sauber beendet wird, wenn das Fenster geschlossen wird."""
        if hasattr(self, 'worker') and self.worker.isRunning():
            self.worker.requestInterruption()
            self.worker.wait()
        if hasattr(self, 'pdf_worker') and self.pdf_worker and self.pdf_worker.isRunning():
            self.pdf_worker.wait()
        event.accept()
