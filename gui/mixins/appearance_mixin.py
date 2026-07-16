"""
QSS-Themes und Reaktionen auf die vier Pfad-Eingabefelder.
"""
import os

from PySide6.QtWidgets import QFileDialog

from core.utils import resource_path


class ThemePathMixin:
    """Menü-Themes, QSS, Pfad-Dialoge und optische Validierung der Pfadzeilen."""

    def on_theme_selected(self, action):
        """Wird aufgerufen, wenn ein Theme aus dem Menü ausgewählt wird."""
        theme_name = action.data()
        self.apply_theme(theme_name)
        self.settings['theme'] = theme_name
        self.save_all_settings()

    def apply_theme(self, theme_name):
        """Wendet das ausgewählte Stylesheet an und aktualisiert die Menü-Auswahl. Setzt im Dark Mode den Log-Text auf weiß."""
        stylesheet = self.get_stylesheet(theme_name)
        self.setStyleSheet(stylesheet)  # Hauptfenster (unverändert bei Child-Widgets)
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
                    return ""  # Theme-Datei fehlt → Standard-Optik
        return ""  # Light / kein externes QSS

    def on_path_change(self, key, value):
        """Wird aufgerufen, wenn ein Pfad geändert wird, und speichert die Einstellung."""
        self.paths[key] = value
        self.save_all_settings()
        if hasattr(self, 'update_workflow_status'):
            self.update_workflow_status()

        # Aktualisiere Excel-Datenanzeige wenn sich der Excel-Pfad ändert
        if key == 'excel_path':
            self.show_excel_data()
        # Bilder-Vorschau-Tab aktualisieren wenn sich der Bilder-Ordner ändert
        if key == 'bilder_ordner' and hasattr(self, 'bilder_vorschau_table'):
            self.refresh_bilder_vorschau()
        if key == 'vorlagen_ordner':
            if hasattr(self, 'refresh_template_checkboxes'):
                self.refresh_template_checkboxes()
            if hasattr(self, 'refresh_rules_template_combos'):
                self.refresh_rules_template_combos()
        if key in {'excel_path', 'vorlagen_ordner', 'bilder_ordner'}:
            if hasattr(self, 'refresh_media_sizes_table'):
                self.refresh_media_sizes_table()

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
        if hasattr(self, 'project_root_dir_edit') and line_edit_widget == self.project_root_dir_edit and not path:
            is_optional_empty = True

        valid_style = "background-color: #eaf7ea; color: #1f3d1f;"  # Leichtes Grün bei gültigem Pfad
        invalid_style = "background-color: #f8d7da; color: #721c24;"  # Reddish

        if os.path.exists(path) or is_optional_empty:
            line_edit_widget.setStyleSheet(valid_style)
        else:
            line_edit_widget.setStyleSheet(invalid_style)
