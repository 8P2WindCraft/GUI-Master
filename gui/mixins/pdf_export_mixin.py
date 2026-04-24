"""PDF-Reiter: letzter Export-Ordner, manuelle Konvertierung mit PdfWorker."""
import os

from PySide6.QtGui import QTextCursor
from PySide6.QtWidgets import QFileDialog, QMessageBox

from core.logging import _log_handler

from ..workers import PdfWorker


class PdfExportMixin:
    """Buttons/Label/Log im PDF-Tab; ``update_export_action_buttons`` auch für die Generierung."""

    def update_export_action_buttons(self, is_running=False):
        """Aktualisiert die Buttons für Export-Aktionen basierend auf dem letzten Exportordner."""
        has_export_path = bool(self.last_export_path and os.path.isdir(self.last_export_path))
        is_pdf_running = hasattr(self, 'pdf_worker') and self.pdf_worker and self.pdf_worker.isRunning()
        export_actions_enabled = has_export_path and not is_running and not is_pdf_running
        select_actions_enabled = not is_running and not is_pdf_running
        self.open_export_folder_btn.setVisible(export_actions_enabled)
        if hasattr(self, 'pdf_convert_btn'):
            self.pdf_convert_btn.setEnabled(export_actions_enabled)
        if hasattr(self, 'pdf_open_folder_btn'):
            self.pdf_open_folder_btn.setEnabled(export_actions_enabled)
        if hasattr(self, 'pdf_select_folder_btn'):
            self.pdf_select_folder_btn.setEnabled(select_actions_enabled)
        if hasattr(self, 'pdf_last_export_label'):
            self.pdf_last_export_label.setText(
                f"Letzter Export: {self.last_export_path}" if has_export_path else "Letzter Export: —"
            )

    def refresh_pdf_tab(self):
        """Aktualisiert den PDF-Reiter basierend auf dem letzten Exportpfad."""
        self.update_export_action_buttons(False)
        if not (self.last_export_path and os.path.isdir(self.last_export_path)):
            self.pdf_summary_label.setText(
                "Kein letzter Export gefunden. Bitte Ordner auswählen oder zuerst Dokumente generieren."
            )
            if self.pdf_progress.value() != 0:
                self.pdf_progress.setValue(0)
            if self.pdf_current_file_label.text() == "Bereit.":
                return
            self.pdf_current_file_label.setText("Bereit.")

    def select_pdf_source_folder(self):
        """Erlaubt die manuelle Auswahl eines Export-Ordners für die PDF-Konvertierung."""
        if hasattr(self, 'worker') and self.worker.isRunning():
            QMessageBox.warning(self, "Bitte warten", "Während der Verarbeitung ist keine Ordnerauswahl möglich.")
            return
        if hasattr(self, 'pdf_worker') and self.pdf_worker and self.pdf_worker.isRunning():
            QMessageBox.warning(self, "Bitte warten", "Während der PDF-Konvertierung ist keine Ordnerauswahl möglich.")
            return

        start_dir = ""
        if self.last_export_path and os.path.isdir(self.last_export_path):
            start_dir = self.last_export_path
        elif self.paths.get('export_ordner') and os.path.isdir(self.paths.get('export_ordner')):
            start_dir = self.paths.get('export_ordner')
        else:
            start_dir = os.getcwd()

        selected_dir = QFileDialog.getExistingDirectory(self, "Export-Ordner für PDF auswählen", start_dir)
        if not selected_dir:
            return

        self.last_export_path = selected_dir
        self.update_export_action_buttons(False)
        self.pdf_summary_label.setText("Export-Ordner manuell gesetzt. Bereit zur PDF-Konvertierung.")
        self.pdf_current_file_label.setText("Bereit.")

    def convert_last_export_to_pdf(self):
        """Konvertiert alle DOCX-Dateien aus dem letzten Exportordner nach PDF."""
        if hasattr(self, 'worker') and self.worker.isRunning():
            QMessageBox.warning(self, "Bitte warten", "Während der Verarbeitung ist die PDF-Konvertierung nicht möglich.")
            return
        if hasattr(self, 'pdf_worker') and self.pdf_worker and self.pdf_worker.isRunning():
            QMessageBox.warning(self, "Bitte warten", "Die PDF-Konvertierung läuft bereits.")
            return

        export_path = self.last_export_path
        if not export_path:
            QMessageBox.warning(self, "Kein Export", "Es wurde noch kein Export durchgeführt.")
            return
        if not os.path.isdir(export_path):
            _log_handler(f"Export-Ordner nicht gefunden: {export_path}", "ERROR", self.append_html_log, self.is_dark_mode)
            QMessageBox.warning(self, "Ordner fehlt", f"Der letzte Export-Ordner existiert nicht mehr:\n{export_path}")
            self.update_export_action_buttons(False)
            return

        existing_pdf_count = 0
        for root, _, files in os.walk(export_path):
            existing_pdf_count += sum(1 for name in files if name.lower().endswith('.pdf'))
        if existing_pdf_count > 0:
            QMessageBox.warning(
                self,
                "Hinweis",
                f"Im Ordner sind bereits {existing_pdf_count} PDF-Datei(en) vorhanden. "
                "Die Umwandlung wird erneut ausgeführt."
            )

        self.current_file_label.setText("PDF-Konvertierung läuft...")
        self.pdf_current_file_label.setText("PDF-Konvertierung gestartet...")
        self.pdf_summary_label.setText("Konvertierung läuft...")
        self.pdf_progress.setValue(0)
        self.pdf_log_text.clear()
        self.update_export_action_buttons(True)
        self.pdf_worker = PdfWorker(export_path, self.is_dark_mode)
        self.pdf_worker.log.connect(self.append_html_log)
        self.pdf_worker.log.connect(self.append_pdf_html_log)
        self.pdf_worker.progress.connect(self.update_pdf_progress_bar)
        self.pdf_worker.current_file.connect(self.update_pdf_current_file_label)
        self.pdf_worker.finished.connect(self.on_pdf_worker_finished)

        for callback in (self.append_html_log, self.append_pdf_html_log):
            _log_handler("=" * 60 + "\n", "SEP", callback, self.is_dark_mode)
            _log_handler("Starte manuelle PDF-Konvertierung für letzten Export...", "INFO", callback, self.is_dark_mode)
            _log_handler(f"Export-Ordner: {export_path}", "INFO", callback, self.is_dark_mode)
            _log_handler("=" * 60 + "\n", "SEP", callback, self.is_dark_mode)
        self.pdf_worker.start()

    def append_pdf_html_log(self, html):
        """Fügt HTML-Logs in den PDF-Reiter ein."""
        if not hasattr(self, 'pdf_log_text'):
            return
        self.pdf_log_text.moveCursor(QTextCursor.End)
        self.pdf_log_text.insertHtml(html)
        self.pdf_log_text.ensureCursorVisible()

    def update_pdf_progress_bar(self, current, total):
        """Aktualisiert den Fortschritt der PDF-Konvertierung."""
        if not hasattr(self, 'pdf_progress'):
            return
        if total > 0:
            self.pdf_progress.setValue(int(current / total * 100))
        else:
            self.pdf_progress.setValue(0)

    def update_pdf_current_file_label(self, filename):
        """Zeigt die aktuell konvertierte Datei im PDF-Reiter."""
        if hasattr(self, 'pdf_current_file_label'):
            self.pdf_current_file_label.setText(f"Konvertiere: {filename}")

    def on_pdf_worker_finished(self, result):
        """Wird aufgerufen, wenn die PDF-Konvertierung abgeschlossen wurde."""
        docx_count = result.get("docx_count", 0)
        success_count = result.get("success_count", 0)
        failure_count = result.get("failure_count", 0)
        error_type = result.get("error")

        if error_type == "missing_docx2pdf":
            QMessageBox.warning(
                self,
                "Modul fehlt",
                "Das Modul 'docx2pdf' ist nicht installiert.\nBitte ausführen: pip install docx2pdf"
            )
            self.current_file_label.setText("PDF-Konvertierung nicht möglich (docx2pdf fehlt).")
            self.pdf_summary_label.setText("Fehler: docx2pdf ist nicht installiert.")
            self.pdf_current_file_label.setText("Keine Konvertierung durchgeführt.")
        elif error_type == "missing_pywin32":
            QMessageBox.warning(
                self,
                "Modul fehlt",
                "Das Modul 'pywin32' ist nicht verfügbar oder unvollständig eingerichtet.\n"
                "Bitte in der aktiven Umgebung ausführen:\n"
                "pip install pywin32\n"
                "pywin32_postinstall.py -install"
            )
            self.current_file_label.setText("PDF-Konvertierung nicht möglich (pywin32 fehlt).")
            self.pdf_summary_label.setText("Fehler: pywin32 ist nicht verfügbar.")
            self.pdf_current_file_label.setText("Keine Konvertierung durchgeführt.")
        elif error_type == "unknown_error":
            error_message = result.get("error_message", "Unbekannter Fehler")
            QMessageBox.critical(self, "PDF-Konvertierung fehlgeschlagen", f"Unerwarteter Fehler:\n{error_message}")
            self.current_file_label.setText("PDF-Konvertierung fehlgeschlagen.")
            self.pdf_summary_label.setText("PDF-Konvertierung fehlgeschlagen.")
            self.pdf_current_file_label.setText("Fehler aufgetreten.")
        elif docx_count == 0:
            QMessageBox.information(self, "Keine Word-Dateien", "Im letzten Export wurden keine .docx-Dateien gefunden.")
            self.current_file_label.setText("Keine Word-Dateien für PDF-Konvertierung gefunden.")
            self.pdf_summary_label.setText("Keine .docx-Dateien gefunden.")
            self.pdf_current_file_label.setText("Bereit.")
        elif failure_count == 0:
            QMessageBox.information(self, "PDF-Konvertierung fertig", f"{success_count} Datei(en) erfolgreich konvertiert.")
            self.current_file_label.setText("PDF-Konvertierung abgeschlossen.")
            self.pdf_summary_label.setText(f"Erfolg: {success_count} von {docx_count} Datei(en) konvertiert.")
            self.pdf_current_file_label.setText("Konvertierung abgeschlossen.")
            self.pdf_progress.setValue(100)
        else:
            QMessageBox.warning(
                self,
                "PDF-Konvertierung teilweise fehlgeschlagen",
                f"{success_count} von {docx_count} Datei(en) wurden konvertiert.\n"
                f"{failure_count} Datei(en) sind fehlgeschlagen. Bitte Logs prüfen."
            )
            self.current_file_label.setText("PDF-Konvertierung mit Fehlern abgeschlossen.")
            self.pdf_summary_label.setText(
                f"Teilweise erfolgreich: {success_count}/{docx_count}, Fehler: {failure_count}."
            )
            self.pdf_current_file_label.setText("Konvertierung abgeschlossen (mit Fehlern).")

        self.update_export_action_buttons(False)
