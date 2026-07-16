"""Word-Generierung: Worker-Thread, Trockenlauf, Vorschau, Fortschritt."""
import os

from PySide6.QtWidgets import QMessageBox

from ..workers import Worker


class GenerationWorkflowMixin:
    """
    Start/Abbruch der docxtpl-Pipeline, Anbindung an ``gui.workers.Worker``,
    Label/Progress in der Hauptsteuerung, Zeilen-Vorschau im Vorlagen-Tab.
    """

    def _lageplan_previous_export_root(self):
        if self.reuse_lageplan_from_last_export and self.last_export_path and os.path.isdir(self.last_export_path):
            return self.last_export_path
        return None

    def _build_worker_params(self, **overrides):
        """Gemeinsame Worker-Argumente (Pfade, Regeln, Lageplan-Kontext); overrides zuletzt."""
        params = {
            **self.paths,
            **self.settings,
            'categories': self.categories,
            'is_dark_mode': self.is_dark_mode,
            'rules_enabled': self.rules_enabled,
            'signage_rules': list(self.signage_rules) if self.signage_rules else [],
            'reuse_lageplan_from_last_export': self.reuse_lageplan_from_last_export,
            'previous_export_root': self._lageplan_previous_export_root(),
        }
        params.update(overrides)
        return params

    def _connect_worker_with_progress(self, worker):
        worker.finished.connect(self.on_worker_finished)
        worker.log.connect(self.append_html_log)
        worker.current_file.connect(self.update_current_file_label)
        worker.progress.connect(self.update_progress_bar)

    def _ensure_worker_can_start(self):
        """Verhindert parallele Generierungslaeufe durch Doppelklicks oder zweite Aktionen."""
        worker = getattr(self, 'worker', None)
        if worker is not None and worker.isRunning():
            QMessageBox.information(
                self,
                "Vorgang läuft",
                "Es läuft bereits ein Vorgang. Bitte warten oder den aktuellen Vorgang abbrechen.",
            )
            return False
        return True

    def start_dry_run(self):
        """Startet den Trockenlauf-Prozess."""
        if not self._ensure_worker_can_start():
            return
        if not all(self.paths.get(k) for k in ['excel_path', 'vorlagen_ordner']):
            QMessageBox.warning(self, "Fehlende Pfade", "Bitte Excel- und Vorlagen-Pfad für den Trockenlauf angeben!")
            return
        if not getattr(self, '_template_checkbox_by_rel', {}):
            QMessageBox.warning(
                self,
                "Keine Vorlagen",
                "Im Vorlagen-Ordner wurden keine verwendbaren Vorlagen (Unterordner anlagen/ oder allgemein/) gefunden.",
            )
            return
        sel = self._template_paths_for_worker()
        if sel == []:
            QMessageBox.warning(
                self,
                "Keine Auswahl",
                "Bitte mindestens eine Vorlage auswählen.",
            )
            return
        self.save_all_settings()
        self.log_text.clear()

        worker_params = self._build_worker_params(
            dry_run=True,
            selected_template_paths=sel,
        )

        self.worker = Worker(**worker_params)
        self.worker.finished.connect(self.on_worker_finished)
        self.worker.log.connect(self.append_html_log)

        # Für den Trockenlauf brauchen wir diese Signale nicht, aber der Worker erwartet sie
        self.worker.current_file.connect(lambda: None)
        self.worker.progress.connect(lambda: None)

        self.worker.start()
        self.set_ui_running_state(True)

    @staticmethod
    def _first_pdf_under_dir(root_dir):
        if not root_dir or not os.path.isdir(root_dir):
            return None
        found = []
        for dirpath, _, files in os.walk(root_dir):
            for name in files:
                if name.lower().endswith('.pdf') and not name.startswith('~'):
                    found.append(os.path.join(dirpath, name))
        if not found:
            return None
        found.sort(key=lambda p: p.lower())
        return found[0]

    def start_preview_batch(self):
        """Stapel-Vorschau: je eine Ausgabe pro gewählter Vorlage, erste Excel-Zeile für Anlagen."""
        if not self._ensure_worker_can_start():
            return
        if not all(self.paths.get(k) for k in ['excel_path', 'vorlagen_ordner', 'export_ordner']):
            QMessageBox.warning(
                self,
                "Fehlende Pfade",
                "Bitte Excel-, Vorlagen- und Export-Pfad für die Vorschau angeben!",
            )
            return
        if not getattr(self, '_template_checkbox_by_rel', {}):
            QMessageBox.warning(
                self,
                "Keine Vorlagen",
                "Im Vorlagen-Ordner wurden keine verwendbaren Vorlagen (Unterordner anlagen/ oder allgemein/) gefunden.",
            )
            return
        sel = self._template_paths_for_worker()
        if sel == []:
            QMessageBox.warning(
                self,
                "Keine Auswahl",
                "Bitte mindestens eine Vorlage auswählen.",
            )
            return
        self.save_all_settings()
        self.log_text.clear()

        worker_params = self._build_worker_params(
            dry_run=False,
            selected_template_paths=sel,
            preview_run=True,
            preview_template_abs=None,
        )

        self.worker = Worker(**worker_params)
        self._connect_worker_with_progress(self.worker)
        self.worker.start()
        self.set_ui_running_state(True)

    def start_preview_single_template(self, template_abs, *, export_as_pdf=False):
        """Einzel-Vorschau für eine Vorlage (unabhängig von der Checkbox).

        ``export_as_pdf``: True = nach DOCX auch PDF erzeugen und bevorzugt öffnen (Word/pywin32).
        """
        if not self._ensure_worker_can_start():
            return
        if not all(self.paths.get(k) for k in ['excel_path', 'vorlagen_ordner', 'export_ordner']):
            QMessageBox.warning(
                self,
                "Fehlende Pfade",
                "Bitte Excel-, Vorlagen- und Export-Pfad für die Vorschau angeben!",
            )
            return
        if not template_abs or not os.path.isfile(template_abs):
            QMessageBox.warning(self, "Vorschau", "Vorlagendatei nicht gefunden.")
            return
        self.save_all_settings()
        self.log_text.clear()

        worker_params = self._build_worker_params(
            dry_run=False,
            selected_template_paths=None,
            preview_run=True,
            preview_template_abs=template_abs,
            export_as_pdf=bool(export_as_pdf),
        )

        label = os.path.basename(template_abs)
        if export_as_pdf:
            label = f"{label} (PDF)"
        self._show_template_row_preview_ui(label)
        self.worker = Worker(**worker_params)
        self._connect_worker_with_progress(self.worker)
        self.worker.current_file.connect(self._on_template_row_preview_file)
        self.worker.progress.connect(self._on_template_row_preview_progress)
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
        if not self._ensure_worker_can_start():
            return
        if not all(self.paths.get(k) for k in ['excel_path', 'vorlagen_ordner', 'export_ordner']):
            QMessageBox.warning(self, "Fehlende Pfade", "Bitte Excel-, Vorlagen- und Export-Pfad angeben!")
            return
        if not getattr(self, '_template_checkbox_by_rel', {}):
            QMessageBox.warning(
                self,
                "Keine Vorlagen",
                "Im Vorlagen-Ordner wurden keine verwendbaren Vorlagen (Unterordner anlagen/ oder allgemein/) gefunden.",
            )
            return
        sel = self._template_paths_for_worker()
        if sel == []:
            QMessageBox.warning(
                self,
                "Keine Auswahl",
                "Bitte mindestens eine Vorlage auswählen.",
            )
            return
        self.save_all_settings()
        self.log_text.clear()

        worker_params = self._build_worker_params(
            dry_run=False,
            selected_template_paths=sel,
        )

        self.worker = Worker(**worker_params)
        self._connect_worker_with_progress(self.worker)
        self.worker.start()
        self.set_ui_running_state(True)

    def set_ui_running_state(self, is_running):
        """Aktiviert/Deaktiviert UI-Elemente, während der Worker läuft."""
        self.start_btn.setEnabled(not is_running)
        self.dry_run_btn.setEnabled(not is_running)
        if hasattr(self, 'preview_btn'):
            self.preview_btn.setEnabled(not is_running)
        if hasattr(self, 'parallel_doc_generation_check'):
            self.parallel_doc_generation_check.setEnabled(not is_running)
        self.close_btn.setText("Abbrechen" if is_running else "Schließen")
        if not is_running:
            self.current_file_label.setText("Bereit zum Starten...")
        self.open_export_folder_btn.setVisible(False)
        self.update_export_action_buttons(is_running)
        if hasattr(self, 'update_workflow_status'):
            self.update_workflow_status()

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
        template_row_preview = getattr(self, '_template_row_preview_active', False)
        self.set_ui_running_state(False)

        is_dry_run = self.worker.dry_run
        is_preview = getattr(self.worker, 'preview_run', False)
        was_cancelled = self.worker.isInterruptionRequested()

        if not is_dry_run and not is_preview:
            self.last_export_path = export_path if (success and export_path) else None

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
        elif is_preview:
            if success and export_path:
                self.current_file_label.setText("Vorschau abgeschlossen.")
                want_pdf = bool(
                    getattr(self.worker, "params", {}).get("export_as_pdf", False)
                )
                opened = False
                if want_pdf:
                    pdf_p = self._first_pdf_under_dir(export_path)
                    if pdf_p:
                        try:
                            os.startfile(pdf_p)
                            opened = True
                        except OSError:
                            QMessageBox.warning(
                                self,
                                "Vorschau",
                                f"PDF konnte nicht geöffnet werden:\n{pdf_p}",
                            )
                    else:
                        QMessageBox.warning(
                            self,
                            "Vorschau",
                            "Es wurde keine PDF-Datei gefunden. Prüfen Sie, ob Word/pywin32 verfügbar ist "
                            "und die Option „PDF“ aktiviert ist.",
                        )
                if not want_pdf or not opened:
                    try:
                        os.startfile(export_path)
                    except OSError:
                        QMessageBox.information(
                            self,
                            "Vorschau",
                            f"Dateien liegen unter:\n{export_path}",
                        )
                QMessageBox.information(
                    self,
                    "Vorschau",
                    "Die Beispieldokumente wurden erstellt. Vorschau ≠ rechtlicher Prüfstand; "
                    "andere Excel-Zeilen können abweichen.",
                )
            else:
                self.current_file_label.setText("Vorschau mit Fehlern oder ohne Ausgabe.")
                QMessageBox.warning(
                    self,
                    "Vorschau",
                    "Die Vorschau konnte keine Dokumente erzeugen oder ist fehlgeschlagen. Bitte Logs prüfen.",
                )
        elif success and self.last_export_path:
            self.current_file_label.setText("Verarbeitung abgeschlossen!")
            self.update_export_action_buttons(False)
            QMessageBox.information(self, "Fertig", "Alle Dokumente wurden erfolgreich erstellt.")
        else:
            self.current_file_label.setText("Verarbeitung mit Fehlern abgeschlossen.")
            QMessageBox.warning(
                self,
                "Fehler",
                "Die Verarbeitung wurde mit Fehlern abgeschlossen. Bitte Logs prüfen."
            )
        self.update_export_action_buttons(False)
        if template_row_preview:
            self._template_row_preview_hide()
    def update_current_file_label(self, filename):
        """Aktualisiert das Label, das den Namen der aktuell verarbeiteten Datei anzeigt."""
        self.current_file_label.setText(f"Verarbeite: {filename}...")

    def update_progress_bar(self, current, total):
        """Aktualisiert die Fortschrittsanzeige."""
        if total > 0:
            self.progress.setValue(int(current / total * 100))
        else:
            self.progress.setValue(0)

    def _show_template_row_preview_ui(self, vorlage_name):
        """Fortschritt im Tab „Vorlagen“ für die Einzel-Vorschau (sichtbar ohne Wechsel zur Hauptsteuerung)."""
        self._template_row_preview_active = True
        if hasattr(self, '_template_preview_status'):
            self._template_preview_status.setVisible(True)
            self._template_preview_status.setText(
                f"Vorschau läuft … ({vorlage_name}) — bitte warten, Word/Excel kann kurz blockieren."
            )
        if hasattr(self, '_template_preview_progress'):
            self._template_preview_progress.setVisible(True)
            self._template_preview_progress.setRange(0, 0)
            self._template_preview_progress.setFormat("Vorschau läuft …")

    def _template_row_preview_hide(self):
        """Blendet die Vorlagen-Vorschau-Fortschrittszeile aus."""
        self._template_row_preview_active = False
        if hasattr(self, '_template_preview_progress'):
            self._template_preview_progress.setRange(0, 100)
            self._template_preview_progress.setValue(0)
            self._template_preview_progress.setFormat("%p%")
            self._template_preview_progress.setVisible(False)
        if hasattr(self, '_template_preview_status'):
            self._template_preview_status.setVisible(False)
            self._template_preview_status.setText("")

    def _on_template_row_preview_progress(self, current, total):
        if not getattr(self, '_template_row_preview_active', False):
            return
        if not hasattr(self, '_template_preview_progress'):
            return
        if total and total > 0:
            self._template_preview_progress.setRange(0, 100)
            self._template_preview_progress.setFormat("%p%")
            self._template_preview_progress.setValue(min(100, int(100 * current / total)))
        else:
            self._template_preview_progress.setRange(0, 0)
            self._template_preview_progress.setFormat("Vorschau läuft …")

    def _on_template_row_preview_file(self, filename):
        if not getattr(self, '_template_row_preview_active', False):
            return
        if hasattr(self, '_template_preview_status'):
            self._template_preview_status.setText(f"Aktuell: {filename}")
