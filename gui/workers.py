"""
Hintergrund-Threads: Generierung (Word/docxtpl) und PDF-Konvertierung.
"""
import inspect
import traceback

from PySide6.QtCore import QThread, Signal

from core.logic import (
    _convert_docx_to_pdf_in_folder,
    verarbeite_vorlagen,
    verarbeite_vorlagen_preview,
)
from core.logging import _log_handler


class Worker(QThread):
    """
    Führt die zeitintensive Dokumentenerstellung in einem separaten Thread aus
    (GUI bleibt bedienbar).

    Signale: log, progress, finished, current_file
    """
    log = Signal(str)  # html_message
    progress = Signal(int, int)
    finished = Signal(bool, str)  # success, export_path
    current_file = Signal(str)

    def __init__(self, **kwargs):
        super().__init__()
        self.params = kwargs
        self.dry_run = kwargs.get('dry_run', False)
        self.preview_run = kwargs.get('preview_run', False)

    def run(self):
        """Startet die Verarbeitung über verarbeite_vorlagen bzw. Vorschau."""
        # Vorschau: verarbeite_vorlagen_preview; sonst voller Lauf oder Trockenlauf (dry_run)
        try:
            thread_safe_params = self.params.copy()
            thread_safe_params['log_callback'] = self.log.emit
            thread_safe_params['progress_callback'] = self.progress.emit
            thread_safe_params['file_callback'] = self.current_file.emit
            thread_safe_params['worker_thread'] = self
            thread_safe_params.pop('theme', None)
            preview_run = thread_safe_params.pop('preview_run', False)
            preview_template_abs = thread_safe_params.pop('preview_template_abs', None)

            def _filter_kwargs_for(func, kwargs):
                """Nur Argumente übergeben, die die Ziel-Funktion tatsächlich akzeptiert."""
                sig = inspect.signature(func)
                return {k: v for k, v in kwargs.items() if k in sig.parameters}

            if preview_run:
                thread_safe_params.pop('dry_run', None)
                filtered_params = _filter_kwargs_for(verarbeite_vorlagen_preview, thread_safe_params)
                result = verarbeite_vorlagen_preview(
                    **filtered_params,
                    preview_template_abs=preview_template_abs,
                )
                if not self.isInterruptionRequested():
                    self.finished.emit(bool(result), result)
                else:
                    self.finished.emit(False, None)
            else:
                filtered_params = _filter_kwargs_for(verarbeite_vorlagen, thread_safe_params)
                result = verarbeite_vorlagen(**filtered_params)

                if self.dry_run:
                    self.finished.emit(result, None)  # Trockenlauf: bool
                elif not self.isInterruptionRequested():
                    self.finished.emit(bool(result), result)  # echter Lauf: Erfolg + Exportpfad
                else:
                    self.finished.emit(False, None)

        except Exception as e:
            error_msg = f"FATALER FEHLER im Worker-Thread: {e}\n{traceback.format_exc()}"
            _log_handler(error_msg, "FATAL", self.log.emit)
            self.finished.emit(False, None)


class PdfWorker(QThread):
    """
    Manuelle PDF-Konvertierung des Export-Ordners (Hintergrund, GUI bleibt responsiv).
    """
    log = Signal(str)  # html_message
    progress = Signal(int, int)
    current_file = Signal(str)
    finished = Signal(object)  # result dict

    def __init__(self, export_path, is_dark_mode=False):
        super().__init__()
        self.export_path = export_path
        self.is_dark_mode = is_dark_mode

    def run(self):
        """Konvertiert .docx im Export-Ordner zu PDF (Implementierung in core.logic)."""
        try:
            def _worker_log(msg, level="INFO"):
                _log_handler(msg, level, self.log.emit, self.is_dark_mode)

            result = _convert_docx_to_pdf_in_folder(
                self.export_path,
                _worker_log,
                progress_callback=self.progress.emit,
                file_callback=self.current_file.emit
            )
            self.finished.emit(result)
        except Exception as e:
            self.finished.emit(
                {
                    "docx_count": 0,
                    "success_count": 0,
                    "failure_count": 0,
                    "error": "unknown_error",
                    "error_message": str(e),
                }
            )
