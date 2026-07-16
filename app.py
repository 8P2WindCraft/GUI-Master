import multiprocessing
import os
import sys
import traceback

from PySide6.QtCore import QTimer, Qt
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication

from core.utils import resource_path
from gui.main_window import MainWindow


def _configure_utf8_stdio():
    """Erzwingt UTF-8 für Konsole/Logs, um Mojibake unter Windows zu vermeiden."""
    if sys.platform != "win32":
        return
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        if stream is None:
            continue
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            try:
                reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


def main():
    """Hauptfunktion, die die QApplication startet und das Hauptfenster anzeigt."""
    # Für PyInstaller + multiprocessing (parallele Word-Jobs) unter Windows
    if sys.platform == "win32":
        multiprocessing.freeze_support()
    _configure_utf8_stdio()
    try:
        app = QApplication(sys.argv)
        logo_path = resource_path('Pictures/Logo.png')
        if os.path.exists(logo_path):
            app.setWindowIcon(QIcon(logo_path))
        win = MainWindow()
        win.show()
        # Fenster in den Vordergrund (Windows: oft nötig, damit nicht Cursor/IDE darüber bleibt)
        win.setWindowState((win.windowState() & ~Qt.WindowMinimized) | Qt.WindowActive)
        win.raise_()
        win.activateWindow()

        def _bring_main_window_to_front():
            win.raise_()
            win.activateWindow()

        QTimer.singleShot(0, _bring_main_window_to_front)
        QTimer.singleShot(150, _bring_main_window_to_front)
        sys.exit(app.exec())
    except Exception:
        print(f"Kritischer Fehler:\n{traceback.format_exc()}", file=sys.stderr)


if __name__ == '__main__':
    main()
