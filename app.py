import multiprocessing
import os
import sys
import traceback

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication

from core.utils import resource_path
from gui.main_window import MainWindow


def main():
    """Hauptfunktion, die die QApplication startet und das Hauptfenster anzeigt."""
    # Für PyInstaller + multiprocessing (parallele Word-Jobs) unter Windows
    if sys.platform == "win32":
        multiprocessing.freeze_support()
    try:
        app = QApplication(sys.argv)
        logo_path = resource_path('Pictures/Logo.png')
        if os.path.exists(logo_path):
            app.setWindowIcon(QIcon(logo_path))
        win = MainWindow()
        win.show()
        sys.exit(app.exec())
    except Exception:
        print(f"Kritischer Fehler:\n{traceback.format_exc()}", file=sys.stderr)


if __name__ == '__main__':
    main()
