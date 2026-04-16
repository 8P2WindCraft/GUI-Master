import os
import sys
import traceback

from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication

from core.utils import resource_path
from gui.main_window import MainWindow


def main():
    """Hauptfunktion, die die QApplication startet und das Hauptfenster anzeigt."""
    # #region agent log
    log_path = os.path.join(os.path.dirname(__file__), '.cursor', 'debug.log')
    try:
        with open(log_path, 'a', encoding='utf-8') as f:
            import json
            f.write(json.dumps({"sessionId":"debug-session","runId":"run1","hypothesisId":"C","location":"app.py:14","message":"App start","data":{},"timestamp":int(__import__('time').time()*1000)}) + '\n')
    except: pass
    # #endregion
    try:
        app = QApplication(sys.argv)
        # #region agent log
        try:
            with open(log_path, 'a', encoding='utf-8') as f:
                import json
                f.write(json.dumps({"sessionId":"debug-session","runId":"run1","hypothesisId":"C","location":"app.py:23","message":"QApplication created","data":{},"timestamp":int(__import__('time').time()*1000)}) + '\n')
        except: pass
        # #endregion
        logo_path = resource_path('Pictures/Logo.png')
        # #region agent log
        try:
            with open(log_path, 'a', encoding='utf-8') as f:
                import json
                f.write(json.dumps({"sessionId":"debug-session","runId":"run1","hypothesisId":"A","location":"app.py:28","message":"Logo path resolved","data":{"logo_path":logo_path,"exists":os.path.exists(logo_path)},"timestamp":int(__import__('time').time()*1000)}) + '\n')
        except: pass
        # #endregion
        if os.path.exists(logo_path):
            app.setWindowIcon(QIcon(logo_path))
        win = MainWindow()
        # #region agent log
        try:
            with open(log_path, 'a', encoding='utf-8') as f:
                import json
                f.write(json.dumps({"sessionId":"debug-session","runId":"run1","hypothesisId":"C","location":"app.py:35","message":"MainWindow created","data":{},"timestamp":int(__import__('time').time()*1000)}) + '\n')
        except: pass
        # #endregion
        win.show()
        sys.exit(app.exec())
    except Exception as e:
        # #region agent log
        try:
            with open(log_path, 'a', encoding='utf-8') as f:
                import json
                f.write(json.dumps({"sessionId":"debug-session","runId":"run1","hypothesisId":"C","location":"app.py:42","message":"Exception caught","data":{"error":str(e)},"timestamp":int(__import__('time').time()*1000)}) + '\n')
        except: pass
        # #endregion
        print(f"Kritischer Fehler:\n{traceback.format_exc()}", file=sys.stderr)


if __name__ == '__main__':
    main()
