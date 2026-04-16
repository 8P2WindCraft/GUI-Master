import os
import sys

from PySide6.QtCore import Qt

from core.logging import _log_handler


def resource_path(relative_path):
    """Gibt den absoluten Pfad zu einer Ressource zurück, auch im PyInstaller-EXE-Modus."""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)


def svg_to_png_file_pyside(svg_path, png_path, log_callback=None, scale=3, compression=-1):
    """
    Konvertiert eine SVG-Datei in eine PNG-Datei unter Verwendung von PySide6.
    Dies ist nützlich, da python-docx SVG-Bilder nicht direkt einbetten kann.

    Args:
        svg_path (str): Pfad zur SVG-Quelldatei.
        png_path (str): Pfad zur PNG-Zieldatei.
        log_callback (callable, optional): Callback für Log-Nachrichten.
        scale (int): Skalierungsfaktor für die Auflösung der PNG-Datei.
        compression (int): PNG-Kompressionslevel (-1 für Standard).

    Returns:
        bool: True bei Erfolg, False bei einem Fehler.
    """
    try:
        from PySide6.QtSvg import QSvgRenderer
        from PySide6.QtGui import QImage, QPainter
        from PySide6.QtCore import QByteArray, QSize

        if not os.path.exists(svg_path):
            _log_handler(f"SVG-Datei nicht gefunden: {svg_path}", "ERROR", log_callback)
            return False
        with open(svg_path, 'rb') as f:
            svg_data = f.read()
        renderer = QSvgRenderer(QByteArray(svg_data))
        if not renderer.isValid():
            _log_handler(f"SVG-Datei ist ungültig: {svg_path}", "ERROR", log_callback)
            return False
        size = renderer.defaultSize()
        if size.isEmpty():
            size.setWidth(300)
            size.setHeight(150)
        scaled_size = QSize(int(size.width() * scale), int(size.height() * scale))
        image = QImage(scaled_size, QImage.Format_ARGB32)
        image.fill(Qt.transparent)
        painter = QPainter(image)
        renderer.render(painter)
        painter.end()
        if not image.save(png_path, "PNG", compression):
            _log_handler(f"Speichern der PNG-Datei fehlgeschlagen für: {png_path}", "ERROR", log_callback)
            return False
        return True
    except Exception as e:
        _log_handler(f"FEHLER bei SVG-Konvertierung: {e}", "ERROR", log_callback)
        return False
