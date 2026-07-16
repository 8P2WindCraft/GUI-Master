import html
import re
from datetime import datetime


def _format_log_html(msg):
    """Escaped Log-Text und erhaelt gezielte HTML-Hervorhebungen."""
    msg_html = html.escape(str(msg)).replace("\n", "<br>")
    # Dateinamen fett machen (z.B. 'template.docx')
    msg_html = re.sub(r"&#x27;([^&#]+?\.\w+)&#x27;", r"<strong>'\1'</strong>", msg_html)
    # Platzhalter-Sets lila machen (z.B. {'var1', 'var2'})
    msg_html = re.sub(r"(\s*\{.*?\})", r'<span style="color: #8A2BE2;">\1</span>', msg_html)
    return msg_html


def _log_handler(msg, level="INFO", log_callback=None, is_dark_mode=False):
    """
    Zentraler Log-Handler, der Nachrichten formatiert und ausgibt.
    Unterstützt zwei Ausgabemodi:
    - GUI (log_callback vorhanden): Formatiert die Nachricht als HTML für eine
      farbige und strukturierte Darstellung in einem QTextEdit-Widget.
    - Konsole (log_callback ist None): Gibt die Nachricht als einfachen Text aus.
    """
    timestamp = datetime.now().strftime('%H:%M:%S')
    level_upper = level.upper()

    # Nachricht immer zuerst für die Konsole formatieren
    plain_message = msg if level_upper == "SEP" else f"[{timestamp}] [{level_upper:^7}] {msg}"
    if not plain_message.endswith('\n'):
        plain_message += '\n'

    if log_callback:
        # GUI-Logging: HTML aus escaped Plaintext plus kontrollierten Hervorhebungen erstellen
        color_map = {
            "INFO": "#0077CC",      # Blau
            "SUCCESS": "#2E8B57",   # Seegrün
            "WARN": "#FFA500",      # Orange
            "ERROR": "#D22B2B",     # Ziegelrot
            "FATAL": "#8B0000",     # Dunkelrot
            "SEP": "#808080"        # Grau
        }
        color = color_map.get(level_upper, "black")

        msg_html = _format_log_html(msg)

        if level_upper == "SEP":
            html_message = f'<div style="display:block; color: {color}; text-align: center; font-family: monospace; margin: 5px 0;">{msg_html}</div>'
        else:
            beschr_color = "#FFF" if is_dark_mode else "#333"
            html_message = (
                f'<br><div style="display:block; font-family: Consolas, Courier New, monospace; line-height: 1.4;">'
                f'<span style="color: #808080;">{timestamp}</span> '
                f'<span style="color: {color}; font-weight: bold;">[{level_upper:^7}]</span> '
                f'<span style="color: {beschr_color};">{msg_html}</span>'
                f'</div>'
            )
        log_callback(html_message)
    else:
        # Reines Konsolen-Logging
        print(plain_message, end='')
