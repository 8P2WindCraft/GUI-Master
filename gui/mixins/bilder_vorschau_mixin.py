"""Bilder-Vorschau-Tab; Tab-Wechsel lädt Vorschau und aktualisiert andere Tabs."""
import os

from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ..constants import BILDER_VORSCHAU_EXTENSIONS


class BilderVorschauMixin:
    """Dateiliste und Bildvorschau; :meth:`_on_tab_changed` orchestriert Neuladen."""

    def setup_bilder_vorschau_tab(self, tab):
        """Erstellt den Inhalt des 'Bilder-Vorschau'-Tabs: Dateiliste des Bilder-Ordners und Bildvorschau."""
        layout = QVBoxLayout(tab)

        self.bilder_vorschau_hint_label = QLabel("Bitte in der Hauptsteuerung einen Bilder-Ordner wählen.")
        self.bilder_vorschau_hint_label.setStyleSheet("color: #666; padding: 8px;")
        layout.addWidget(self.bilder_vorschau_hint_label)

        splitter = QSplitter(Qt.Horizontal)

        self.bilder_vorschau_table = QTableWidget()
        self.bilder_vorschau_table.setColumnCount(2)
        self.bilder_vorschau_table.setHorizontalHeaderLabels(["Dateiname", "Dateityp"])
        self.bilder_vorschau_table.horizontalHeader().setStretchLastSection(True)
        self.bilder_vorschau_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.bilder_vorschau_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.bilder_vorschau_table.itemSelectionChanged.connect(self._on_bilder_vorschau_selection_changed)
        splitter.addWidget(self.bilder_vorschau_table)

        preview_frame = QFrame()
        preview_frame.setFrameShape(QFrame.StyledPanel)
        preview_layout = QVBoxLayout(preview_frame)
        preview_layout.addWidget(QLabel("Vorschau:"))
        self.bilder_vorschau_preview_label = QLabel("Keine Vorschau")
        self.bilder_vorschau_preview_label.setAlignment(Qt.AlignCenter)
        self.bilder_vorschau_preview_label.setMinimumSize(200, 200)
        self.bilder_vorschau_preview_label.setMaximumSize(400, 400)
        self.bilder_vorschau_preview_label.setStyleSheet("border: 1px solid #ccc; background: #f5f5f5; padding: 8px;")
        self.bilder_vorschau_preview_label.setScaledContents(False)
        preview_layout.addWidget(self.bilder_vorschau_preview_label)
        preview_layout.addStretch()
        splitter.addWidget(preview_frame)

        splitter.setSizes([400, 400])
        layout.addWidget(splitter)

    def _on_tab_changed(self, index):
        """Lädt die Bilder-Vorschau-Liste neu, wenn der Bilder-Vorschau-Tab aktiv wird."""
        if hasattr(self, 'bilder_vorschau_tab_index') and index == self.bilder_vorschau_tab_index:
            self.refresh_bilder_vorschau()
        if hasattr(self, 'vorlagen_tab_index') and index == self.vorlagen_tab_index:
            self.refresh_template_checkboxes()
        if hasattr(self, 'rules_tab_index') and index == self.rules_tab_index:
            self.refresh_rules_template_combos()
        if hasattr(self, 'pdf_tab_index') and index == self.pdf_tab_index:
            self.refresh_pdf_tab()

    def refresh_bilder_vorschau(self):
        """Liest den Bilder-Ordner ein und füllt die Tabelle; setzt die Vorschau zurück."""
        if not hasattr(self, 'bilder_vorschau_table'):
            return
        path = (self.bilder_ordner_edit.text() if hasattr(self, 'bilder_ordner_edit') else "").strip()
        self.bilder_vorschau_table.setRowCount(0)
        self.bilder_vorschau_preview_label.clear()
        self.bilder_vorschau_preview_label.setText("Keine Vorschau")
        if not path:
            self.bilder_vorschau_hint_label.setVisible(True)
            self.bilder_vorschau_hint_label.setText("Bitte in der Hauptsteuerung einen Bilder-Ordner wählen.")
            return
        if not os.path.isdir(path):
            self.bilder_vorschau_hint_label.setVisible(True)
            self.bilder_vorschau_hint_label.setText(f"Ordner existiert nicht: {path}")
            return
        self.bilder_vorschau_hint_label.setVisible(False)
        try:
            names = sorted(os.listdir(path))
        except OSError:
            self.bilder_vorschau_hint_label.setVisible(True)
            self.bilder_vorschau_hint_label.setText("Ordner konnte nicht gelesen werden.")
            return
        for name in names:
            full = os.path.join(path, name)
            if os.path.isfile(full):
                ext = os.path.splitext(name)[1].lower()
                if ext in BILDER_VORSCHAU_EXTENSIONS:
                    row = self.bilder_vorschau_table.rowCount()
                    self.bilder_vorschau_table.insertRow(row)
                    self.bilder_vorschau_table.setItem(row, 0, QTableWidgetItem(name))
                    self.bilder_vorschau_table.setItem(row, 1, QTableWidgetItem(ext or "—"))
                    self.bilder_vorschau_table.item(row, 0).setData(Qt.UserRole, full)
                    self.bilder_vorschau_table.item(row, 1).setData(Qt.UserRole, full)

    def _on_bilder_vorschau_selection_changed(self):
        """Zeigt bei Auswahl einer Zeile die Bildvorschau an."""
        if not hasattr(self, 'bilder_vorschau_preview_label'):
            return
        row = self.bilder_vorschau_table.currentRow()
        if row < 0:
            self.bilder_vorschau_preview_label.clear()
            self.bilder_vorschau_preview_label.setText("Keine Vorschau")
            return
        item = self.bilder_vorschau_table.item(row, 0)
        file_path = item.data(Qt.UserRole) if item else None
        if not file_path or not os.path.isfile(file_path):
            self.bilder_vorschau_preview_label.clear()
            self.bilder_vorschau_preview_label.setText("Keine Vorschau")
            return
        ext = os.path.splitext(file_path)[1].lower()
        if ext == '.svg':
            try:
                from core.utils import svg_to_png_file_pyside
                import tempfile
                fd, tmp_png = tempfile.mkstemp(suffix='.png')
                os.close(fd)
                try:
                    if svg_to_png_file_pyside(file_path, tmp_png, scale=2, compression=-1):
                        pix = QPixmap(tmp_png)
                        if not pix.isNull():
                            scaled = pix.scaled(400, 400, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                            self.bilder_vorschau_preview_label.setPixmap(scaled)
                            self.bilder_vorschau_preview_label.setText("")
                        else:
                            self.bilder_vorschau_preview_label.clear()
                            self.bilder_vorschau_preview_label.setText("SVG konnte nicht geladen werden.")
                    else:
                        self.bilder_vorschau_preview_label.clear()
                        self.bilder_vorschau_preview_label.setText("Keine Vorschau")
                finally:
                    try:
                        os.unlink(tmp_png)
                    except OSError:
                        pass
            except Exception:
                self.bilder_vorschau_preview_label.clear()
                self.bilder_vorschau_preview_label.setText("Keine Vorschau")
        else:
            pix = QPixmap(file_path)
            if not pix.isNull():
                scaled = pix.scaled(400, 400, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                self.bilder_vorschau_preview_label.setPixmap(scaled)
                self.bilder_vorschau_preview_label.setText("")
            else:
                self.bilder_vorschau_preview_label.clear()
                self.bilder_vorschau_preview_label.setText("Keine Vorschau")
