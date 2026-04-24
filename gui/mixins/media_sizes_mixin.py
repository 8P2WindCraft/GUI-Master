"""Tab Bilder & QR-Größen: Übersicht, Vorschau, Schreiben nach Excel Blatt 2."""
import io
import os
import tempfile

from PySide6.QtCore import Qt
from PySide6.QtGui import QImage, QPixmap
from PySide6.QtWidgets import (
    QCheckBox,
    QDoubleSpinBox,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)

from core.media_sizes import (
    gruppier_nach_basis_platzhalter,
    sammle_media_groessen_uebersicht,
    schreibe_groessen_nach_excel_blatt2,
)


class MediaSizesMixin:
    """Größen zu Bild- und QR-Platzhaltern, Tabellen-UI, Speichern ins Excel-Fallbackblatt."""

    def setup_media_sizes_tab(self, tab):
        """Tab: Bild-/QR-Platzhalter aus Vorlagen, Größen aus Excel, Vorschau, Schreiben nach Blatt 2."""
        layout = QVBoxLayout(tab)
        intro = QLabel(
            "Standard: Gruppierung nach Basis-Platzhalter – eine Größenänderung gilt für alle gelisteten "
            "Vorlagen. Referenz und Größe wie erste Anlagenzeile + Fallback (Blatt 2). "
            "Ohne Häkchen: jede Vorlage einzeln. Speichern schreibt nach Blatt 2."
        )
        intro.setWordWrap(True)
        intro.setStyleSheet("color: #555;")
        layout.addWidget(intro)

        btn_row = QHBoxLayout()
        self.media_sizes_refresh_btn = QPushButton("Aktualisieren")
        self.media_sizes_refresh_btn.setToolTip("Excel und Vorlagen neu einlesen")
        self.media_sizes_refresh_btn.clicked.connect(self.refresh_media_sizes_table)
        btn_row.addWidget(self.media_sizes_refresh_btn)
        self.media_sizes_save_btn = QPushButton("Änderungen in Excel speichern")
        self.media_sizes_save_btn.setToolTip("Nur geänderte cm-Werte → Blatt 2, Spalte B/C")
        self.media_sizes_save_btn.clicked.connect(self._save_media_sizes_excel)
        btn_row.addWidget(self.media_sizes_save_btn)
        self.media_sizes_group_check = QCheckBox("Nach Basisplatzhalter gruppieren (Standard)")
        self.media_sizes_group_check.setChecked(True)
        self.media_sizes_group_check.setToolTip(
            "Eine Zeile pro Platzhalter (z. B. ba_logo_img). Größe und Referenz gelten für alle "
            "gelisteten Vorlagen gleich – Speichern schreibt eine Textmarke in Blatt 2."
        )
        self.media_sizes_group_check.toggled.connect(self._on_media_sizes_group_toggled)
        btn_row.addWidget(self.media_sizes_group_check)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        self.media_sizes_status_label = QLabel("")
        self.media_sizes_status_label.setStyleSheet("color: #666; font-size: 11px;")
        layout.addWidget(self.media_sizes_status_label)

        self.media_sizes_detail_table = QTableWidget()
        self.media_sizes_detail_table.setColumnCount(8)
        self.media_sizes_detail_table.setHorizontalHeaderLabels(
            [
                "Dokument",
                "Basis-Platzhalter",
                "Typ",
                "Referenz",
                "Vorschau",
                "Größen-Textmarke",
                "Wert (cm)",
                "Quelle",
            ]
        )
        self.media_sizes_detail_table.horizontalHeader().setStretchLastSection(True)
        self.media_sizes_detail_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.media_sizes_detail_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.media_sizes_detail_table.hide()
        layout.addWidget(self.media_sizes_detail_table)

        self.media_sizes_group_table = QTableWidget()
        self.media_sizes_group_table.setColumnCount(8)
        self.media_sizes_group_table.setHorizontalHeaderLabels(
            [
                "Basis-Platzhalter",
                "Typ",
                "Referenz",
                "Vorschau",
                "Größen-Textmarke",
                "Wert (cm)",
                "Quelle",
                "Dokumente (Vorlagen)",
            ]
        )
        self.media_sizes_group_table.horizontalHeader().setStretchLastSection(True)
        self.media_sizes_group_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.media_sizes_group_table.setEditTriggers(QTableWidget.NoEditTriggers)
        layout.addWidget(self.media_sizes_group_table)

        self._media_sizes_detail_rows = []
        self._media_sizes_row_meta = []
        self._media_sizes_basis_row_meta = []

    def _on_media_sizes_group_toggled(self, checked: bool):
        self.media_sizes_detail_table.setVisible(not checked)
        self.media_sizes_group_table.setVisible(checked)
        if checked and self._media_sizes_detail_rows:
            self._fill_media_sizes_basis_group_table()
        elif not checked and self._media_sizes_detail_rows:
            self._fill_media_sizes_detail_table(self._media_sizes_detail_rows)

    def _media_sizes_preview_pixmap(self, kind: str, reference: str, bilder_ordner: str) -> QPixmap | None:
        ref = (reference or "").strip()
        if not ref:
            pix = QPixmap(56, 56)
            pix.fill(Qt.lightGray)
            return pix
        if kind == "Bild" and bilder_ordner:
            full = os.path.normpath(os.path.join(bilder_ordner, ref))
            if not os.path.isfile(full):
                return None
            ext = os.path.splitext(full)[1].lower()
            if ext == ".svg":
                try:
                    from core.utils import svg_to_png_file_pyside

                    fd, tmp_png = tempfile.mkstemp(suffix=".png")
                    os.close(fd)
                    try:
                        if svg_to_png_file_pyside(full, tmp_png, scale=2, compression=-1):
                            pix = QPixmap(tmp_png)
                        else:
                            return None
                    finally:
                        try:
                            os.unlink(tmp_png)
                        except OSError:
                            pass
                except Exception:
                    return None
            else:
                pix = QPixmap(full)
            if pix.isNull():
                return None
            return pix.scaled(56, 56, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        if kind == "QR-Link":
            try:
                import qrcode

                qr = qrcode.make(ref)
                buf = io.BytesIO()
                qr.save(buf, format="PNG")
                buf.seek(0)
                img = QImage.fromData(buf.read())
                if img.isNull():
                    return None
                return QPixmap.fromImage(img).scaled(
                    56, 56, Qt.KeepAspectRatio, Qt.SmoothTransformation
                )
            except Exception:
                return None
        return None

    def _fill_media_sizes_detail_table(self, rows: list):
        self._media_sizes_row_meta = []
        table = self.media_sizes_detail_table
        table.setRowCount(0)
        bilder_ordner = (self.paths.get("bilder_ordner") or "").strip()
        for i, r in enumerate(rows):
            table.insertRow(i)
            table.setItem(i, 0, QTableWidgetItem(r.get("template_rel", "")))
            table.setItem(i, 1, QTableWidgetItem(r.get("base_key", "")))
            table.setItem(i, 2, QTableWidgetItem(r.get("kind", "")))
            ref_full = r.get("reference") or ""
            ref_disp = ref_full if len(ref_full) <= 100 else ref_full[:97] + "…"
            it_ref = QTableWidgetItem(ref_disp)
            it_ref.setToolTip(ref_full if len(ref_full) > 100 else "")
            table.setItem(i, 3, it_ref)
            table.setItem(i, 5, QTableWidgetItem(r.get("size_key", "")))
            spin = QDoubleSpinBox()
            spin.setRange(0.1, 999.0)
            spin.setDecimals(2)
            spin.setValue(float(r.get("cm_value", 0)))
            spin.setEnabled(not r.get("cm_variable"))
            table.setCellWidget(i, 6, spin)
            src = r.get("cm_source", "")
            it_src = QTableWidgetItem(src)
            if r.get("cm_variable"):
                it_src.setToolTip("Werte in Blatt 1 unterscheiden sich – bitte in Excel anpassen.")
            table.setItem(i, 7, it_src)

            lbl = QLabel()
            lbl.setAlignment(Qt.AlignCenter)
            pix = self._media_sizes_preview_pixmap(r.get("kind", ""), ref_full, bilder_ordner)
            if pix is not None and not pix.isNull():
                lbl.setPixmap(pix)
            else:
                lbl.setText("—")
                if r.get("kind") == "Bild" and ref_full and bilder_ordner:
                    lbl.setToolTip("Bilddatei im Bilder-Ordner nicht gefunden.")
            lbl.setFixedSize(60, 60)
            table.setCellWidget(i, 4, lbl)
            table.setRowHeight(i, 68)

            self._media_sizes_row_meta.append(
                {
                    "size_key": r.get("size_key", ""),
                    "cm_variable": bool(r.get("cm_variable")),
                    "orig_cm": float(r.get("cm_value", 0)),
                }
            )

    def _fill_media_sizes_basis_group_table(self):
        detail = self._media_sizes_detail_rows
        groups = gruppier_nach_basis_platzhalter(detail)
        t = self.media_sizes_group_table
        t.setRowCount(0)
        self._media_sizes_basis_row_meta = []
        bilder_ordner = (self.paths.get("bilder_ordner") or "").strip()
        for i, g in enumerate(groups):
            t.insertRow(i)
            base_key = g.get("base_key", "")
            t.setItem(i, 0, QTableWidgetItem(base_key))
            t.setItem(i, 1, QTableWidgetItem(g.get("kind", "")))

            ref_full = g.get("reference") or ""
            ref_disp = ref_full if len(ref_full) <= 100 else ref_full[:97] + "…"
            it_ref = QTableWidgetItem(ref_disp)
            tip_ref = ref_full
            if g.get("reference_conflict"):
                tip_ref = (tip_ref + "\n\n" if tip_ref else "") + "Hinweis: Unterschiedliche Referenzwerte je Vorlage – bitte Excel prüfen."
            it_ref.setToolTip(tip_ref)
            t.setItem(i, 2, it_ref)

            lbl = QLabel()
            lbl.setAlignment(Qt.AlignCenter)
            pix = self._media_sizes_preview_pixmap(g.get("kind", ""), ref_full, bilder_ordner)
            if pix is not None and not pix.isNull():
                lbl.setPixmap(pix)
            else:
                lbl.setText("—")
            lbl.setFixedSize(60, 60)
            t.setCellWidget(i, 3, lbl)

            t.setItem(i, 4, QTableWidgetItem(g.get("size_key", "")))

            spin = QDoubleSpinBox()
            spin.setRange(0.1, 999.0)
            spin.setDecimals(2)
            spin.setValue(float(g.get("cm_value", 0)))
            spin.setEnabled(not g.get("cm_variable"))
            t.setCellWidget(i, 5, spin)

            it_src = QTableWidgetItem(g.get("cm_source", ""))
            if g.get("cm_variable"):
                it_src.setToolTip("Werte in Excel weichen ab – bitte zuerst in Excel bereinigen oder Detailansicht nutzen.")
            t.setItem(i, 6, it_src)

            doc_compact = g.get("templates_compact", "")
            it_doc = QTableWidgetItem(doc_compact)
            it_doc.setToolTip(g.get("templates_tooltip") or doc_compact)
            t.setItem(i, 7, it_doc)

            t.setRowHeight(i, 68)

            self._media_sizes_basis_row_meta.append(
                {
                    "size_key": g.get("size_key", ""),
                    "cm_variable": bool(g.get("cm_variable")),
                    "orig_cm": float(g.get("cm_value", 0)),
                }
            )

    def refresh_media_sizes_table(self):
        vo = (self.paths.get("vorlagen_ordner") or "").strip()
        excel_path = (self.paths.get("excel_path") or "").strip()
        hr = self.settings.get("header_row", 3)
        sel = self._template_paths_for_worker()

        if not excel_path or not os.path.isfile(excel_path):
            self._media_sizes_detail_rows = []
            if self.media_sizes_group_check.isChecked():
                self._fill_media_sizes_basis_group_table()
            else:
                self._fill_media_sizes_detail_table([])
            self.media_sizes_status_label.setText("Bitte eine Excel-Datei unter „Hauptsteuerung“ wählen.")
            return
        if not vo or not os.path.isdir(vo):
            self._media_sizes_detail_rows = []
            if self.media_sizes_group_check.isChecked():
                self._fill_media_sizes_basis_group_table()
            else:
                self._fill_media_sizes_detail_table([])
            self.media_sizes_status_label.setText("Bitte einen Vorlagen-Ordner wählen.")
            return
        if sel == []:
            self._media_sizes_detail_rows = []
            if self.media_sizes_group_check.isChecked():
                self._fill_media_sizes_basis_group_table()
            else:
                self._fill_media_sizes_detail_table([])
            self.media_sizes_status_label.setText("Keine Vorlagen ausgewählt (Tab „Vorlagen“).")
            return

        try:
            rows = sammle_media_groessen_uebersicht(vo, excel_path, hr, sel, None)
        except Exception as e:
            self._media_sizes_detail_rows = []
            if self.media_sizes_group_check.isChecked():
                self._fill_media_sizes_basis_group_table()
            else:
                self._fill_media_sizes_detail_table([])
            self.media_sizes_status_label.setText(f"Fehler beim Einlesen: {e}")
            QMessageBox.warning(self, "Bilder & QR-Größen", str(e))
            return

        self._media_sizes_detail_rows = rows
        n_basis = len(gruppier_nach_basis_platzhalter(rows)) if rows else 0
        self.media_sizes_status_label.setText(
            f"{len(rows)} Einträge (Vorlagen × Platzhalter), {n_basis} Basis-Platzhalter."
        )
        if self.media_sizes_group_check.isChecked():
            self._fill_media_sizes_basis_group_table()
        else:
            self._fill_media_sizes_detail_table(rows)

    def _save_media_sizes_excel(self):
        excel_path = (self.paths.get("excel_path") or "").strip()
        if not excel_path or not os.path.isfile(excel_path):
            QMessageBox.warning(self, "Excel", "Keine gültige Excel-Datei.")
            return

        updates = []
        if self.media_sizes_group_check.isChecked():
            table = self.media_sizes_group_table
            for i in range(table.rowCount()):
                if i >= len(self._media_sizes_basis_row_meta):
                    break
                meta = self._media_sizes_basis_row_meta[i]
                if meta.get("cm_variable"):
                    continue
                w = table.cellWidget(i, 5)
                if isinstance(w, QDoubleSpinBox) and abs(w.value() - meta["orig_cm"]) > 1e-9:
                    updates.append((meta["size_key"], w.value()))
        else:
            table = self.media_sizes_detail_table
            for i in range(table.rowCount()):
                if i >= len(self._media_sizes_row_meta):
                    break
                meta = self._media_sizes_row_meta[i]
                if meta.get("cm_variable"):
                    continue
                w = table.cellWidget(i, 6)
                if isinstance(w, QDoubleSpinBox) and abs(w.value() - meta["orig_cm"]) > 1e-9:
                    updates.append((meta["size_key"], w.value()))

        if not updates:
            QMessageBox.information(self, "Excel", "Keine geänderten Größen zum Speichern.")
            return

        try:
            schreibe_groessen_nach_excel_blatt2(excel_path, updates)
        except PermissionError:
            QMessageBox.warning(
                self,
                "Excel",
                "Die Datei ist gesperrt (z. B. in Excel geöffnet).",
            )
            return
        except Exception as e:
            QMessageBox.warning(self, "Excel", f"Speichern fehlgeschlagen:\n{e}")
            return

        QMessageBox.information(self, "Excel", f"{len(updates)} Größe(n) wurden in Blatt 2 gespeichert.")
        self.refresh_media_sizes_table()
        self.show_excel_data()

