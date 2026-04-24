"""Tab Regeln: optionale Beschilderungs-Regeln (Vorlage-Zweige pro Excel-Zeile)."""
import os

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListWidget,
    QPushButton,
    QSpinBox,
    QSplitter,
    QTableWidget,
    QVBoxLayout,
    QWidget,
)

from core.logic import liste_textmarken_aus_docx, sammle_vorlagen_pfade

from ..constants import RULE_CONDITION_CHOICES


class SignageRulesMixin:
    """Editor für ``signage_rules``, Textmarken aus Bezugs-Vorlagen, Bedingungen."""

    def setup_rules_tab(self, tab):
        """Tab: Optionale Beschilderungs-Regeln (z. B. Notausgang: Text vs. Bild-Zweig)."""
        layout = QVBoxLayout(tab)
        intro = QLabel(
            "Wenn aktiviert, werden nur für passende Excel-Zeilen die angegebenen Vorlagen erzeugt "
            "(Anlagen-Vorlagen). Ohne Aktivierung bleibt alles wie bisher."
        )
        intro.setWordWrap(True)
        intro.setStyleSheet("color: #555;")
        layout.addWidget(intro)

        self.rules_enabled_check = QCheckBox("Regeln für optionale Schilder aktivieren")
        self.rules_enabled_check.setChecked(self.rules_enabled)
        self.rules_enabled_check.stateChanged.connect(self.save_all_settings)
        layout.addWidget(self.rules_enabled_check)

        self.reuse_lageplan_from_last_export_check = QCheckBox(
            "Lagepläne (rl_) aus letztem Export übernehmen"
        )
        self.reuse_lageplan_from_last_export_check.setChecked(self.reuse_lageplan_from_last_export)
        self.reuse_lageplan_from_last_export_check.stateChanged.connect(self.save_all_settings)
        layout.addWidget(self.reuse_lageplan_from_last_export_check)
        lageplan_hint = QLabel(
            "Wenn aktiv, werden Anlagen-Vorlagen mit Präfix „rl_“ nicht neu gerendert, sondern aus dem "
            "zuletzt erfolgreichen Export kopiert (Unterordner wie in „Kategorien“ für rl_, sonst "
            "„Lageplan“ oder „Plan“). Fehlt die Datei, wird wie gewohnt aus der Vorlage erzeugt."
        )
        lageplan_hint.setWordWrap(True)
        lageplan_hint.setStyleSheet("color: #666; font-size: 11px;")
        layout.addWidget(lageplan_hint)

        self._rules_loading_ui = False
        self._rules_prev_selected_row = -2

        split = QSplitter(Qt.Horizontal)
        left_panel = QWidget()
        left_l = QVBoxLayout(left_panel)
        left_l.setContentsMargins(0, 0, 0, 0)
        self.rules_list_widget = QListWidget()
        self.rules_list_widget.setMinimumWidth(170)
        self.rules_list_widget.currentRowChanged.connect(self._on_rules_list_row_changed)
        left_l.addWidget(self.rules_list_widget)
        list_btn_row = QHBoxLayout()
        rules_add_btn = QPushButton("Neue Regel")
        rules_add_btn.clicked.connect(self._rules_add_clicked)
        rules_del_btn = QPushButton("Löschen")
        rules_del_btn.clicked.connect(self._rules_delete_clicked)
        list_btn_row.addWidget(rules_add_btn)
        list_btn_row.addWidget(rules_del_btn)
        left_l.addLayout(list_btn_row)

        self.rules_editor_widget = QWidget()
        box_l = QVBoxLayout(self.rules_editor_widget)
        self.rules_editor_box = QGroupBox("Regel bearbeiten")
        editor_inner = QVBoxLayout(self.rules_editor_box)

        self.rules_rule_enabled_check = QCheckBox("Diese Regel anwenden")
        self.rules_rule_enabled_check.setChecked(True)
        self.rules_rule_enabled_check.stateChanged.connect(self._rules_mark_dirty_save)
        editor_inner.addWidget(self.rules_rule_enabled_check)

        row_name = QHBoxLayout()
        row_name.addWidget(QLabel("Name:"))
        self.rules_rule_name_edit = QLineEdit()
        self.rules_rule_name_edit.setPlaceholderText("z. B. Notausgang")
        self.rules_rule_name_edit.editingFinished.connect(self._rules_mark_dirty_save)
        self.rules_rule_name_edit.editingFinished.connect(self._rules_sync_list_item_title)
        row_name.addWidget(self.rules_rule_name_edit)
        editor_inner.addLayout(row_name)

        ref_row = QHBoxLayout()
        ref_row.addWidget(QLabel("1. Bezugs-Vorlage (Textmarken):"))
        self.rules_reference_template_combo = QComboBox()
        self.rules_reference_template_combo.setMinimumWidth(280)
        self.rules_reference_template_combo.setToolTip(
            "Zuerst die Word-Vorlage wählen, aus der die Textmarken gelesen werden."
        )
        self.rules_reference_template_combo.currentIndexChanged.connect(self._rules_on_reference_template_changed)
        ref_row.addWidget(self.rules_reference_template_combo)
        editor_inner.addLayout(ref_row)

        ref_hint = QLabel(
            "Nach der Auswahl erscheinen die Platzhalter der Vorlage in der Spalte „Textmarke“. "
            "Leer lassen = Excel-Spaltenname ist gleichzeitig der Kontext-Schlüssel."
        )
        ref_hint.setWordWrap(True)
        ref_hint.setStyleSheet("color: #666; font-size: 11px;")
        editor_inner.addWidget(ref_hint)

        editor_inner.addWidget(QLabel("2. Excel-Spalte und zugehörige Textmarke (mind. eine Zeile mit Inhalt):"))
        self.rules_fields_table = QTableWidget(0, 2)
        self.rules_fields_table.setHorizontalHeaderLabels(["Excel-Spalte", "Textmarke"])
        self.rules_fields_table.horizontalHeader().setStretchLastSection(True)
        self.rules_fields_table.setMinimumHeight(120)
        self.rules_fields_table.setMaximumHeight(220)
        editor_inner.addWidget(self.rules_fields_table)
        field_btn_row = QHBoxLayout()
        rules_field_add_btn = QPushButton("Zeile hinzufügen")
        rules_field_add_btn.clicked.connect(self._rules_field_row_add)
        rules_field_del_btn = QPushButton("Zeile entfernen")
        rules_field_del_btn.clicked.connect(self._rules_field_row_remove)
        field_btn_row.addWidget(rules_field_add_btn)
        field_btn_row.addWidget(rules_field_del_btn)
        field_btn_row.addStretch()
        editor_inner.addLayout(field_btn_row)

        cond_row = QHBoxLayout()
        cond_row.addWidget(QLabel("3. Bedingung (für die Zeilen oben, mindestens eine muss passen):"))
        self.rules_condition_combo = QComboBox()
        self.rules_condition_combo.setMinimumWidth(280)
        for val, label in RULE_CONDITION_CHOICES:
            self.rules_condition_combo.addItem(label, val)
        self.rules_condition_combo.currentIndexChanged.connect(self._rules_on_condition_mode_changed)
        cond_row.addWidget(self.rules_condition_combo)
        editor_inner.addLayout(cond_row)

        param_row = QHBoxLayout()
        param_row.addWidget(QLabel("Zahl / Länge:"))
        self.rules_min_len_spin = QSpinBox()
        self.rules_min_len_spin.setRange(0, 9999)
        self.rules_min_len_spin.setValue(3)
        self.rules_min_len_spin.setToolTip("Für „länger als“ und „Mindestlänge (≥)“.")
        self.rules_min_len_spin.valueChanged.connect(self._rules_mark_dirty_save)
        param_row.addWidget(self.rules_min_len_spin)
        param_row.addWidget(QLabel("enthält / gleich:"))
        self.rules_needle_edit = QLineEdit()
        self.rules_needle_edit.setPlaceholderText("Suchtext oder exakter Vergleichswert")
        self.rules_needle_edit.editingFinished.connect(self._rules_mark_dirty_save)
        param_row.addWidget(self.rules_needle_edit)
        editor_inner.addLayout(param_row)

        regex_row = QHBoxLayout()
        regex_row.addWidget(QLabel("Regex:"))
        self.rules_regex_edit = QLineEdit()
        self.rules_regex_edit.setPlaceholderText("z. B. ^[A-Z]{2}-[0-9]+$")
        self.rules_regex_edit.editingFinished.connect(self._rules_mark_dirty_save)
        regex_row.addWidget(self.rules_regex_edit)
        editor_inner.addLayout(regex_row)

        row_img = QHBoxLayout()
        row_img.addWidget(QLabel("Bild-Spalte (optional, Dateiname im Bilder-Ordner):"))
        self.rules_image_col_edit = QLineEdit()
        self.rules_image_col_edit.setPlaceholderText("z. B. notausgang_img")
        self.rules_image_col_edit.editingFinished.connect(self._rules_mark_dirty_save)
        row_img.addWidget(self.rules_image_col_edit)
        editor_inner.addLayout(row_img)

        row_if = QHBoxLayout()
        row_if.addWidget(QLabel("Vorlage wenn Bedingung / Bild-Zweig „ja“:"))
        self.rules_template_if_combo = QComboBox()
        self.rules_template_if_combo.setMinimumWidth(280)
        self.rules_template_if_combo.currentIndexChanged.connect(self._rules_mark_dirty_save)
        row_if.addWidget(self.rules_template_if_combo)
        editor_inner.addLayout(row_if)

        row_else = QHBoxLayout()
        row_else.addWidget(QLabel("Vorlage wenn Bedingung / Bild-Zweig „nein“:"))
        self.rules_template_else_combo = QComboBox()
        self.rules_template_else_combo.setMinimumWidth(280)
        self.rules_template_else_combo.currentIndexChanged.connect(self._rules_mark_dirty_save)
        row_else.addWidget(self.rules_template_else_combo)
        editor_inner.addLayout(row_else)

        hint = QLabel(
            "Beide Ziel-Vorlagen müssen unter „Vorlagen“ angehakt sein. "
            "Die Bezugs-Vorlage dient nur zum Auslesen der Textmarken und kann dieselbe Datei sein wie IF oder ELSE."
        )
        hint.setWordWrap(True)
        hint.setStyleSheet("color: #666; font-size: 11px;")
        editor_inner.addWidget(hint)

        box_l.addWidget(self.rules_editor_box)
        box_l.addStretch()

        split.addWidget(left_panel)
        split.addWidget(self.rules_editor_widget)
        split.setStretchFactor(1, 1)
        layout.addWidget(split)
        layout.addStretch()

        self._rules_on_condition_mode_changed()
        self._apply_signage_rules_to_ui()

    def _rules_mark_dirty_save(self, *args):
        if getattr(self, "_rules_loading_ui", False):
            return
        self.save_all_settings()

    def _rules_sync_list_item_title(self):
        row = self.rules_list_widget.currentRow()
        if row < 0:
            return
        it = self.rules_list_widget.item(row)
        if it:
            name = (self.rules_rule_name_edit.text() or f"Regel {row + 1}").strip()
            it.setText(name)

    def _set_rule_editor_enabled(self, enabled: bool):
        if hasattr(self, "rules_editor_box"):
            self.rules_editor_box.setEnabled(enabled)

    def _default_signage_rule_dict(self):
        return {
            "id": "",
            "name": "Neue Regel",
            "enabled": True,
            "reference_template": "",
            "templates_if": [],
            "templates_else": [],
            "when": {
                "text_condition": {
                    "mode": "not_empty",
                    "min": 3,
                    "needle": "",
                    "equals_value": "",
                    "regex_pattern": "",
                    "fields": [{"column": "", "textmarke": ""}],
                }
            },
            "branch_on_image": {"column": ""},
        }

    def _ensure_signage_rules_length(self, n: int):
        while len(self.signage_rules) < n:
            self.signage_rules.append(self._default_signage_rule_dict())

    def _pick_template_combo(self, cb: QComboBox, rel: str):
        if not rel:
            cb.setCurrentIndex(0)
            return
        idx = cb.findData(rel)
        if idx >= 0:
            cb.setCurrentIndex(idx)
        else:
            cb.addItem(rel, rel)
            cb.setCurrentIndex(cb.count() - 1)

    def _on_rules_list_row_changed(self, row: int):
        if getattr(self, "_rules_loading_ui", False):
            return
        prev = getattr(self, "_rules_prev_selected_row", -2)
        if prev >= 0 and prev != row:
            self._ensure_signage_rules_length(prev + 1)
            self.signage_rules[prev] = self._collect_rule_dict_from_editor()
            it = self.rules_list_widget.item(prev)
            if it:
                it.setText(self.signage_rules[prev].get("name") or f"Regel {prev + 1}")
        self._rules_prev_selected_row = row
        if row < 0:
            self._set_rule_editor_enabled(False)
            return
        self._set_rule_editor_enabled(True)
        self._ensure_signage_rules_length(row + 1)
        self._load_rule_editor_from_index(row)

    def _rules_add_clicked(self):
        row = self.rules_list_widget.currentRow()
        if row >= 0:
            self._ensure_signage_rules_length(row + 1)
            self.signage_rules[row] = self._collect_rule_dict_from_editor()
        self.signage_rules.append(self._default_signage_rule_dict())
        self.rules_list_widget.addItem(self.signage_rules[-1]["name"])
        self.rules_list_widget.setCurrentRow(self.rules_list_widget.count() - 1)
        self.save_all_settings()

    def _rules_delete_clicked(self):
        row = self.rules_list_widget.currentRow()
        if row < 0 or row >= len(self.signage_rules):
            return
        del self.signage_rules[row]
        self.rules_list_widget.takeItem(row)
        if self.rules_list_widget.count():
            self.rules_list_widget.setCurrentRow(min(row, self.rules_list_widget.count() - 1))
        else:
            self._rules_prev_selected_row = -1
            self._set_rule_editor_enabled(False)
        self.save_all_settings()

    def _rules_field_row_add(self):
        self._rules_add_fields_table_row("", "")
        self._rules_mark_dirty_save()

    def _rules_field_row_remove(self):
        r = self.rules_fields_table.currentRow()
        if r < 0 and self.rules_fields_table.rowCount() > 0:
            r = self.rules_fields_table.rowCount() - 1
        if r >= 0:
            self.rules_fields_table.removeRow(r)
        self._rules_mark_dirty_save()

    def _rules_add_fields_table_row(self, column: str, textmarke: str):
        r = self.rules_fields_table.rowCount()
        self.rules_fields_table.insertRow(r)
        le = QLineEdit()
        le.setText(column)
        le.setPlaceholderText("Excel-Überschrift")
        le.editingFinished.connect(self._rules_mark_dirty_save)
        self.rules_fields_table.setCellWidget(r, 0, le)
        cb = QComboBox()
        cb.setEditable(True)
        cb.lineEdit().setPlaceholderText("Textmarke aus Vorlage oder leer")
        cb.setCurrentText(textmarke)
        cb.currentTextChanged.connect(self._rules_mark_dirty_save)
        self.rules_fields_table.setCellWidget(r, 1, cb)
        self._rules_fill_textmarke_combo_items(cb)

    def _rules_fill_textmarke_combo_items(self, cb: QComboBox):
        vo = self.paths.get("vorlagen_ordner", "").strip()
        rel = self.rules_reference_template_combo.currentData() or ""
        cur = cb.currentText()
        cb.blockSignals(True)
        cb.clear()
        marken = []
        if vo and rel:
            abs_p = os.path.normpath(os.path.join(vo, rel))
            marken = liste_textmarken_aus_docx(abs_p)
        for m in marken:
            cb.addItem(m)
        cb.setEditText(cur)
        cb.blockSignals(False)

    def _rules_refresh_all_field_textmarken(self):
        for r in range(self.rules_fields_table.rowCount()):
            w = self.rules_fields_table.cellWidget(r, 1)
            if isinstance(w, QComboBox):
                self._rules_fill_textmarke_combo_items(w)

    def _rules_on_reference_template_changed(self):
        self._rules_refresh_all_field_textmarken()
        self._rules_mark_dirty_save()

    def _rules_on_condition_mode_changed(self, *_):
        if not hasattr(self, "rules_condition_combo"):
            return
        mode = self.rules_condition_combo.currentData()
        show_len = mode in ("length_gt", "length_gte")
        show_needle = mode in ("contains", "equals", "equals_ignorecase")
        show_rx = mode == "regex"
        self.rules_min_len_spin.setVisible(show_len)
        self.rules_needle_edit.setVisible(show_needle)
        self.rules_regex_edit.setVisible(show_rx)
        self._rules_mark_dirty_save()

    def _clear_rule_fields_table(self):
        self.rules_fields_table.setRowCount(0)

    def _load_rule_editor_from_index(self, row: int):
        self._rules_loading_ui = True
        try:
            rule = (
                self.signage_rules[row]
                if row < len(self.signage_rules)
                else self._default_signage_rule_dict()
            )
            self.rules_rule_enabled_check.setChecked(bool(rule.get("enabled", True)))
            self.rules_rule_name_edit.setText(str(rule.get("name", "")))

            when = rule.get("when") or {}
            mode = "not_empty"
            min_v = 3
            needle = ""
            eq_v = ""
            rx_v = ""
            fields = [{"column": "", "textmarke": ""}]

            if "text_condition" in when:
                tc = when.get("text_condition") or {}
                mode = str(tc.get("mode") or "not_empty").strip().lower()
                min_v = int(tc.get("min", tc.get("min_len", 3)))
                needle = str(tc.get("needle", tc.get("contains", "")))
                eq_v = str(tc.get("equals_value", tc.get("equals", "")))
                rx_v = str(tc.get("regex_pattern", tc.get("regex", "")))
                fields = tc.get("fields")
                if not fields:
                    cols = tc.get("columns") or []
                    if isinstance(cols, str):
                        cols = [c.strip() for c in cols.split(",") if c.strip()]
                    fields = [{"column": c, "textmarke": ""} for c in cols]
                if not fields:
                    fields = [{"column": "", "textmarke": ""}]
            elif "text_fields_any_length_gt" in when:
                old = when.get("text_fields_any_length_gt") or {}
                mode = "length_gt"
                min_v = int(old.get("min", 3))
                cols = old.get("columns") or []
                if isinstance(cols, str):
                    cols = [c.strip() for c in cols.split(",") if c.strip()]
                fields = [{"column": c, "textmarke": ""} for c in cols] or [
                    {"column": "", "textmarke": ""}
                ]

            idx = self.rules_condition_combo.findData(mode)
            self.rules_condition_combo.setCurrentIndex(idx if idx >= 0 else 0)
            self.rules_min_len_spin.setValue(min_v)
            if mode == "contains":
                self.rules_needle_edit.setText(needle)
            elif mode in ("equals", "equals_ignorecase"):
                self.rules_needle_edit.setText(eq_v)
            else:
                self.rules_needle_edit.setText("")
            self.rules_regex_edit.setText(rx_v)

            img_col = (rule.get("branch_on_image") or {}).get("column", "")
            self.rules_image_col_edit.setText(str(img_col or ""))

            self.refresh_rules_template_combos()
            ref = str(rule.get("reference_template", "") or "")
            self._pick_template_combo(self.rules_reference_template_combo, ref)
            tif_list = rule.get("templates_if") or []
            tel_list = rule.get("templates_else") or []
            self._pick_template_combo(
                self.rules_template_if_combo, tif_list[0] if tif_list else ""
            )
            self._pick_template_combo(
                self.rules_template_else_combo, tel_list[0] if tel_list else ""
            )

            self._clear_rule_fields_table()
            for f in fields:
                if isinstance(f, str):
                    self._rules_add_fields_table_row(f.strip(), "")
                else:
                    fd = f or {}
                    self._rules_add_fields_table_row(
                        str(fd.get("column", "")),
                        str(fd.get("textmarke", "")),
                    )
            if self.rules_fields_table.rowCount() == 0:
                self._rules_add_fields_table_row("", "")

            self._rules_on_condition_mode_changed()
        finally:
            self._rules_loading_ui = False
        self._rules_refresh_all_field_textmarken()

    def _collect_rule_dict_from_editor(self) -> dict:
        mode = self.rules_condition_combo.currentData() or "not_empty"
        needle = (self.rules_needle_edit.text() or "").strip()
        eq_val = needle if mode in ("equals", "equals_ignorecase") else ""
        needle_val = needle if mode == "contains" else ""
        fields = []
        for r in range(self.rules_fields_table.rowCount()):
            le_w = self.rules_fields_table.cellWidget(r, 0)
            cb_w = self.rules_fields_table.cellWidget(r, 1)
            col = le_w.text().strip() if isinstance(le_w, QLineEdit) else ""
            tm = ""
            if isinstance(cb_w, QComboBox):
                tm = cb_w.currentText().strip()
            if not col and not tm:
                continue
            fields.append({"column": col, "textmarke": tm})
        if not fields:
            fields = [{"column": "", "textmarke": ""}]

        tif = self.rules_template_if_combo.currentData() or self.rules_template_if_combo.currentText().strip()
        tel = self.rules_template_else_combo.currentData() or self.rules_template_else_combo.currentText().strip()
        ref = self.rules_reference_template_combo.currentData() or self.rules_reference_template_combo.currentText().strip()

        row = self.rules_list_widget.currentRow()
        rid = ""
        if 0 <= row < len(self.signage_rules):
            rid = str((self.signage_rules[row] or {}).get("id") or "")
        if not rid:
            rid = f"rule_{row + 1}" if row >= 0 else "rule_1"

        return {
            "id": rid,
            "name": (self.rules_rule_name_edit.text() or f"Regel {row + 1}").strip()
            or f"Regel {row + 1}",
            "enabled": self.rules_rule_enabled_check.isChecked(),
            "reference_template": ref,
            "templates_if": [tif] if tif else [],
            "templates_else": [tel] if tel else [],
            "when": {
                "text_condition": {
                    "mode": mode,
                    "min": int(self.rules_min_len_spin.value()),
                    "needle": needle_val,
                    "equals_value": eq_val,
                    "regex_pattern": (self.rules_regex_edit.text() or "").strip(),
                    "fields": fields,
                }
            },
            "branch_on_image": {"column": (self.rules_image_col_edit.text() or "").strip()},
        }

    def refresh_rules_template_combos(self):
        """Füllt die Vorlagen-Combos (Bezug, IF, ELSE) aus dem aktuellen Vorlagen-Ordner."""
        if not hasattr(self, "rules_template_if_combo"):
            return
        vo = self.paths.get("vorlagen_ordner", "").strip()
        preserve_ref = self.rules_reference_template_combo.currentData()
        preserve_if = self.rules_template_if_combo.currentData()
        preserve_else = self.rules_template_else_combo.currentData()
        for cb in (
            self.rules_reference_template_combo,
            self.rules_template_if_combo,
            self.rules_template_else_combo,
        ):
            cb.blockSignals(True)
            cb.clear()
            cb.addItem("(keine Auswahl)", "")
        if vo and os.path.isdir(vo):
            anlagen, allgemein = sammle_vorlagen_pfade(vo)
            for abs_p in sorted(anlagen + allgemein):
                rel = os.path.normpath(os.path.relpath(abs_p, vo))
                self.rules_reference_template_combo.addItem(rel, rel)
                self.rules_template_if_combo.addItem(rel, rel)
                self.rules_template_else_combo.addItem(rel, rel)
        for cb, preserve in (
            (self.rules_reference_template_combo, preserve_ref),
            (self.rules_template_if_combo, preserve_if),
            (self.rules_template_else_combo, preserve_else),
        ):
            if preserve:
                idx = cb.findData(preserve)
                if idx >= 0:
                    cb.setCurrentIndex(idx)
            cb.blockSignals(False)

    def _snapshot_signage_rules_from_ui(self):
        """Übernimmt die Regel-UI nach self.signage_rules (alle Einträge)."""
        if not hasattr(self, "rules_enabled_check"):
            return
        self.rules_enabled = self.rules_enabled_check.isChecked()
        if hasattr(self, "reuse_lageplan_from_last_export_check"):
            self.reuse_lageplan_from_last_export = (
                self.reuse_lageplan_from_last_export_check.isChecked()
            )
        if not hasattr(self, "rules_list_widget"):
            return
        row = self.rules_list_widget.currentRow()
        if row >= 0:
            self._ensure_signage_rules_length(row + 1)
            self.signage_rules[row] = self._collect_rule_dict_from_editor()
            it = self.rules_list_widget.item(row)
            if it:
                it.setText(self.signage_rules[row].get("name") or f"Regel {row + 1}")

    def _apply_signage_rules_to_ui(self):
        """Lädt self.signage_rules in Liste und Editor."""
        if not hasattr(self, "rules_enabled_check"):
            return
        self.rules_enabled_check.setChecked(self.rules_enabled)
        if hasattr(self, "reuse_lageplan_from_last_export_check"):
            self.reuse_lageplan_from_last_export_check.setChecked(
                self.reuse_lageplan_from_last_export
            )
        self._rules_loading_ui = True
        try:
            self.rules_list_widget.clear()
            if not isinstance(self.signage_rules, list):
                self.signage_rules = []
            for i, rule in enumerate(self.signage_rules):
                name = (rule or {}).get("name") or f"Regel {i + 1}"
                self.rules_list_widget.addItem(str(name))
            if self.signage_rules:
                self.rules_list_widget.setCurrentRow(0)
                self._rules_prev_selected_row = 0
                self._set_rule_editor_enabled(True)
                self._load_rule_editor_from_index(0)
            else:
                self._rules_prev_selected_row = -1
                self._set_rule_editor_enabled(False)
                self._clear_rule_fields_table()
        finally:
            self._rules_loading_ui = False
        if not self.signage_rules:
            self.refresh_rules_template_combos()
