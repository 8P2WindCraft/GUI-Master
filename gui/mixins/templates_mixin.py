"""Vorlagen-Tab: Kategorien, Checkbox-Auswahl, Pfade für die Generierung."""
import os

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from core.logic import sammle_vorlagen_pfade


class TemplatesMixin:
    """Auflistung der .docx unter anlagen/ und allgemein/, Teilmengen-Auswahl."""

    def setup_templates_tab(self, tab):
        """Tab: Welche Word-Vorlagen (anlagen/allgemein) für Export und Trockenlauf genutzt werden."""
        layout = QVBoxLayout(tab)
        hint_top = QLabel(
            "Die Pfade zu Excel und Vorlagen-Ordner legst du unter <b>Hauptsteuerung</b> fest. "
            "Hier wählst du die .docx-Dateien für den nächsten Lauf."
        )
        hint_top.setWordWrap(True)
        hint_top.setStyleSheet("color: #555;")
        layout.addWidget(hint_top)

        vorlagen_pick = QGroupBox("Vorlagen für diesen Lauf")
        vorlagen_pick_layout = QVBoxLayout(vorlagen_pick)
        self._template_hint_label = QLabel(
            "Nur Unterordner …\\anlagen\\ und …\\allgemein\\ werden verwendet."
        )
        self._template_hint_label.setWordWrap(True)
        self._template_hint_label.setStyleSheet("color: #555;")
        vorlagen_pick_layout.addWidget(self._template_hint_label)
        tpl_btn_row = QHBoxLayout()
        refresh_tpl_btn = QPushButton("Vorlagen einlesen")
        refresh_tpl_btn.setToolTip("Liste aus dem Vorlagen-Ordner neu aufbauen")
        refresh_tpl_btn.clicked.connect(self.refresh_template_checkboxes)
        tpl_btn_row.addWidget(refresh_tpl_btn)
        all_tpl_btn = QPushButton("Alle auswählen")
        all_tpl_btn.clicked.connect(self._template_select_all)
        tpl_btn_row.addWidget(all_tpl_btn)
        none_tpl_btn = QPushButton("Keine auswählen")
        none_tpl_btn.clicked.connect(self._template_select_none)
        tpl_btn_row.addWidget(none_tpl_btn)
        tpl_btn_row.addStretch()
        vorlagen_pick_layout.addLayout(tpl_btn_row)
        self._template_preview_status = QLabel("")
        self._template_preview_status.setStyleSheet("color: #444; font-size: 12px;")
        self._template_preview_status.setWordWrap(True)
        self._template_preview_status.setVisible(False)
        self._template_preview_progress = QProgressBar()
        self._template_preview_progress.setVisible(False)
        self._template_preview_progress.setMinimumHeight(16)
        self._template_preview_progress.setTextVisible(True)
        self._template_preview_progress.setFormat("%p%")
        vorlagen_pick_layout.addWidget(self._template_preview_status)
        vorlagen_pick_layout.addWidget(self._template_preview_progress)
        self._template_row_preview_active = False
        tpl_scroll = QScrollArea()
        tpl_scroll.setWidgetResizable(True)
        tpl_scroll.setMinimumHeight(200)
        tpl_scroll_content = QWidget()
        self._template_list_layout = QVBoxLayout(tpl_scroll_content)
        self._template_list_layout.addStretch()
        tpl_scroll.setWidget(tpl_scroll_content)
        vorlagen_pick_layout.addWidget(tpl_scroll)
        layout.addWidget(vorlagen_pick)
        layout.addStretch()

        self.refresh_template_checkboxes()

    def _vorlage_kategorie_label(self, vorlage_basename):
        """Kategorie-Anzeigename wie beim Export (Präfix aus Tab „Kategorien“), sonst Sonstiges."""
        name_lower = vorlage_basename.lower()
        for prefix, cat_name in self.categories.items():
            if name_lower.startswith(prefix.lower()):
                return cat_name
        return "Sonstiges"

    def _gruppiere_vorlagen_nach_kategorie(self, abs_paths):
        """
        Gruppiert absolute Vorlagen-Pfade nach Kategorie, sortiert Gruppen und Dateien.
        Liefert [(kategorie_name, [abs_pfad, ...]), ...].
        """
        buckets = {}
        for p in abs_paths:
            lab = self._vorlage_kategorie_label(os.path.basename(p))
            buckets.setdefault(lab, []).append(p)
        for lab in buckets:
            buckets[lab].sort(key=lambda x: os.path.basename(x).lower())

        all_labels = set(buckets.keys())
        category_order = []
        for _pref, lab in self.categories.items():
            if lab not in category_order:
                category_order.append(lab)
        ordered = [l for l in category_order if l in all_labels]
        rest = sorted(all_labels - set(ordered) - {"Sonstiges"})
        ordered.extend(rest)
        if "Sonstiges" in all_labels:
            ordered.append("Sonstiges")
        return [(l, buckets[l]) for l in ordered]

    def refresh_template_checkboxes(self):
        """Baut die Checkbox-Liste aus dem aktuellen Vorlagen-Ordner auf."""
        while self._template_list_layout.count():
            item = self._template_list_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
        self._template_checkbox_by_rel = {}
        vo = self.paths.get('vorlagen_ordner', '').strip()
        if not vo or not os.path.isdir(vo):
            self._template_hint_label.setText("Bitte einen gültigen Vorlagen-Ordner wählen.")
            return

        saved = self.selected_template_rel_paths
        saved_set = None
        if saved is not None:
            saved_set = {os.path.normpath(s) for s in saved}

        def add_checkbox_for_abs(abs_p):
            rel = os.path.normpath(os.path.relpath(abs_p, vo))
            row = QWidget()
            row_l = QHBoxLayout(row)
            row_l.setContentsMargins(0, 2, 0, 2)
            cb = QCheckBox(rel)
            if saved_set is None:
                cb.setChecked(True)
            else:
                cb.setChecked(rel in saved_set)
            cb.stateChanged.connect(self.save_all_settings)
            prev_btn = QPushButton("Vorschau")
            prev_btn.setToolTip("Nur diese Vorlage mit der ersten Excel-Zeile erzeugen (Beispiel).")
            prev_btn.setFixedWidth(88)
            prev_btn.clicked.connect(lambda checked=False, p=abs_p: self.start_preview_single_template(p))
            row_l.addWidget(cb, 1)
            row_l.addWidget(prev_btn)
            self._template_checkbox_by_rel[rel] = cb
            self._template_list_layout.addWidget(row)

        def add_group_mit_kategorien(title, paths):
            if not paths:
                return
            self._template_list_layout.addWidget(QLabel(f"<b>{title}</b>"))
            for cat_label, path_list in self._gruppiere_vorlagen_nach_kategorie(paths):
                sub = QLabel(f"  {cat_label}")
                sub.setStyleSheet("color: #444; margin-top: 4px;")
                self._template_list_layout.addWidget(sub)
                for abs_p in path_list:
                    add_checkbox_for_abs(abs_p)

        anlagen, allgemein = sammle_vorlagen_pfade(vo)
        add_group_mit_kategorien("Anlagen", anlagen)
        add_group_mit_kategorien("Allgemein", allgemein)

        if not self._template_checkbox_by_rel:
            self._template_hint_label.setText(
                "Keine .docx-Vorlagen unter …\\anlagen\\ oder …\\allgemein\\ gefunden."
            )
        else:
            self._template_hint_label.setText(
                "Nur angehakte Vorlagen werden erzeugt. Sortierung nach Kategorien wie unter „Kategorien“ "
                "(Dateinamen-Präfixe, z. B. b_, ba_); ohne passendes Präfix: „Sonstiges“."
            )

    def _snapshot_template_selection(self):
        if not hasattr(self, '_template_checkbox_by_rel') or not self._template_checkbox_by_rel:
            return
        checked = [r for r, cb in self._template_checkbox_by_rel.items() if cb.isChecked()]
        all_rels = list(self._template_checkbox_by_rel.keys())
        if not all_rels:
            return
        self.selected_template_rel_paths = None if len(checked) == len(all_rels) else checked

    def _template_select_all(self):
        if not hasattr(self, '_template_checkbox_by_rel'):
            return
        for cb in self._template_checkbox_by_rel.values():
            cb.setChecked(True)
        self.save_all_settings()

    def _template_select_none(self):
        if not hasattr(self, '_template_checkbox_by_rel'):
            return
        for cb in self._template_checkbox_by_rel.values():
            cb.setChecked(False)
        self.save_all_settings()

    def _template_paths_for_worker(self):
        """None = alle Vorlagen; Liste = nur diese absoluten Pfade; [] = keine / nichts gewählt."""
        vo = self.paths.get('vorlagen_ordner', '').strip()
        if not hasattr(self, '_template_checkbox_by_rel') or not self._template_checkbox_by_rel:
            return []
        all_rels = list(self._template_checkbox_by_rel.keys())
        checked = [r for r, cb in self._template_checkbox_by_rel.items() if cb.isChecked()]
        if not checked:
            return []
        if len(checked) == len(all_rels):
            return None
        return [os.path.normpath(os.path.join(vo, r)) for r in checked]
