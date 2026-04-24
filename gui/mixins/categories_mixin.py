"""Kategorien-Tab: Präfixe & Ordnernamen für die Exportstruktur."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QCheckBox,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from ..constants import BUILD_DATE, VERSION


class CategoriesTabMixin:
    """Dokumentkategorien bearbeiten (Tab „Kategorien“)."""

    def setup_categories_tab(self, tab):
        """Erstellt den Inhalt des 'Kategorien'-Tabs für die Verwaltung der Ausgabeordner."""
        layout = QVBoxLayout(tab)

        # Versions-Info
        version_info = QLabel(f"Version {VERSION} - Build {BUILD_DATE}")
        version_info.setStyleSheet("color: #666; font-size: 10px; padding: 5px;")
        version_info.setAlignment(Qt.AlignCenter)
        layout.addWidget(version_info)

        group_box = QGroupBox("Dokumentkategorien")
        layout.addWidget(group_box)
        group_layout = QVBoxLayout(group_box)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll_content = QWidget()
        self.categories_ui_layout = QVBoxLayout(scroll_content)
        scroll.setWidget(scroll_content)

        for prefix, name in self.categories.items():
            self.add_category_widget(prefix, name)

        btn_layout = QHBoxLayout()
        add_btn = QPushButton("+")
        add_btn.setFixedWidth(30)
        add_btn.setToolTip("Neue Kategorie hinzufügen")
        add_btn.clicked.connect(self.add_category)
        remove_btn = QPushButton("-")
        remove_btn.setFixedWidth(30)
        remove_btn.setToolTip("Ausgewählte Kategorie entfernen")
        remove_btn.clicked.connect(self.remove_category)
        btn_layout.addWidget(add_btn)
        btn_layout.addWidget(remove_btn)
        btn_layout.addStretch()

        group_layout.addLayout(btn_layout)
        group_layout.addWidget(scroll)

        save_btn = QPushButton("Kategorien speichern")
        save_btn.clicked.connect(self.save_categories)
        layout.addWidget(save_btn)
        layout.addStretch()

    def add_category_widget(self, prefix, name):
        """Fügt der UI eine neue Zeile zur Eingabe einer Kategorie hinzu."""
        row_layout = QHBoxLayout()
        checkbox = QCheckBox()
        prefix_edit = QLineEdit(prefix)
        name_edit = QLineEdit(name)

        row_layout.addWidget(checkbox)
        row_layout.addWidget(QLabel("Präfix:"))
        row_layout.addWidget(prefix_edit)
        row_layout.addWidget(QLabel("Ordnername:"))
        row_layout.addWidget(name_edit)

        self.categories_ui_layout.addLayout(row_layout)
        self.category_widgets[row_layout] = (checkbox, prefix_edit, name_edit)

    def add_category(self):
        """Event-Handler, um eine neue, leere Kategorie-Zeile hinzuzufügen."""
        self.add_category_widget("", "")

    def remove_category(self):
        """Event-Handler, um alle ausgewählten Kategorie-Zeilen zu entfernen."""
        to_remove = [l for l, (c, _, _) in self.category_widgets.items() if c.isChecked()]
        for layout in to_remove:
            while layout.count():
                child = layout.takeAt(0)
                if child.widget():
                    child.widget().deleteLater()
            self.categories_ui_layout.removeItem(layout)
            layout.deleteLater()
            del self.category_widgets[layout]

    def save_categories(self):
        """Sammelt die Daten aus den Kategorie-Eingabefeldern und speichert sie."""
        self.categories.clear()
        # Werte bestehen aus (checkbox, prefix_edit, name_edit)
        for checkbox, prefix_edit, name_edit in self.category_widgets.values():
            prefix = prefix_edit.text().strip()
            name = name_edit.text().strip()
            if prefix and name:
                self.categories[prefix] = name
        self.save_all_settings()
        QMessageBox.information(self, "Gespeichert", "Kategorien aktualisiert.")
