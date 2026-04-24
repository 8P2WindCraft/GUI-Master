"""
Qt-Modelle für die Anbindung von Daten (z. B. Pandas) an Tabellenansichten.
"""
from PySide6.QtCore import Qt, QAbstractTableModel


class PandasModel(QAbstractTableModel):
    """
    Tabellenmodell für PySide6: Pandas-DataFrame als Quelle für eine QTableView
    (Excel-Daten in der GUI).
    """
    def __init__(self, df):
        super().__init__()
        self._df = df

    def rowCount(self, parent=None):
        return self._df.shape[0]

    def columnCount(self, parent=None):
        return self._df.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid() and role == Qt.DisplayRole:
            return str(self._df.iloc[index.row(), index.column()])

    def headerData(self, col, orient, role=Qt.DisplayRole):
        if orient == Qt.Horizontal and role == Qt.DisplayRole:
            return self._df.columns[col]
