from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QVBoxLayout,
)

from models.asset_record import AssetRecord
from .label_selection_dialog import LabelSelectionDialog


class FilterSelectionDialog(QDialog):
    def __init__(
        self,
        filters: list[str],
        records_by_filter: dict[str, list[AssetRecord]],
        checked_filters: list[str] | None = None,
        selected_records_by_filter: dict[str, list[AssetRecord]] | None = None,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("Seleccionar filtros")
        self.resize(760, 560)
        self._records_by_filter = records_by_filter
        self._selected_records_by_filter = selected_records_by_filter or {}
        checked = set(checked_filters or [])

        layout = QVBoxLayout(self)
        title = QLabel(
            "Marcá uno o más filtros para imprimir. "
            "Doble click sobre un filtro abre sus etiquetas para elegirlas una por una."
        )
        title.setWordWrap(True)
        layout.addWidget(title)

        self.list_widget = QListWidget(self)
        self.list_widget.itemDoubleClicked.connect(self._open_filter_labels)
        for filter_name in filters:
            count = len(records_by_filter.get(filter_name, []))
            item = QListWidgetItem(f"{filter_name}  ({count} etiquetas)")
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(
                Qt.CheckState.Checked
                if filter_name in checked
                else Qt.CheckState.Unchecked
            )
            item.setData(Qt.ItemDataRole.UserRole, filter_name)
            self.list_widget.addItem(item)
        layout.addWidget(self.list_widget)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel,
            self,
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def checked_filters(self) -> list[str]:
        selected = []
        for row in range(self.list_widget.count()):
            item = self.list_widget.item(row)
            if item.checkState() == Qt.CheckState.Checked:
                selected.append(str(item.data(Qt.ItemDataRole.UserRole)))
        return selected

    def selected_records(self) -> list[AssetRecord]:
        records = []
        for filter_name in self.checked_filters():
            explicit = self._selected_records_by_filter.get(filter_name)
            if explicit is not None:
                records.extend(explicit)
            else:
                records.extend(self._records_by_filter.get(filter_name, []))
        return records

    def selected_records_by_filter(self) -> dict[str, list[AssetRecord]]:
        return self._selected_records_by_filter

    def _open_filter_labels(self, item: QListWidgetItem) -> None:
        filter_name = str(item.data(Qt.ItemDataRole.UserRole))
        records = self._records_by_filter.get(filter_name, [])
        dialog = LabelSelectionDialog(
            filter_name,
            records,
            checked_records=self._selected_records_by_filter.get(filter_name),
            parent=self,
        )
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self._selected_records_by_filter[filter_name] = dialog.selected_records()
            item.setCheckState(Qt.CheckState.Checked)
