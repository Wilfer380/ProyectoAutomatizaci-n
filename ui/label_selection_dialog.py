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


class LabelSelectionDialog(QDialog):
    def __init__(
        self,
        filter_name: str,
        records: list[AssetRecord],
        checked_records: list[AssetRecord] | None = None,
        parent=None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle(f"Etiquetas de {filter_name}")
        self.resize(720, 520)
        self._records = records
        checked_ids = (
            {record.row_index for record in checked_records}
            if checked_records is not None
            else None
        )

        layout = QVBoxLayout(self)
        title = QLabel(
            f"Seleccioná las etiquetas que querés imprimir para: {filter_name}"
        )
        title.setWordWrap(True)
        layout.addWidget(title)

        self.list_widget = QListWidget(self)
        for record in records:
            item = QListWidgetItem(
                f"{record.asset_id}  —  {record.asset_name}  —  {record.section}"
            )
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(
                Qt.CheckState.Checked
                if checked_ids is None or record.row_index in checked_ids
                else Qt.CheckState.Unchecked
            )
            item.setData(Qt.ItemDataRole.UserRole, record)
            self.list_widget.addItem(item)
        layout.addWidget(self.list_widget)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel,
            self,
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def selected_records(self) -> list[AssetRecord]:
        selected = []
        for row in range(self.list_widget.count()):
            item = self.list_widget.item(row)
            if item.checkState() == Qt.CheckState.Checked:
                selected.append(item.data(Qt.ItemDataRole.UserRole))
        return selected
