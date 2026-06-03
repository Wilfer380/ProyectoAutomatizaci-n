from __future__ import annotations

from pathlib import Path

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


CHECK_ICON_PATH = Path(__file__).resolve().parents[1] / "assets" / "checkbox_checked_blue.svg"


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
        self.list_widget.setStyleSheet(
            """
            QListWidget {
                background: #0B1120;
                color: #E2E8F0;
                border: 1px solid #334155;
                border-radius: 10px;
                padding: 6px;
            }
            QListWidget::item {
                padding: 6px 10px 6px 8px;
                border-radius: 6px;
                min-height: 26px;
            }
            QListWidget::item:hover {
                background: #111C33;
            }
            QListWidget::item:selected {
                background: #1D4ED8;
                color: white;
            }
            QListWidget::indicator {
                width: 16px;
                height: 16px;
            }
            QListWidget::indicator:unchecked {
                border: 1px solid #60A5FA;
                background: transparent;
                border-radius: 3px;
            }
            QListWidget::indicator:checked {
                image: url("{CHECK_ICON_PATH.as_posix()}");
                width: 16px;
                height: 16px;
            }
            QListWidget::indicator:checked:hover {
                image: url("{CHECK_ICON_PATH.as_posix()}");
            }
            """
            .replace("{CHECK_ICON_PATH.as_posix()}", CHECK_ICON_PATH.as_posix())
        )
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
