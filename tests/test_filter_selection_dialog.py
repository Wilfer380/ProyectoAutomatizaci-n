import sys
import unittest

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QApplication

from models.asset_record import AssetRecord
from ui.filter_selection_dialog import FilterSelectionDialog
from ui.label_selection_dialog import LabelSelectionDialog

app = QApplication.instance()
if app is None:
    app = QApplication(sys.argv)


class TestFilterSelectionDialog(unittest.TestCase):
    def _record(self, row: int, section: str) -> AssetRecord:
        return AssetRecord(
            row_index=row,
            asset_id=f"A-{row:03d}",
            asset_name=f"Equipo {row}",
            section=section,
        )

    def test_checked_filter_selects_all_records_when_no_detail_override(self):
        records = [self._record(1, "B1_WTC"), self._record(2, "B1_WTC")]
        dialog = FilterSelectionDialog(
            ["B1_WTC"], {"B1_WTC": records}, checked_filters=["B1_WTC"]
        )

        self.assertEqual(dialog.checked_filters(), ["B1_WTC"])
        self.assertEqual(dialog.selected_records(), records)

    def test_detail_override_limits_records_for_filter(self):
        records = [self._record(1, "B1_WTC"), self._record(2, "B1_WTC")]
        dialog = FilterSelectionDialog(
            ["B1_WTC"],
            {"B1_WTC": records},
            checked_filters=["B1_WTC"],
            selected_records_by_filter={"B1_WTC": [records[1]]},
        )

        self.assertEqual(dialog.selected_records(), [records[1]])

    def test_label_selection_dialog_returns_checked_records(self):
        records = [self._record(1, "B1_WTC"), self._record(2, "B1_WTC")]
        dialog = LabelSelectionDialog("B1_WTC", records)
        first_item = dialog.list_widget.item(0)
        first_item.setCheckState(Qt.CheckState.Unchecked)

        self.assertEqual(dialog.selected_records(), [records[1]])

    def test_label_selection_dialog_restores_checked_records(self):
        records = [self._record(1, "B1_WTC"), self._record(2, "B1_WTC")]
        dialog = LabelSelectionDialog("B1_WTC", records, checked_records=[records[1]])

        self.assertEqual(dialog.list_widget.item(0).checkState(), Qt.CheckState.Unchecked)
        self.assertEqual(dialog.list_widget.item(1).checkState(), Qt.CheckState.Checked)


if __name__ == "__main__":
    unittest.main()
