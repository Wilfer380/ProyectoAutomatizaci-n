import unittest
from unittest.mock import MagicMock, patch

from PySide6.QtCore import QObject
from PySide6.QtGui import QImage
from PySide6.QtWidgets import QApplication

from models.asset_record import AssetRecord
from view_models.main_view_model import MainViewModel

app = QApplication.instance() or QApplication([])


class TestMainViewModel(unittest.TestCase):
    def test_initial_state(self):
        vm = MainViewModel()
        self.assertIsInstance(vm, QObject)
        self.assertEqual(vm.selected_file_path, "")
        self.assertFalse(vm.is_processing)
        self.assertEqual(vm.progress_value, 0)

    def _record(self, row: int, section: str = "S1") -> AssetRecord:
        return AssetRecord(
            row_index=row,
            asset_id=f"ID{row}",
            asset_name=f"N{row}",
            section=section,
            image=QImage(),
        )

    def test_process_file_emits_signals(self):
        vm = MainViewModel()

        file_selected_emitted = False

        def on_file_selected(path):
            nonlocal file_selected_emitted
            if path == "test.xlsx":
                file_selected_emitted = True

        vm.fileSelected.connect(on_file_selected)

        processing_started_emitted = False

        def on_processing_started():
            nonlocal processing_started_emitted
            processing_started_emitted = True

        vm.processingStarted.connect(on_processing_started)

        processing_finished_emitted = False
        result_items = []

        def on_processing_finished(items):
            nonlocal processing_finished_emitted, result_items
            processing_finished_emitted = True
            result_items = items

        vm.processingFinished.connect(on_processing_finished)

        error_emitted = False

        def on_error(msg):
            nonlocal error_emitted
            error_emitted = True

        vm.errorOccurred.connect(on_error)

        # Mock ExcelService
        mock_records = [self._record(1)]

        with patch("view_models.main_view_model.ExcelService") as mock_service_cls:
            mock_service = MagicMock()
            mock_service.extract_data.return_value = mock_records
            mock_service_cls.return_value = mock_service

            vm.select_file("test.xlsx")
            self.assertTrue(file_selected_emitted)
            vm.set_selected_filters(["S1"])
            vm.set_selected_records(mock_records)

            vm.process_file()

            self.assertTrue(processing_started_emitted)
            self.assertTrue(processing_finished_emitted)
            self.assertFalse(error_emitted)

            self.assertEqual(len(result_items), 1)
            self.assertEqual(result_items[0].asset_id, "ID1")
            self.assertEqual(result_items[0].asset_name, "N1")
            self.assertEqual(result_items[0].section, "S1")

            self.assertFalse(vm.is_processing)
            self.assertEqual(vm.progress_value, 100)

    def test_empty_configured_selection_does_not_fall_back_to_all_records(self):
        vm = MainViewModel()
        errors = []
        finished = []
        vm.errorOccurred.connect(errors.append)
        vm.processingFinished.connect(finished.append)
        vm._records = [self._record(1), self._record(2)]

        vm.set_selected_filters([])
        vm.process_file()

        self.assertEqual(finished, [])
        self.assertEqual(
            errors, ["Seleccioná al menos un filtro o una etiqueta antes de generar."]
        )

    def test_selecting_new_file_clears_previous_selection(self):
        vm = MainViewModel()
        old_record = self._record(1, "OLD")
        new_record = self._record(2, "NEW")
        vm.set_selected_filters(["OLD"])
        vm.set_selected_records([old_record])

        with patch("view_models.main_view_model.ExcelService") as mock_service_cls:
            mock_service = MagicMock()
            mock_service.extract_data.return_value = [new_record]
            mock_service_cls.return_value = mock_service

            vm.select_file("new.xlsx")

        self.assertEqual(vm.selected_filters, [])
        self.assertEqual(vm.selected_records_by_filter(), {})
        self.assertEqual(vm.records_by_filter(), {"NEW": [new_record]})


if __name__ == "__main__":
    unittest.main()
