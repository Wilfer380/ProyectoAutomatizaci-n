import unittest
from unittest.mock import MagicMock, patch
from PySide6.QtCore import QObject, QEventLoop, QTimer
from view_models.main_view_model import MainViewModel
from models.asset_record import AssetRecord
from PySide6.QtGui import QImage, QGuiApplication

app = QGuiApplication.instance() or QGuiApplication([])


class TestMainViewModel(unittest.TestCase):
    def test_initial_state(self):
        vm = MainViewModel()
        self.assertIsInstance(vm, QObject)
        self.assertEqual(vm.selected_file_path, "")
        self.assertFalse(vm.is_processing)
        self.assertEqual(vm.progress_value, 0)

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
        mock_records = [
            AssetRecord(
                row_index=2,
                asset_id="ID1",
                asset_name="N1",
                section="S1",
                image=QImage(),
            )
        ]

        with patch("view_models.main_view_model.ExcelService") as mock_service_cls:
            mock_service = MagicMock()
            mock_service.extract_data.return_value = mock_records
            mock_service_cls.return_value = mock_service

            vm.select_file("test.xlsx")
            self.assertTrue(file_selected_emitted)

            # Start processing
            vm.process_file()
            worker_ref = vm.worker

            # Since QThread is likely used, we need an event loop to wait for it.
            loop = QEventLoop()

            # Quit the loop when processing finishes or after a timeout
            vm.processingFinished.connect(lambda items: loop.quit())
            vm.errorOccurred.connect(lambda err: loop.quit())

            # Timeout in case the signals are never emitted
            QTimer.singleShot(2000, loop.quit)

            loop.exec()

            # Ensure the worker has completely finished before exiting the test
            if worker_ref:
                worker_ref.wait()

            self.assertTrue(processing_started_emitted)
            self.assertTrue(processing_finished_emitted)
            self.assertFalse(error_emitted)

            self.assertEqual(len(result_items), 1)
            self.assertEqual(result_items[0].asset_id, "ID1")
            self.assertEqual(result_items[0].asset_name, "N1")
            self.assertEqual(result_items[0].section, "S1")

            self.assertFalse(vm.is_processing)
            self.assertEqual(vm.progress_value, 100)


if __name__ == "__main__":
    unittest.main()
