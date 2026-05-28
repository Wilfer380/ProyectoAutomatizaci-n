import sys
import unittest
from unittest.mock import MagicMock, patch

from PySide6.QtWidgets import QApplication

from models.app_settings import AppSettings
from ui.main_window import MainWindow
from view_models.main_view_model import MainViewModel

app = QApplication.instance()
if app is None:
    app = QApplication(sys.argv)


class TestMainWindow(unittest.TestCase):
    def setUp(self):
        self.settings = AppSettings()
        self.view_model = MainViewModel()
        self.window = MainWindow(self.settings, self.view_model)

    def test_no_word_ui_elements(self):
        # We should NOT have Word related chips or fields
        self.assertFalse(hasattr(self.window, "word_state_chip"))
        self.assertFalse(hasattr(self.window, "word_path_edit"))
        self.assertFalse(hasattr(self.window, "select_word_button"))
        self.assertFalse(hasattr(self.window, "word_hint_label"))

    def test_view_model_binding(self):
        # Check that signals from VM update UI
        self.view_model.set_progress(45)
        self.assertEqual(self.window.progress_bar.value(), 45)

        self.view_model._load_records_and_filters = MagicMock()
        self.view_model.select_file("C:/fake/path.xlsx")
        self.assertEqual(self.window.excel_path_edit.text(), "C:/fake/path.xlsx")

    def test_excel_button_selects_file_in_view_model(self):
        self.view_model._load_records_and_filters = MagicMock()
        self.window.choose_excel_file = MagicMock(return_value="C:/fake/path.xlsx")

        self.window.select_excel_button.click()

        self.assertEqual(self.view_model.selected_file_path, "C:/fake/path.xlsx")
        self.assertEqual(self.window.excel_path_edit.text(), "C:/fake/path.xlsx")

    def test_manual_excel_path_updates_view_model(self):
        self.view_model._load_records_and_filters = MagicMock()
        self.window.excel_path_edit.setText("C:/manual/path.xlsx")
        self.window.excel_path_edit.editingFinished.emit()

        self.assertEqual(self.view_model.selected_file_path, "C:/manual/path.xlsx")

    def test_start_button_triggers_processing(self):
        self.view_model.process_file = MagicMock()
        self.window.start_button.click()
        self.view_model.process_file.assert_called_once()

    def test_validate_configured_printer_shows_friendly_missing_driver_error(self):
        self.window.show_error = MagicMock()
        with patch("ui.main_window.check_printer_driver") as check_printer_driver:
            check_printer_driver.return_value.installed = False
            check_printer_driver.return_value.message = (
                "Controlador SATO WS408 no detectado. Contactá a TI."
            )

            result = self.window.validate_configured_printer()

        self.assertFalse(result)
        self.window.show_error.assert_called_once()
        self.assertEqual(
            self.window.printer_state_chip.text(), "Impresora no detectada"
        )


if __name__ == "__main__":
    unittest.main()
