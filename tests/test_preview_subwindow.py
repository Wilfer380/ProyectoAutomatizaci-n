import sys
import unittest

from PySide6.QtGui import QImage
from PySide6.QtWidgets import QApplication, QDialog, QLabel, QPushButton, QScrollArea

from ui.preview_subwindow import PreviewSubwindow
from view_models.label_item_view_model import LabelItemViewModel
from view_models.preview_view_model import PreviewViewModel

app = QApplication.instance()
if app is None:
    app = QApplication(sys.argv)


class TestPreviewSubwindow(unittest.TestCase):
    def setUp(self):
        self.view_model = PreviewViewModel(preview_image_callback=self._preview_image)
        item = LabelItemViewModel()
        item.asset_id = "A-001"
        item.asset_name = "Equipo de prueba"
        item.section = "IT"
        self.view_model.set_items([item])
        self.dialog = PreviewSubwindow(self.view_model)

    def _preview_image(self, _item):
        image = QImage(384, 184, QImage.Format.Format_ARGB32)
        image.fill(0xFFFFFFFF)
        return image

    def test_initialization(self):
        self.assertIsInstance(self.dialog, QDialog)

        scroll_area = self.dialog.findChild(QScrollArea, "previewScrollArea")
        self.assertIsNotNone(scroll_area)

        label_preview = self.dialog.findChild(QLabel, "labelPreviewImage")
        self.assertIsNotNone(label_preview)
        assert label_preview is not None
        self.assertIsNotNone(label_preview.pixmap())

    def test_buttons_connected(self):
        btn_confirmar = self.dialog.findChild(QPushButton, "btnConfirmar")
        btn_rehacer = self.dialog.findChild(QPushButton, "btnRehacer")

        self.assertIsNotNone(btn_confirmar)
        self.assertIsNotNone(btn_rehacer)
        assert btn_confirmar is not None
        assert btn_rehacer is not None

        self.assertEqual(btn_confirmar.text(), "Confirmar e imprimir")
        self.assertEqual(btn_rehacer.text(), "Rechazar")


if __name__ == "__main__":
    unittest.main()
