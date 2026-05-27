import sys
import unittest

from PySide6.QtWidgets import (
    QApplication,
    QDialog,
    QGraphicsScene,
    QGraphicsView,
    QPushButton,
)

from ui.preview_subwindow import PreviewSubwindow
from view_models.preview_view_model import PreviewViewModel

app = QApplication.instance()
if app is None:
    app = QApplication(sys.argv)


class TestPreviewSubwindow(unittest.TestCase):
    def setUp(self):
        self.view_model = PreviewViewModel()
        self.dialog = PreviewSubwindow(self.view_model)

    def test_initialization(self):
        self.assertIsInstance(self.dialog, QDialog)

        # Should have a QGraphicsScene and QGraphicsView
        graphics_view = self.dialog.findChild(QGraphicsView)
        self.assertIsNotNone(graphics_view)
        assert graphics_view is not None  # for type checker

        scene = graphics_view.scene()
        self.assertIsInstance(scene, QGraphicsScene)

    def test_buttons_connected(self):
        btn_confirmar = self.dialog.findChild(QPushButton, "btnConfirmar")
        btn_rehacer = self.dialog.findChild(QPushButton, "btnRehacer")

        self.assertIsNotNone(btn_confirmar)
        self.assertIsNotNone(btn_rehacer)
        assert btn_confirmar is not None
        assert btn_rehacer is not None

        self.assertEqual(btn_confirmar.text(), "Confirmar")
        self.assertEqual(btn_rehacer.text(), "Rehacer")


if __name__ == "__main__":
    unittest.main()
