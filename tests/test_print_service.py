import unittest
from collections.abc import Callable
from typing import cast

from PySide6.QtGui import QPainter
from PySide6.QtPrintSupport import QPrinter

from services.driver_check import PrinterDriverMissingError
from services.print_service import LabelPrintConfig, LabelRenderer, PrintService
from view_models.label_item_view_model import LabelItemViewModel


class FakePrinter:
    def __init__(self, mode=None):
        self.mode = mode
        self.printer_name = ""
        self.page_size = None
        self.margins = None
        self.margin_unit = None
        self.new_pages = 0

    def setPrinterName(self, printer_name):
        self.printer_name = printer_name

    def setPageSize(self, page_size):
        self.page_size = page_size

    def setPageMargins(self, margins, unit):
        self.margins = margins
        self.margin_unit = unit

    def newPage(self):
        self.new_pages += 1


class FakePainter:
    def __init__(self):
        self.began = False
        self.ended = False
        self.drawn_text = []
        self.drawn_rects = []
        self.drawn_images = []

    def begin(self, printer):
        self.began = True
        self.printer = printer
        return True

    def end(self):
        self.ended = True

    def drawRect(self, rect):
        self.drawn_rects.append(rect)

    def drawText(self, rect, flags, text):
        self.drawn_text.append((rect, flags, text))

    def drawImage(self, rect, image):
        self.drawn_images.append((rect, image))


class FakeRenderer:
    def __init__(self):
        self.items = []

    def render_label(self, painter, item):
        self.items.append(item)


class TestPrintService(unittest.TestCase):
    def _make_item(self, asset_id="A-001"):
        item = LabelItemViewModel()
        item.asset_id = asset_id
        item.asset_name = "Equipo de prueba"
        item.section = "IT"
        return item

    def test_configures_sato_printer_with_48x23mm_page(self):
        fake_printer = FakePrinter()
        fake_painter = FakePainter()
        fake_renderer = FakeRenderer()

        service = PrintService(
            LabelPrintConfig(printer_name="SATO WS408"),
            printer_factory=cast(Callable[..., QPrinter], lambda *_args: fake_printer),
            painter_factory=cast(Callable[[], QPainter], lambda: fake_painter),
            renderer=cast(LabelRenderer, fake_renderer),
            printer_names_provider=lambda: ["SATO WS408"],
        )

        service.print_labels([self._make_item(), self._make_item("A-002")])

        self.assertEqual(fake_printer.printer_name, "SATO WS408")
        self.assertIsNotNone(fake_printer.page_size)
        self.assertIsNotNone(fake_printer.margins)
        self.assertEqual(fake_printer.new_pages, 1)
        self.assertTrue(fake_painter.began)
        self.assertTrue(fake_painter.ended)
        self.assertEqual(len(fake_renderer.items), 2)

    def test_empty_label_list_is_rejected(self):
        service = PrintService(
            printer_factory=cast(Callable[..., QPrinter], lambda *_args: FakePrinter()),
            painter_factory=cast(Callable[[], QPainter], FakePainter),
            printer_names_provider=lambda: ["SATO WS408"],
        )

        with self.assertRaises(ValueError):
            service.print_labels([])

    def test_missing_driver_is_rejected_before_printing(self):
        service = PrintService(
            LabelPrintConfig(printer_name="SATO WS408"),
            printer_factory=cast(Callable[..., QPrinter], lambda *_args: FakePrinter()),
            painter_factory=cast(Callable[[], QPainter], FakePainter),
            printer_names_provider=lambda: ["Microsoft Print to PDF"],
        )

        with self.assertRaises(PrinterDriverMissingError):
            service.print_labels([self._make_item()])

    def test_renderer_draws_asset_text(self):
        item = self._make_item()
        painter = FakePainter()

        LabelRenderer().render_label(cast(QPainter, painter), item)

        self.assertEqual(len(painter.drawn_rects), 1)
        self.assertEqual(len(painter.drawn_text), 1)
        self.assertIn("A-001", painter.drawn_text[0][2])
        self.assertIn("Equipo de prueba", painter.drawn_text[0][2])


if __name__ == "__main__":
    unittest.main()
