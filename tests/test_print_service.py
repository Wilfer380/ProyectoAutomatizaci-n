import sys
import unittest
from collections.abc import Callable
from typing import cast

from PySide6.QtCore import QRect
from PySide6.QtGui import QFont, QPageLayout, QPainter
from PySide6.QtPrintSupport import QPrinter
from PySide6.QtWidgets import QApplication

from services.driver_check import PrinterDriverMissingError
from services.print_service import LabelPrintConfig, LabelRenderer, PrintService
from view_models.label_item_view_model import LabelItemViewModel

app = QApplication.instance() or QApplication(sys.argv)


class FakePageLayout:
    def __init__(self, rect):
        self._rect = rect

    def paintRectPixels(self, _resolution):
        return self._rect


class FakePrinter:
    def __init__(self, mode=None):
        self.mode = mode
        self.printer_name = ""
        self.page_size = None
        self.page_layout = None
        self.orientation = None
        self.margins = None
        self.margin_unit = None
        self.new_pages = 0
        self.resolution_value = None
        self.full_page = None
        self.paint_rect = QRect(0, 0, 384, 184)

    def setResolution(self, resolution):
        self.resolution_value = resolution

    def resolution(self):
        return self.resolution_value

    def setPrinterName(self, printer_name):
        self.printer_name = printer_name

    def setPageSize(self, page_size):
        self.page_size = page_size

    def setPageOrientation(self, orientation):
        self.orientation = orientation

    def setPageLayout(self, page_layout):
        self.page_layout = page_layout
        self.page_size = page_layout.pageSize()
        self.orientation = page_layout.orientation()
        self.margins = page_layout.margins()

    def setPageMargins(self, margins, unit):
        self.margins = margins
        self.margin_unit = unit

    def setFullPage(self, enabled):
        self.full_page = enabled

    def pageLayout(self):
        return FakePageLayout(self.paint_rect)

    def newPage(self):
        self.new_pages += 1


class FakePainter:
    def __init__(self):
        self.began = False
        self.begin_count = 0
        self.end_count = 0
        self.ended = False
        self.drawn_text = []
        self.drawn_rects = []
        self.drawn_images = []
        self.font_value = None
        self.translate_calls = []
        self.rotate_calls = []

    def begin(self, printer):
        self.begin_count += 1
        self.began = True
        self.printer = printer
        return True

    def end(self):
        self.end_count += 1
        self.ended = True

    def font(self):
        return QFont()

    def setFont(self, font):
        self.font_value = font

    def save(self):
        pass

    def restore(self):
        pass

    def device(self):
        return self.printer

    def translate(self, x, y):
        self.translate_calls.append((x, y))

    def rotate(self, angle):
        self.rotate_calls.append(angle)

    def drawRect(self, rect):
        self.drawn_rects.append(rect)

    def drawText(self, rect, flags, text):
        self.drawn_text.append((rect, flags, text))

    def drawImage(self, rect, image):
        self.drawn_images.append((rect, image))


class FakeRenderer:
    def __init__(self):
        self.items = []

    def render_label(self, painter, item, width_px, height_px):
        self.items.append((item, width_px, height_px))


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
        self.assertIsNotNone(fake_printer.page_layout)
        self.assertIsNotNone(fake_printer.margins)
        self.assertEqual(fake_printer.resolution_value, 203)
        self.assertEqual(fake_printer.orientation, QPageLayout.Orientation.Landscape)
        self.assertTrue(fake_printer.full_page)
        self.assertEqual(fake_printer.new_pages, 0)
        self.assertTrue(fake_painter.began)
        self.assertTrue(fake_painter.ended)
        self.assertEqual(fake_painter.begin_count, 2)
        self.assertEqual(fake_painter.end_count, 2)
        self.assertEqual(fake_renderer.items[0][1:], (384, 184))
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
        from unittest.mock import patch

        with patch("services.driver_check.powershell_printer_names", return_value=()):
            service = PrintService(
                LabelPrintConfig(printer_name="SATO WS408"),
                printer_factory=cast(Callable[..., QPrinter], lambda *_args: FakePrinter()),
                painter_factory=cast(Callable[[], QPainter], FakePainter),
                printer_names_provider=lambda: ["Microsoft Print to PDF"],
            )

            with self.assertRaises(PrinterDriverMissingError):
                service.print_labels([self._make_item()])

    def test_renderer_draws_logo_when_asset_image_missing(self):
        item = self._make_item()
        painter = FakePainter()

        LabelRenderer().render_label(cast(QPainter, painter), item, 384, 184)

        self.assertEqual(len(painter.drawn_images), 1)

    def test_renderer_draws_asset_text(self):
        item = self._make_item()
        painter = FakePainter()

        LabelRenderer().render_label(cast(QPainter, painter), item, 384, 184)

        self.assertEqual(len(painter.drawn_rects), 0)
        self.assertEqual(len(painter.drawn_text), 4)
        printed_text = "\n".join(call[2] for call in painter.drawn_text)
        self.assertIn("A-001", printed_text)
        self.assertIn("Equipo de prueba", printed_text)
        self.assertIn("IT", printed_text)

    def test_rotates_portrait_paint_rect_using_printable_width(self):
        fake_printer = FakePrinter()
        fake_printer.paint_rect = QRect(0, 0, 184, 384)
        fake_painter = FakePainter()
        fake_renderer = FakeRenderer()

        service = PrintService(
            LabelPrintConfig(printer_name="SATO WS408"),
            printer_factory=cast(Callable[..., QPrinter], lambda *_args: fake_printer),
            painter_factory=cast(Callable[[], QPainter], lambda: fake_painter),
            renderer=cast(LabelRenderer, fake_renderer),
            printer_names_provider=lambda: ["SATO WS408"],
        )

        service.print_labels([self._make_item()])

        self.assertEqual(fake_painter.translate_calls, [(184, 0)])
        self.assertEqual(fake_painter.rotate_calls, [90])


if __name__ == "__main__":
    unittest.main()
