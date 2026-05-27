from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass

from PySide6.QtCore import QMarginsF, QRectF, QSizeF, Qt
from PySide6.QtGui import QPageLayout, QPageSize, QPainter
from PySide6.QtPrintSupport import QPrinter

from view_models.label_item_view_model import LabelItemViewModel


@dataclass(frozen=True)
class LabelPrintConfig:
    printer_name: str = "SATO WS408"
    label_width_mm: float = 48.0
    label_height_mm: float = 23.0
    margin_mm: float = 0.0


class LabelRenderer:
    """Renders one 48x23mm label into the current QPainter coordinate system."""

    def render_label(self, painter: QPainter, item: LabelItemViewModel) -> None:
        label_rect = QRectF(0.0, 0.0, 48.0, 23.0)
        text_rect = QRectF(1.5, 2.0, 25.0, 19.0)
        image_rect = QRectF(
            28.0 + float(item.image_offset_x),
            2.0 + float(item.image_offset_y),
            18.0 * float(item.image_scale),
            19.0 * float(item.image_scale),
        )

        painter.drawRect(label_rect)
        painter.drawText(
            text_rect,
            Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter,
            self._label_text(item),
        )

        image = item.image_data
        if image is not None and not image.isNull():
            painter.drawImage(image_rect, image)

    def _label_text(self, item: LabelItemViewModel) -> str:
        return "\n".join(
            part
            for part in (item.asset_id, item.asset_name, item.section)
            if str(part).strip()
        )


class PrintService:
    def __init__(
        self,
        config: LabelPrintConfig | None = None,
        *,
        printer_factory: Callable[..., QPrinter] | None = None,
        painter_factory: Callable[[], QPainter] | None = None,
        renderer: LabelRenderer | None = None,
    ) -> None:
        self.config = config or LabelPrintConfig()
        self._printer_factory: Callable[..., QPrinter] = printer_factory or QPrinter
        self._painter_factory: Callable[[], QPainter] = painter_factory or QPainter
        self._renderer = renderer or LabelRenderer()

    def print_labels(self, items: Sequence[LabelItemViewModel]) -> None:
        if not items:
            raise ValueError("No hay etiquetas para imprimir.")

        printer = self._create_configured_printer()
        painter = self._painter_factory()

        if not painter.begin(printer):
            raise RuntimeError("No se pudo iniciar el trabajo de impresión.")

        try:
            for index, item in enumerate(items):
                if index > 0:
                    printer.newPage()
                self._renderer.render_label(painter, item)
        finally:
            painter.end()

    def _create_configured_printer(self) -> QPrinter:
        printer = self._printer_factory(QPrinter.PrinterMode.HighResolution)
        printer.setPrinterName(self.config.printer_name)

        page_size = QPageSize(
            QSizeF(self.config.label_width_mm, self.config.label_height_mm),
            QPageSize.Unit.Millimeter,
        )
        printer.setPageSize(page_size)
        printer.setPageMargins(
            QMarginsF(
                self.config.margin_mm,
                self.config.margin_mm,
                self.config.margin_mm,
                self.config.margin_mm,
            ),
            QPageLayout.Unit.Millimeter,
        )
        return printer
