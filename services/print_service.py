from __future__ import annotations

from collections.abc import Callable, Sequence
from dataclasses import dataclass
from pathlib import Path

from PySide6.QtCore import QMarginsF, QRectF, QSizeF, Qt
from PySide6.QtGui import QFont, QFontMetrics, QImage, QPageLayout, QPageSize, QPainter
from PySide6.QtPrintSupport import QPrinter

from services.driver_check import ensure_printer_driver
from utils.constants import LABEL_HEIGHT_MM, LABEL_WIDTH_MM, TARGET_PRINTER_NAME
from view_models.label_item_view_model import LabelItemViewModel

DEFAULT_WEG_LOGO_PATH = Path(
    r"Q:\PUBLIC\CO_MDE_CONTROL_DOCUMENTAL_CD\IMAGEN CORPORATIVA\2 FORMATOS\1. Logo WEG\2.Logo pequeño para plantillas Negro.png"
)


@dataclass(frozen=True)
class LabelPrintConfig:
    printer_name: str = TARGET_PRINTER_NAME
    label_width_mm: float = LABEL_WIDTH_MM
    label_height_mm: float = LABEL_HEIGHT_MM
    margin_mm: float = 0.0
    resolution_dpi: int = 203
    logo_path: str = str(DEFAULT_WEG_LOGO_PATH)
    separate_jobs: bool = True


class LabelRenderer:
    """Renders one 48x23mm WEG asset label into the current QPainter coordinate system."""

    def __init__(self, logo_path: str | Path | None = DEFAULT_WEG_LOGO_PATH) -> None:
        self.logo = QImage(str(logo_path)) if logo_path else QImage()

    def render_label(
        self,
        painter: QPainter,
        item: LabelItemViewModel,
        width_px: int,
        height_px: int,
    ) -> None:
        content_shift_x = -12.0
        border = 4
        label_rect = QRectF(
            border + content_shift_x,
            border,
            width_px - (border * 2),
            height_px - (border * 2),
        )
        painter.drawRect(label_rect)

        logo_rect = QRectF(8.0 + content_shift_x, 8.0, 68.0, 42.0)
        logo = (
            item.image_data
            if item.image_data is not None and not item.image_data.isNull()
            else self.logo
        )
        if not logo.isNull():
            painter.drawImage(logo_rect, logo)

        header_rect = QRectF(70.0 + content_shift_x, 8.0, width_px - 78.0, 54.0)

        section_font = QFont(painter.font())
        section_font.setBold(False)
        section_font.setPixelSize(22)
        painter.setFont(section_font)
        painter.drawText(
            QRectF(header_rect.left(), 8.0, header_rect.width(), 22.0),
            Qt.AlignmentFlag.AlignCenter,
            item.section,
        )

        id_font = QFont(painter.font())
        id_font.setPixelSize(27)
        id_font.setBold(True)
        painter.setFont(id_font)
        painter.drawText(
            QRectF(header_rect.left(), 31.0, header_rect.width(), 28.0),
            Qt.AlignmentFlag.AlignCenter,
            item.asset_id,
        )

        description, tag_code = self._split_description_and_code(item.asset_name)

        desc_rect = QRectF(14.0 + content_shift_x, 74.0, width_px - 28.0, 54.0)
        desc_flags = Qt.AlignmentFlag.AlignCenter | Qt.TextFlag.TextWordWrap
        desc_font = self._fit_font(
            painter.font(),
            description,
            desc_rect,
            base_size=17,
            min_size=10,
            bold=True,
            flags=desc_flags,
        )
        painter.setFont(desc_font)
        painter.drawText(desc_rect, desc_flags, description)

        code_font = QFont(painter.font())
        code_font.setBold(False)
        code_font.setPixelSize(18)
        painter.setFont(code_font)
        painter.drawText(
            QRectF(14.0 + content_shift_x, 129.0, width_px - 28.0, 24.0),
            Qt.AlignmentFlag.AlignCenter,
            tag_code,
        )

    def _fit_font(
        self,
        base_font: QFont,
        text: str,
        rect: QRectF,
        *,
        base_size: int,
        min_size: int,
        bold: bool,
        flags: Qt.AlignmentFlag | Qt.TextFlag,
    ) -> QFont:
        for size in range(base_size, min_size - 1, -1):
            font = QFont(base_font)
            font.setBold(bold)
            font.setPixelSize(size)
            bounds = QFontMetrics(font).boundingRect(rect.toRect(), int(flags), text)
            if bounds.width() <= rect.width() and bounds.height() <= rect.height():
                return font
        font = QFont(base_font)
        font.setBold(bold)
        font.setPixelSize(min_size)
        return font

    def _split_description_and_code(self, asset_name: str) -> tuple[str, str]:
        parts = str(asset_name).strip().rsplit(" ", 1)
        if len(parts) == 2 and any(char.isdigit() for char in parts[1]):
            return parts[0], parts[1]
        return str(asset_name).strip(), ""


class PrintService:
    def __init__(
        self,
        config: LabelPrintConfig | None = None,
        *,
        printer_factory: Callable[..., QPrinter] | None = None,
        painter_factory: Callable[[], QPainter] | None = None,
        renderer: LabelRenderer | None = None,
        printer_names_provider: Callable[[], Sequence[str]] | None = None,
    ) -> None:
        self.config = config or LabelPrintConfig()
        self._printer_factory: Callable[..., QPrinter] = printer_factory or QPrinter
        self._painter_factory: Callable[[], QPainter] = painter_factory or QPainter
        self._renderer = renderer or LabelRenderer()
        self._printer_names_provider = printer_names_provider

    def print_labels(self, items: Sequence[LabelItemViewModel]) -> None:
        if not items:
            raise ValueError("No hay etiquetas para imprimir.")

        ensure_printer_driver(
            self.config.printer_name,
            printer_names_provider=self._printer_names_provider,
        )
        width_px, height_px = self._label_pixel_size()
        if self.config.separate_jobs:
            for item in items:
                self._print_single_label(item, width_px, height_px)
            return

        printer = self._create_configured_printer()
        painter = self._painter_factory()

        if not painter.begin(printer):
            raise RuntimeError("No se pudo iniciar el trabajo de impresión.")

        try:
            for index, item in enumerate(items):
                if index > 0:
                    printer.newPage()
                self._render_single_label(painter, item, width_px, height_px)
        finally:
            painter.end()

    def _print_single_label(
        self, item: LabelItemViewModel, width_px: int, height_px: int
    ) -> None:
        printer = self._create_configured_printer()
        painter = self._painter_factory()
        if not painter.begin(printer):
            raise RuntimeError("No se pudo iniciar el trabajo de impresión.")
        try:
            self._render_single_label(painter, item, width_px, height_px)
        finally:
            painter.end()

    def _render_single_label(
        self,
        painter: QPainter,
        item: LabelItemViewModel,
        width_px: int,
        height_px: int,
    ) -> None:
        painter.save()
        try:
            self._renderer.render_label(painter, item, width_px, height_px)
        finally:
            painter.restore()

    def _label_pixel_size(self) -> tuple[int, int]:
        return (
            round(self.config.label_width_mm / 25.4 * self.config.resolution_dpi),
            round(self.config.label_height_mm / 25.4 * self.config.resolution_dpi),
        )

    def _create_configured_printer(self) -> QPrinter:
        printer = self._printer_factory(QPrinter.PrinterMode.HighResolution)
        printer.setPrinterName(self.config.printer_name)
        printer.setResolution(self.config.resolution_dpi)

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
