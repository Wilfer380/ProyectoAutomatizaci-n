from __future__ import annotations

import ctypes
import logging
import re
import time
from dataclasses import dataclass
from pathlib import Path

from ctypes import wintypes

from models.asset_record import AssetRecord
from utils.constants import (
    BLOCK_SIZE,
    EXCEL_LABEL_GROUP_NAMES,
    EXCEL_SHEET_LABEL,
    EXCEL_SHEET_SOURCE,
    LABEL_COLUMNS,
    LABEL_IMAGE_HEIGHT_PX,
    LABEL_IMAGE_WIDTH_PX,
    LABEL_OUTPUT_START_ROW,
    SOURCE_HEADERS,
)
from utils.normalization import normalize_excel_scalar
from utils.constants import APP_NAME

try:
    from PIL import Image, ImageDraw, ImageFont, ImageGrab
except ImportError:  # pragma: no cover
    Image = None
    ImageDraw = None
    ImageFont = None
    ImageGrab = None

pythoncom = None
win32 = None


def _load_pywin32():
    global pythoncom, win32
    if pythoncom is not None and win32 is not None:
        return pythoncom, win32

    from utils.win32_bootstrap import load_pywin32

    pythoncom, win32 = load_pywin32()
    return pythoncom, win32

@dataclass(slots=True)
class LabelImageExportResult:
    position: int
    group_name: str
    output_path: Path
    original_size_px: tuple[int, int]
    target_size_px: tuple[int, int]


class ExcelService:
    def __init__(self) -> None:
        self.logger = logging.getLogger(APP_NAME)
        self.excel_app = None
        self.workbook = None
        self._owns_excel_app = False
        self._com_initialized = False

    def open(self, workbook_path: str, visible: bool = False) -> None:
        try:
            pythoncom_module, win32_module = _load_pywin32()
        except ImportError as exc:
            raise RuntimeError("pywin32 no estÃ¡ instalado.") from exc

        pythoncom_module.CoInitialize()
        self._com_initialized = True
        self._owns_excel_app = True
        self.excel_app = None
        self.workbook = None
        try:
            self.excel_app = win32_module.DispatchEx("Excel.Application")
            self.excel_app.Visible = False
            self.excel_app.DisplayAlerts = False
            self.excel_app.ScreenUpdating = False
            try:
                self.excel_app.DisplayStatusBar = False
            except Exception:
                pass
            try:
                self.excel_app.Interactive = False
            except Exception:
                pass
            try:
                self.excel_app.EnableEvents = False
            except Exception:
                pass
            try:
                self.excel_app.UserControl = False
            except Exception:
                pass
            try:
                self.excel_app.AskToUpdateLinks = False
            except Exception:
                pass
            self.logger.info("Instancia de Excel creada. Visible=False")
            self.workbook = self.excel_app.Workbooks.Open(
                Filename=str(Path(workbook_path).resolve()),
                UpdateLinks=0,
                ReadOnly=False,
                IgnoreReadOnlyRecommended=True,
                AddToMru=False,
                Notify=False,
                Local=True,
            )
            self._replace_unsupported_xlookup_formulas()
            self.logger.info("Excel abierto en segundo plano: %s", Path(workbook_path).resolve())
        except Exception:
            self._teardown(save_changes=False)
            raise

    def close(self, save_changes: bool = False) -> None:
        self._teardown(save_changes=save_changes)

    def _teardown(self, save_changes: bool = False) -> None:
        workbook = self.workbook
        excel_app = self.excel_app
        owns_excel_app = self._owns_excel_app
        com_initialized = self._com_initialized

        self.workbook = None
        self.excel_app = None
        self._owns_excel_app = False
        self._com_initialized = False

        try:
            if workbook is not None:
                workbook.Close(SaveChanges=save_changes)
        except Exception:
            pass

        try:
            if excel_app is not None and owns_excel_app:
                self._close_own_excel_login_windows(excel_app)
        except Exception:
            self.logger.exception("No se pudo cerrar la ventana de inicio de sesión de Excel durante la limpieza.")

        try:
            if excel_app is not None and owns_excel_app:
                excel_app.Quit()
                self.logger.info("Excel cerrado al finalizar el proceso.")
        except Exception:
            pass

        try:
            if excel_app is not None and owns_excel_app:
                self._close_own_excel_login_windows(excel_app)
        except Exception:
            self.logger.exception("No se pudo verificar/cerrar Excel tras Quit().")

        # Intentionally do not CoUninitialize here.
        # Word and Excel COM run in the same worker apartment; uninitializing
        # one service would disconnect the other still-alive COM proxy.

    def _close_own_excel_login_windows(self, excel_app) -> None:
        excel_pid = self._get_excel_pid(excel_app)
        if excel_pid <= 0:
            self.logger.info("No se pudo identificar el PID de la instancia propia de Excel al cerrar.")
            return

        closed_titles: list[str] = []
        failed_titles: list[str] = []

        @ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
        def enum_windows_proc(hwnd, lparam):  # noqa: ANN001, ARG001
            try:
                pid = self._get_window_pid(hwnd)
                if pid != excel_pid:
                    return True
                title = self._get_window_title(hwnd)
                if not title or not self._is_excel_login_window(title):
                    return True
                if self._close_window(hwnd):
                    closed_titles.append(title)
                else:
                    failed_titles.append(title)
                return True
            except Exception:
                return True

        ctypes.windll.user32.EnumWindows(enum_windows_proc, 0)
        if closed_titles:
            self.logger.info("Se detectó y cerró la ventana de Excel propia: %s", "; ".join(closed_titles))
        if failed_titles:
            self.logger.warning("Se detectó la ventana de Excel propia pero no pudo cerrarse: %s", "; ".join(failed_titles))
        if not closed_titles and not failed_titles:
            self.logger.info("No se detectó ventana propia de Excel para cerrar durante la limpieza final.")

    def _get_excel_pid(self, excel_app) -> int:
        try:
            hwnd = int(excel_app.Hwnd)
        except Exception:
            return 0
        return self._get_window_pid(hwnd) if hwnd else 0

    def _get_window_pid(self, hwnd: int) -> int:
        pid = ctypes.c_ulong(0)
        ctypes.windll.user32.GetWindowThreadProcessId(wintypes.HWND(hwnd), ctypes.byref(pid))
        return int(pid.value)

    def _get_window_title(self, hwnd: int) -> str:
        length = ctypes.windll.user32.GetWindowTextLengthW(wintypes.HWND(hwnd))
        if length <= 0:
            return ""
        buffer = ctypes.create_unicode_buffer(length + 1)
        ctypes.windll.user32.GetWindowTextW(wintypes.HWND(hwnd), buffer, length + 1)
        return buffer.value.strip()

    def _close_window(self, hwnd: int) -> bool:
        WM_CLOSE = 0x0010
        result = bool(ctypes.windll.user32.PostMessageW(wintypes.HWND(hwnd), WM_CLOSE, 0, 0))
        if not result:
            result = bool(ctypes.windll.user32.ShowWindow(wintypes.HWND(hwnd), 0))
        return result

    def _is_excel_login_window(self, title: str) -> bool:
        lowered = title.lower()
        return any(
            keyword in lowered
            for keyword in (
                "inicia sesión",
                "iniciar sesión",
                "sign in",
                "comenzar a usar excel",
            )
        )

    def get_sheet_names(self) -> list[str]:
        return [sheet.Name for sheet in self.workbook.Worksheets]

    def _replace_unsupported_xlookup_formulas(self) -> None:
        if self.workbook is None:
            return

        worksheet = self.workbook.Worksheets(EXCEL_SHEET_LABEL)
        used_range = worksheet.UsedRange
        rows = used_range.Rows.Count
        cols = used_range.Columns.Count
        replaced = 0

        for row in range(1, rows + 1):
            for col in range(1, cols + 1):
                cell = used_range.Cells(row, col)
                try:
                    formula = str(cell.Formula or "")
                except Exception:
                    continue
                if "_xlfn.XLOOKUP" not in formula:
                    continue
                converted = self._convert_xlookup_formula(formula)
                if not converted:
                    continue
                try:
                    cell.Formula = converted
                    replaced += 1
                except Exception:
                    continue

        if replaced:
            self.logger.info("Se reemplazaron %s fórmula(s) XLOOKUP incompatibles por fórmulas compatibles.", replaced)

    @staticmethod
    def _convert_xlookup_formula(formula: str) -> str | None:
        normalized = formula.strip()
        if not normalized:
            return None

        if normalized.startswith("="):
            normalized = normalized[1:]
        if normalized.startswith("+"):
            normalized = normalized[1:]

        prefix = "_xlfn.XLOOKUP("
        if not normalized.upper().startswith(prefix.upper()) or not normalized.endswith(")"):
            return None

        inner = normalized[len(prefix) : -1]
        parts = [part.strip() for part in inner.split(",", 3)]
        if len(parts) != 4:
            return None

        lookup_value, lookup_array, return_array, default_value = parts
        if not default_value:
            default_value = '""'

        return f"=IFERROR(INDEX({return_array},MATCH({lookup_value},{lookup_array},0)),{default_value})"

    def get_headers(self) -> list[str]:
        worksheet = self.workbook.Worksheets(EXCEL_SHEET_SOURCE)
        return [str(worksheet.Cells(1, column).Value).strip() for column in range(1, 4)]

    def get_filters(self) -> list[str]:
        worksheet = self.workbook.Worksheets(EXCEL_SHEET_SOURCE)
        last_row = self._get_last_used_row(worksheet)

        values: set[str] = set()
        for row in range(2, last_row + 1):
            value = worksheet.Cells(row, 3).Value
            normalized = self._normalize_excel_text(value)
            if normalized:
                values.add(normalized)

        return sorted(values)

    def get_filtered_records(self, selected_filter: str) -> list[AssetRecord]:
        worksheet = self.workbook.Worksheets(EXCEL_SHEET_SOURCE)
        last_row = self._get_last_used_row(worksheet)
        records: list[AssetRecord] = []

        for row in range(2, last_row + 1):
            section_value = self._normalize_excel_text(worksheet.Cells(row, 3).Value)
            if section_value != selected_filter:
                continue

            records.append(
                AssetRecord(
                    row_index=row,
                    asset_id=self._normalize_excel_identifier(worksheet.Cells(row, 1).Value),
                    asset_name=self._normalize_excel_text(worksheet.Cells(row, 2).Value),
                    section=section_value,
                )
            )

        return records

    def write_block_to_label_sheet(self, records: list[AssetRecord]) -> None:
        worksheet = self.workbook.Worksheets(EXCEL_SHEET_LABEL)
        self.clear_label_input_area()
        worksheet.Range(f"{LABEL_COLUMNS['asset']}{LABEL_OUTPUT_START_ROW}:{LABEL_COLUMNS['asset']}{LABEL_OUTPUT_START_ROW + BLOCK_SIZE - 1}").NumberFormat = "@"

        for offset, record in enumerate(records, start=LABEL_OUTPUT_START_ROW):
            worksheet.Range(f"{LABEL_COLUMNS['asset']}{offset}").Value = record.asset_id

        self.excel_app.CalculateFull()

    def clear_label_input_area(self) -> None:
        worksheet = self.workbook.Worksheets(EXCEL_SHEET_LABEL)
        start_row = LABEL_OUTPUT_START_ROW
        end_row = LABEL_OUTPUT_START_ROW + BLOCK_SIZE - 1
        worksheet.Range(f"{LABEL_COLUMNS['asset']}{start_row}:{LABEL_COLUMNS['asset']}{end_row}").ClearContents()

    def get_generated_assets(self, count: int) -> list[str]:
        worksheet = self.workbook.Worksheets(EXCEL_SHEET_LABEL)
        return [
            self._normalize_excel_identifier(worksheet.Range(f"{LABEL_COLUMNS['asset']}{row}").Value)
            for row in range(LABEL_OUTPUT_START_ROW, LABEL_OUTPUT_START_ROW + count)
        ]

    def copy_label_shape(self, position: int) -> str:
        worksheet = self.workbook.Worksheets(EXCEL_SHEET_LABEL)
        try:
            group_name = EXCEL_LABEL_GROUP_NAMES[position]
        except IndexError as exc:
            raise RuntimeError(f"No hay grupo Excel configurado para la posiciÃ³n {position + 1}.") from exc

        try:
            shape = worksheet.Shapes.Item(group_name)
        except Exception as exc:
            raise RuntimeError(f"No existe el grupo/shape de etiqueta en Excel: {group_name}") from exc

        shape.CopyPicture(Appearance=1, Format=2)
        return group_name

    def export_label_shape_image(
        self,
        position: int,
        output_png: str | Path,
        target_px: tuple[int, int] = (LABEL_IMAGE_WIDTH_PX, LABEL_IMAGE_HEIGHT_PX),
    ) -> LabelImageExportResult:
        if Image is None or ImageGrab is None:
            raise RuntimeError("Pillow no estÃ¡ instalado. InstalÃ¡ las dependencias desde requirements.txt.")

        worksheet = self.workbook.Worksheets(EXCEL_SHEET_LABEL)
        try:
            group_name = EXCEL_LABEL_GROUP_NAMES[position]
        except IndexError as exc:
            raise RuntimeError(f"No hay grupo Excel configurado para la posiciÃ³n {position + 1}.") from exc

        try:
            shape = worksheet.Shapes.Item(group_name)
        except Exception as exc:
            raise RuntimeError(f"No existe el grupo/shape de etiqueta en Excel: {group_name}") from exc

        image = self._render_label_group_with_pillow(shape, target_px)
        original_size = image.size

        output_path = Path(output_png)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        image.convert("RGBA").save(output_path, format="PNG")

        return LabelImageExportResult(
            position=position + 1,
            group_name=group_name,
            output_path=output_path,
            original_size_px=original_size,
            target_size_px=target_px,
        )

    def _grab_copied_shape_image(self, group_name: str):
        errors: list[Exception] = []
        for attempt in range(1, 6):
            try:
                time.sleep(0.2 * attempt)
                clipboard_image = ImageGrab.grabclipboard()
                if clipboard_image is None:
                    raise RuntimeError("el portapapeles todavÃ­a no contiene imagen")
                if isinstance(clipboard_image, list):
                    raise RuntimeError("el portapapeles devolviÃ³ archivos, no una imagen")
                return clipboard_image
            except Exception as exc:
                errors.append(exc)

        raise RuntimeError(
            f"Excel copiÃ³ {group_name}, pero no se pudo capturar la imagen desde el portapapeles. "
            f"Ãšltimo error: {errors[-1]}"
        )

    def _render_label_group_with_pillow(self, shape, target_px: tuple[int, int]):
        """Render the Excel label from the group's live contents instead of CopyPicture.

        Excel's `Shape.CopyPicture` returns a valid bitmap size for these grouped
        labels, but the bitmap content is blank on this workbook/COM setup. The
        group still exposes the calculated text boxes and embedded logo reliably,
        so we render those contents into the configured full PNG ourselves.
        """
        if Image is None or ImageDraw is None or ImageFont is None:
            raise RuntimeError("Pillow no estÃ¡ instalado. InstalÃ¡ las dependencias desde requirements.txt.")

        width_px, height_px = target_px
        image = Image.new("RGBA", target_px, (255, 255, 255, 255))
        draw = ImageDraw.Draw(image)
        scale_x = width_px / float(shape.Width)
        scale_y = height_px / float(shape.Height)
        group_left = float(shape.Left)
        group_top = float(shape.Top)

        # Borde exterior real de la etiqueta. Se dibuja completo para evitar el
        # corte/desfase que producÃ­a la captura del shape agrupado.
        draw.rectangle((0, 0, width_px - 1, height_px - 1), outline=(220, 220, 220, 255), width=1)

        for index in range(1, shape.GroupItems.Count + 1):
            item = shape.GroupItems.Item(index)
            box = self._shape_item_box_px(item, group_left, group_top, scale_x, scale_y)
            if int(item.Type) == 13:  # msoPicture
                logo = self._copy_picture_item_to_image(item)
                if logo is not None:
                    logo = logo.convert("RGBA").resize((max(1, box[2] - box[0]), max(1, box[3] - box[1])), self._pil_resize_filter())
                    image.alpha_composite(logo, (box[0], box[1]))
                continue

            text = self._shape_item_text(item)
            if not text or self._is_hidden_text(item):
                continue

            font_size = self._shape_item_font_size_px(item, scale_y)
            bold = self._shape_item_is_bold(item)
            font = self._load_calibri_font(font_size, bold=bold)
            self._draw_centered_text(draw, text, box, font, fill=(0, 0, 0, 255))

        return image

    def _shape_item_box_px(self, item, group_left: float, group_top: float, scale_x: float, scale_y: float) -> tuple[int, int, int, int]:
        left = int(round((float(item.Left) - group_left) * scale_x))
        top = int(round((float(item.Top) - group_top) * scale_y))
        right = int(round((float(item.Left) + float(item.Width) - group_left) * scale_x))
        bottom = int(round((float(item.Top) + float(item.Height) - group_top) * scale_y))
        return (left, top, max(left + 1, right), max(top + 1, bottom))

    def _copy_picture_item_to_image(self, item):
        errors: list[Exception] = []
        for attempt in range(1, 4):
            try:
                item.CopyPicture(Appearance=1, Format=2)
                time.sleep(0.15 * attempt)
                clipboard_image = ImageGrab.grabclipboard()
                if clipboard_image is None or isinstance(clipboard_image, list):
                    raise RuntimeError("el portapapeles no devolviÃ³ la imagen del logo")
                return clipboard_image
            except Exception as exc:
                errors.append(exc)
        return None

    def _shape_item_text(self, item) -> str:
        try:
            return str(item.TextFrame2.TextRange.Text or "").strip()
        except Exception:
            return ""

    def _is_hidden_text(self, item) -> bool:
        try:
            return int(item.TextFrame2.TextRange.Font.Fill.ForeColor.RGB) == 16777215
        except Exception:
            return False

    def _shape_item_font_size_px(self, item, scale_y: float) -> int:
        try:
            size_pt = float(item.TextFrame2.TextRange.Font.Size)
        except Exception:
            size_pt = 10.0
        return max(7, int(round(size_pt * scale_y)))

    def _shape_item_is_bold(self, item) -> bool:
        try:
            return int(item.TextFrame2.TextRange.Font.Bold) != 0
        except Exception:
            return False

    def _load_calibri_font(self, size_px: int, bold: bool = False):
        font_name = "calibrib.ttf" if bold else "calibri.ttf"
        font_path = Path("C:/Windows/Fonts") / font_name
        try:
            return ImageFont.truetype(str(font_path), size_px)
        except Exception:
            return ImageFont.load_default()

    def _draw_centered_text(self, draw, text: str, box: tuple[int, int, int, int], font, fill) -> None:
        left, top, right, bottom = box
        padding_x = 2
        padding_y = 1
        max_width = max(1, right - left - (padding_x * 2))
        max_height = max(1, bottom - top - (padding_y * 2))
        lines = self._wrap_text_to_width(draw, text, font, max_width)
        line_gap = 0
        line_metrics = [draw.textbbox((0, 0), line, font=font) for line in lines]
        total_height = sum(metric[3] - metric[1] for metric in line_metrics) + (max(0, len(lines) - 1) * line_gap)

        while total_height > max_height and getattr(font, "size", 8) > 7:
            font = self._load_calibri_font(getattr(font, "size", 8) - 1, bold=False)
            lines = self._wrap_text_to_width(draw, text, font, max_width)
            line_metrics = [draw.textbbox((0, 0), line, font=font) for line in lines]
            total_height = sum(metric[3] - metric[1] for metric in line_metrics) + (max(0, len(lines) - 1) * line_gap)

        y = top + padding_y + max(0, (max_height - total_height) // 2)
        for line, metric in zip(lines, line_metrics):
            line_width = metric[2] - metric[0]
            line_height = metric[3] - metric[1]
            x = left + padding_x + max(0, (max_width - line_width) // 2)
            draw.text((x, y - metric[1]), line, font=font, fill=fill)
            y += line_height + line_gap

    def _wrap_text_to_width(self, draw, text: str, font, max_width: int) -> list[str]:
        words = text.split()
        if not words:
            return [""]
        lines: list[str] = []
        current = words[0]
        for word in words[1:]:
            candidate = f"{current} {word}"
            if draw.textbbox((0, 0), candidate, font=font)[2] <= max_width:
                current = candidate
            else:
                lines.append(current)
                current = word
        lines.append(current)
        return lines

    @staticmethod
    def _pil_resize_filter():
        return getattr(getattr(Image, "Resampling", Image), "LANCZOS")

    def _get_last_used_row(self, worksheet) -> int:
        used_range = worksheet.UsedRange
        return used_range.Row + used_range.Rows.Count - 1

    @staticmethod
    def _normalize_excel_text(value) -> str:
        return normalize_excel_scalar(value)

    @staticmethod
    def _normalize_excel_identifier(value) -> str:
        return normalize_excel_scalar(value)

    def audit_identifier_column(self) -> list[tuple[int, str, str]]:
        return self.normalize_identifier_column()

    def normalize_identifier_column(self) -> list[tuple[int, str, str]]:
        worksheet = self.workbook.Worksheets(EXCEL_SHEET_SOURCE)
        last_row = self._get_last_used_row(worksheet)
        issues: list[tuple[int, str, str]] = []

        if last_row < 2:
            return issues

        worksheet.Range(f"A2:A{last_row}").NumberFormat = "@"

        for row in range(2, last_row + 1):
            raw_value = worksheet.Cells(row, 1).Value
            normalized = self._normalize_excel_identifier(raw_value)
            raw_text = "" if raw_value is None else str(raw_value)
            if raw_text != normalized:
                issues.append((row, raw_text, normalized))
            try:
                worksheet.Cells(row, 1).Value = normalized
            except Exception:
                # No interrumpir el proceso si una celda puntual no admite la reasignaciÃ³n.
                pass

        return issues
