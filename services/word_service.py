from __future__ import annotations

import re
import time
from dataclasses import dataclass
from pathlib import Path

try:
    from docx import Document
    from docx.shared import Cm
except ImportError:  # pragma: no cover
    Document = None
    Cm = None

from utils.constants import WORD_LABEL_HEIGHT_CM, WORD_LABEL_WIDTH_CM

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
class VisualValidationResult:
    position: int
    slot_name: str
    expected_width_cm: float
    expected_height_cm: float
    applied_width_cm: float
    applied_height_cm: float
    container_width_cm: float | None
    container_height_cm: float | None
    adjusted: bool
    details: str
    shape_name: str
    page_number: int
    created_slot: bool


@dataclass(slots=True)
class BlankPageCleanupResult:
    pages_before: int
    pages_after: int
    removed_pages: int
    removed_page_numbers: list[int]
    cleaned_placeholders: list[str]
    remaining_placeholders: list[str]
    expected_last_image_page: int | None = None
    validation_passes: int = 0


@dataclass(slots=True)
class ImageCountValidationResult:
    expected_count: int
    detected_count: int
    missing_count: int
    extra_count: int
    missing_positions: list[int]
    repaired_positions: list[int]


class WordService:
    MAX_PLACEHOLDER_IMAGES = 27
    POINTS_PER_CM = 28.3464567

    def __init__(self) -> None:
        self.word_app = None
        self.document = None
        self._win32_module = None
        self._word_types_primed = False
        self._released_to_user = False
        self._com_initialized = False
        self.last_layout_details = "plantilla Word no inspeccionada"

    def open(self, template_path: str, visible: bool = False) -> None:
        self._open_word_document(template_path, visible=visible, create_from_template=False)

    def create_from_template(self, template_path: str, visible: bool = False) -> None:
        self._open_word_document(template_path, visible=visible, create_from_template=True)

    def close(self, save_changes: bool = False) -> None:
        if self._released_to_user:
            self.document = None
            self.word_app = None
            self._released_to_user = False
            return

        self._teardown(save_changes=save_changes)

    def prepare_placeholder_document(self) -> None:
        if self.document is None:
            raise RuntimeError("Word no está abierto.")

        remaining = self.find_remaining_image_placeholders()
        self.last_layout_details = (
            "plantilla Word COM lista para reemplazar placeholders "
            f"<img1>...<img{self.MAX_PLACEHOLDER_IMAGES}>; pendientes={len(remaining)}"
        )

    def get_label_layout_details(self) -> str:
        return self.last_layout_details

    def replace_image_placeholder(self, placeholder_number: int, image_path: str | Path) -> VisualValidationResult:
        if self.document is None:
            raise RuntimeError("Word no está abierto.")
        if placeholder_number < 1 or placeholder_number > self.MAX_PLACEHOLDER_IMAGES:
            raise RuntimeError(
                f"Placeholder fuera de rango: <img{placeholder_number}>. "
                f"La plantilla soporta <img1>...<img{self.MAX_PLACEHOLDER_IMAGES}>."
            )

        placeholder = self._placeholder_text(placeholder_number)
        try:
            placeholder_range = self._find_exact_text_range(placeholder)
        except ModuleNotFoundError:
            if self._win32_module is None:
                raise
            self._prime_word_types(self._win32_module)
            placeholder_range = self._find_exact_text_range(placeholder)
        if placeholder_range is None:
            raise RuntimeError(f"No se encontró el placeholder {placeholder} en la plantilla Word.")

        expected_width = WORD_LABEL_WIDTH_CM * self.POINTS_PER_CM
        expected_height = WORD_LABEL_HEIGHT_CM * self.POINTS_PER_CM
        insert_start = int(placeholder_range.Start)
        page_number = self._range_page_number(placeholder_range)
        placeholder_range.Text = ""
        insert_range = self.document.Range(insert_start, insert_start)

        try:
            picture = self.document.InlineShapes.AddPicture(
                    FileName=str(Path(image_path).resolve()),
                    LinkToFile=False,
                    SaveWithDocument=True,
                    Range=insert_range,
                )
        except ModuleNotFoundError:
            if self._win32_module is None:
                raise
            self._prime_word_types(self._win32_module)
            try:
                picture = self.document.InlineShapes.AddPicture(
                        FileName=str(Path(image_path).resolve()),
                        LinkToFile=False,
                        SaveWithDocument=True,
                        Range=insert_range,
                    )
            except Exception as exc:
                raise RuntimeError(f"Word no pudo reemplazar {placeholder} por una imagen real: {exc}") from exc
        except Exception as exc:
            raise RuntimeError(f"Word no pudo reemplazar {placeholder} por una imagen real: {exc}") from exc

        try:
            picture.LockAspectRatio = 0
            picture.Width = expected_width
            picture.Height = expected_height
        except Exception:
            pass

        return VisualValidationResult(
            position=placeholder_number,
            slot_name=placeholder,
            expected_width_cm=WORD_LABEL_WIDTH_CM,
            expected_height_cm=WORD_LABEL_HEIGHT_CM,
            applied_width_cm=self._points_to_cm(getattr(picture, "Width", expected_width)),
            applied_height_cm=self._points_to_cm(getattr(picture, "Height", expected_height)),
            container_width_cm=None,
            container_height_cm=None,
            adjusted=True,
            details="placeholder de texto reemplazado por InlineShapes.AddPicture",
            shape_name="",
            page_number=page_number,
            created_slot=False,
        )

    def clear_unused_image_placeholders(self, used_count: int, max_count: int | None = None) -> list[str]:
        if self.document is None:
            raise RuntimeError("Word no está abierto.")

        max_count = max_count or self.MAX_PLACEHOLDER_IMAGES
        cleaned: list[str] = []
        for placeholder_number in range(used_count + 1, max_count + 1):
            placeholder = self._placeholder_text(placeholder_number)
            self._replace_all_exact_text(placeholder, "")
            cleaned.append(placeholder)
        return cleaned

    def validate_embedded_image_count(self, expected_count: int) -> ImageCountValidationResult:
        if self.document is None:
            raise RuntimeError("Word no está abierto.")
        if expected_count < 0:
            raise RuntimeError("El conteo esperado de imágenes no puede ser negativo.")

        detected_count = self.count_visible_images()
        missing_count = max(expected_count - detected_count, 0)
        extra_count = max(detected_count - expected_count, 0)
        missing_positions = list(range(detected_count + 1, expected_count + 1)) if missing_count else []
        return ImageCountValidationResult(expected_count, detected_count, missing_count, extra_count, missing_positions, [])

    def count_visible_images(self) -> int:
        if self.document is None:
            raise RuntimeError("Word no está abierto.")

        self._repaginate()
        count = 0
        try:
            for index in range(1, self.document.InlineShapes.Count + 1):
                inline_shape = self.document.InlineShapes.Item(index)
                if int(getattr(inline_shape, "Type", 0) or 0) in {3, 4}:
                    count += 1
        except Exception:
            pass
        return count

    def cleanup_blank_pages(
        self,
        expected_last_image_page: int | None = None,
        validation_passes: int = 3,
        cleaned_placeholders: list[str] | None = None,
    ) -> BlankPageCleanupResult:
        if self.document is None:
            raise RuntimeError("Word no está abierto.")

        cleaned_placeholders = cleaned_placeholders or []
        pages_before = self.get_page_count()
        passes_done = 0
        removed_pages: list[int] = []

        for _ in range(max(1, validation_passes)):
            passes_done += 1
            removed_in_pass = False
            self._trim_trailing_blank_paragraphs()
            self._repaginate()

            current_pages = self.get_page_count()
            for page_number in range(current_pages, 0, -1):
                page_range = self._page_range(page_number)
                if page_range is None:
                    continue

                is_after_last_image = expected_last_image_page is not None and page_number > expected_last_image_page
                removable_empty_tail = self._range_is_blank_or_placeholder_only(page_range)

                if not (is_after_last_image or removable_empty_tail):
                    continue

                if self._delete_blank_page_range(page_range, page_number):
                    removed_pages.append(page_number)
                    removed_in_pass = True
                    self._repaginate()

            if not removed_in_pass:
                break

        self._repaginate()
        try:
            remaining_placeholders = self.find_remaining_image_placeholders()
        except Exception:
            remaining_placeholders = []

        return BlankPageCleanupResult(
            pages_before=pages_before,
            pages_after=self.get_page_count(),
            removed_pages=len(removed_pages),
            removed_page_numbers=sorted(set(removed_pages)),
            cleaned_placeholders=cleaned_placeholders,
            remaining_placeholders=remaining_placeholders,
            expected_last_image_page=expected_last_image_page,
            validation_passes=passes_done,
        )

    def find_remaining_image_placeholders(self) -> list[str]:
        if self.document is None:
            raise RuntimeError("Word no está abierto.")

        remaining: list[str] = []
        try:
            for placeholder_number in range(1, self.MAX_PLACEHOLDER_IMAGES + 1):
                placeholder = self._placeholder_text(placeholder_number)
                if self._contains_exact_text(placeholder):
                    remaining.append(placeholder)
        except Exception:
            return []
        return remaining

    def get_page_count(self) -> int:
        if self.document is None:
            raise RuntimeError("Word no está abierto.")
        self._repaginate()
        try:
            return int(self.document.ComputeStatistics(2))
        except Exception:
            return 1

    def print_document(self, printer_name: str) -> None:
        if self.word_app is None or self.document is None:
            raise RuntimeError("Word no está abierto.")

        current_printer = self.word_app.ActivePrinter
        self.word_app.ActivePrinter = printer_name
        try:
            self.document.PrintOut(Background=False)
        finally:
            self.word_app.ActivePrinter = current_printer

    def save_document_copy(self, output_path: str) -> None:
        if self.document is None:
            raise RuntimeError("Word no está abierto.")

        output_file = Path(output_path).resolve()
        output_file.parent.mkdir(parents=True, exist_ok=True)
        try:
            self.document.SaveAs2(str(output_file))
        except Exception:
            self.document.SaveAs(str(output_file))

    def show_to_user(self) -> None:
        if self.word_app is None or self.document is None:
            raise RuntimeError("Word no está abierto.")
        self.word_app.Visible = True
        # Keep alerts suppressed here too; Word sign-in / first-run prompts
        # are disruptive during manual review.
        self.word_app.DisplayAlerts = 0
        try:
            self.word_app.Options.PrintDrawingObjects = True
        except Exception:
            pass
        try:
            view = self.word_app.ActiveWindow.View
            view.ShowDrawings = True
            view.ShowPicturePlaceHolders = False
            view.Type = 3  # wdPrintView
        except Exception:
            pass
        self.document.Activate()

    def release_to_user(self) -> None:
        self._released_to_user = True
        self.document = None
        self.word_app = None

    def export_pdf(self, output_path: str) -> None:
        if self.document is None:
            raise RuntimeError("Word no está abierto.")

        wd_format_pdf = 17
        output_file = Path(output_path).resolve()
        output_file.parent.mkdir(parents=True, exist_ok=True)
        try:
            self.document.SaveAs2(str(output_file), FileFormat=wd_format_pdf)
        except Exception:
            self.document.SaveAs(str(output_file), FileFormat=wd_format_pdf)

    def build_document_without_com(
        self,
        template_path: str,
        output_path: str,
        image_paths: list[Path],
        placeholder_count: int,
    ) -> None:
        if Document is None or Cm is None:
            raise RuntimeError("python-docx no está instalado. Instala la dependencia para usar el fallback sin COM.")

        template_file = Path(template_path).resolve()
        output_file = Path(output_path).resolve()

        if not template_file.exists():
            raise FileNotFoundError(f"La plantilla Word no existe: {template_file}")

        doc = Document(str(template_file))
        self._replace_placeholders_without_com(doc, image_paths, placeholder_count)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(output_file))

    def _replace_placeholders_without_com(self, doc, image_paths: list[Path], placeholder_count: int) -> None:
        placeholder_pattern = re.compile(r"<img(\d+)>")

        for paragraph in self._iter_docx_paragraphs(doc):
            text = paragraph.text
            if "<img" not in text:
                continue

            paragraph.clear()
            cursor = 0
            for match in placeholder_pattern.finditer(text):
                if match.start() > cursor:
                    paragraph.add_run(text[cursor:match.start()])

                number = int(match.group(1))
                if 1 <= number <= placeholder_count and number <= len(image_paths):
                    run = paragraph.add_run()
                    run.add_picture(
                        str(image_paths[number - 1]),
                        width=Cm(WORD_LABEL_WIDTH_CM),
                        height=Cm(WORD_LABEL_HEIGHT_CM),
                    )

                cursor = match.end()

            if cursor < len(text):
                paragraph.add_run(text[cursor:])

    def _iter_docx_paragraphs(self, doc):
        yield from doc.paragraphs
        for table in doc.tables:
            yield from self._iter_table_paragraphs(table)
        for section in doc.sections:
            yield from section.header.paragraphs
            for table in section.header.tables:
                yield from self._iter_table_paragraphs(table)
            yield from section.footer.paragraphs
            for table in section.footer.tables:
                yield from self._iter_table_paragraphs(table)

    def _iter_table_paragraphs(self, table):
        for row in table.rows:
            for cell in row.cells:
                yield from cell.paragraphs
                for nested_table in cell.tables:
                    yield from self._iter_table_paragraphs(nested_table)

    def _open_word_document(self, template_path: str, *, visible: bool, create_from_template: bool) -> None:
        try:
            pythoncom_module, win32_module = _load_pywin32()
        except ImportError as exc:  # pragma: no cover
            raise RuntimeError("pywin32 no está instalado.") from exc

        target_path = str(Path(template_path).resolve())
        last_error = None
        for attempt in range(1, 4):
            pythoncom_module.CoInitialize()
            self._com_initialized = True
            self.word_app = None
            self.document = None
            self._win32_module = win32_module
            self._released_to_user = False
            try:
                if not self._word_types_primed:
                    self._prime_word_types(win32_module)
                    self._word_types_primed = True
                self.word_app = win32_module.dynamic.Dispatch("Word.Application")
                self.word_app.Visible = visible
                self.word_app.DisplayAlerts = 0
                try:
                    self.word_app.AutomationSecurity = 3
                except Exception:
                    pass

                if create_from_template:
                    self.document = self.word_app.Documents.Add(
                        Template=target_path,
                        NewTemplate=False,
                        DocumentType=0,
                        Visible=visible,
                    )
                else:
                    self.document = self.word_app.Documents.Open(target_path, False, False, False)
                return
            except Exception as exc:
                last_error = exc
                self._teardown(save_changes=False)
                if attempt < 3:
                    time.sleep(attempt)

        raise last_error

    def _prime_word_types(self, win32_module) -> None:
        try:
            primer = win32_module.dynamic.Dispatch("Word.Application")
            primer.Visible = False
            primer.DisplayAlerts = 0
            temp_doc = primer.Documents.Add()
            _ = temp_doc.Content
            _ = temp_doc.Content.Find
            _ = temp_doc.Content.Find.Replacement
            _ = temp_doc.Range(0, 0)
            _ = temp_doc.InlineShapes
            temp_doc.Close(SaveChanges=False)
            primer.Quit(False)
        except Exception:
            pass

    def _teardown(self, save_changes: bool = False) -> None:
        document = self.document
        word_app = self.word_app
        com_initialized = self._com_initialized

        self.document = None
        self.word_app = None
        self._released_to_user = False
        self._com_initialized = False

        try:
            if document is not None:
                document.Close(SaveChanges=save_changes)
        except Exception:
            pass

        try:
            if word_app is not None:
                word_app.Quit(SaveChanges=save_changes)
        except Exception:
            pass

        # Intentionally do not CoUninitialize here.
        # The worker keeps Excel and Word COM alive in the same apartment;
        # per-service uninitialization can disconnect the other proxy mid-run.

    def _placeholder_text(self, placeholder_number: int) -> str:
        return f"<img{placeholder_number}>"

    def _trim_trailing_blank_paragraphs(self) -> None:
        try:
            while self.document.Paragraphs.Count > 0:
                paragraph = self.document.Paragraphs(self.document.Paragraphs.Count)
                text = str(getattr(paragraph.Range, "Text", "") or "")
                if text.strip(" \t\r\n\f\v\a\x07\u00a0"):
                    break
                paragraph.Range.Delete()
        except Exception:
            pass

    def _page_range(self, page_number: int):
        try:
            start_range = self.document.GoTo(What=1, Which=1, Count=page_number)
            start = int(start_range.Start)
            page_count = self.get_page_count()
            if page_number < page_count:
                next_range = self.document.GoTo(What=1, Which=1, Count=page_number + 1)
                end = int(next_range.Start)
            else:
                end = int(self.document.Content.End)
            if end <= start:
                return None
            return self.document.Range(start, end)
        except Exception:
            return None

    def _range_is_blank_or_placeholder_only(self, page_range) -> bool:
        if self._range_contains_images(page_range):
            return False

        text = str(getattr(page_range, "Text", "") or "")
        for placeholder_number in range(1, self.MAX_PLACEHOLDER_IMAGES + 1):
            text = text.replace(self._placeholder_text(placeholder_number), "")
        return text.strip(" \t\r\n\f\v\a\x07\u00a0") == ""

    def _range_contains_images(self, page_range) -> bool:
        try:
            if int(getattr(page_range.InlineShapes, "Count", 0) or 0) > 0:
                return True
        except Exception:
            pass

        try:
            shape_range = getattr(page_range, "ShapeRange", None)
            if shape_range is not None and int(getattr(shape_range, "Count", 0) or 0) > 0:
                return True
        except Exception:
            pass

        return False

    def _delete_blank_page_range(self, page_range, page_number: int) -> bool:
        try:
            page_count = self.get_page_count()
            if page_count <= 1:
                page_range.Text = ""
            else:
                page_range.Delete()
            return True
        except Exception:
            return False

    def _find_exact_text_range(self, text: str):
        search_range = self.document.Content
        try:
            find = search_range.Find
        except ModuleNotFoundError:
            if self._win32_module is None:
                raise
            self._prime_word_types(self._win32_module)
            search_range = self.document.Content
            find = search_range.Find
        find.ClearFormatting()
        find.Text = text
        find.Forward = True
        find.Wrap = 0
        find.MatchCase = True
        find.MatchWholeWord = False
        find.MatchWildcards = False
        if find.Execute():
            return search_range.Duplicate
        return None

    def _replace_all_exact_text(self, text: str, replacement: str) -> None:
        search_range = self.document.Content
        try:
            find = search_range.Find
        except ModuleNotFoundError:
            if self._win32_module is None:
                raise
            self._prime_word_types(self._win32_module)
            search_range = self.document.Content
            find = search_range.Find
        find.ClearFormatting()
        find.Execute(
            FindText=text,
            ReplaceWith=replacement,
            Replace=2,
            Forward=True,
            Wrap=0,
            MatchCase=True,
            MatchWholeWord=False,
            MatchWildcards=False,
        )

    def _contains_exact_text(self, text: str) -> bool:
        try:
            return text in str(getattr(self.document.Content, "Text", "") or "")
        except Exception:
            return False

    def _range_page_number(self, word_range) -> int:
        try:
            return int(word_range.Information(3))
        except Exception:
            return 1

    def _repaginate(self) -> None:
        try:
            self.document.Repaginate()
        except Exception:
            pass

    def _points_to_cm(self, value: float | None) -> float | None:
        if value is None:
            return None
        return round(float(value) / self.POINTS_PER_CM, 3)
