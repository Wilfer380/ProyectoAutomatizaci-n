from __future__ import annotations

from dataclasses import dataclass
import time
from pathlib import Path
import time

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
    POINTS_PER_CM = 28.3464567
    TRUSTED_SLOT_PREFIX = "ATL_LabelSlot_"
    SLOT_SIZE_TOLERANCE = 0.20
    MAX_PLACEHOLDER_IMAGES = 27

    def __init__(self) -> None:
        self.word_app = None
        self.document = None
        self._released_to_user = False
        self._slots: list[dict] = []
        self._current_slot_index = 0
        self._initial_slot_count = 0
        self.last_layout_details = "plantilla Word no inspeccionada"

    def open(self, template_path: str, visible: bool = False) -> None:
        try:
            pythoncom_module, win32_module = _load_pywin32()
        except ImportError as exc:
            raise RuntimeError("pywin32 no estÃƒÂ¡ instalado.") from exc

        target_path = str(Path(template_path).resolve())
        last_error = None
        for attempt in range(1, 4):
            pythoncom_module.CoInitialize()
            self.word_app = None
            self.document = None
            self._released_to_user = False
            try:
                self.word_app = win32_module.DispatchEx("Word.Application")
                self.word_app.Visible = visible
                self.word_app.DisplayAlerts = 0
                try:
                    self.word_app.AutomationSecurity = 3
                except Exception:
                    pass
                self.document = self.word_app.Documents.Open(target_path, False, False, False)
                return
            except Exception as exc:
                last_error = exc
                failed_document = self.document
                failed_word_app = self.word_app
                self.document = None
                self.word_app = None
                try:
                    if failed_document is not None:
                        failed_document.Close(SaveChanges=False)
                except Exception:
                    pass
                try:
                    if failed_word_app is not None:
                        failed_word_app.Quit(SaveChanges=False)
                except Exception:
                    pass
                if attempt < 3:
                    time.sleep(attempt)

        raise last_error

    def create_from_template(self, template_path: str, visible: bool = False) -> None:
        try:
            pythoncom_module, win32_module = _load_pywin32()
        except ImportError as exc:
            raise RuntimeError("pywin32 no estÃƒÂ¡ instalado.") from exc

        target_path = str(Path(template_path).resolve())
        last_error = None
        for attempt in range(1, 4):
            pythoncom_module.CoInitialize()
            self.word_app = None
            self.document = None
            self._released_to_user = False
            try:
                self.word_app = win32_module.DispatchEx("Word.Application")
                self.word_app.Visible = visible
                self.word_app.DisplayAlerts = 0
                try:
                    self.word_app.AutomationSecurity = 3
                except Exception:
                    pass
                self.document = self.word_app.Documents.Add(
                    Template=target_path,
                    NewTemplate=False,
                    DocumentType=0,
                    Visible=visible,
                )
                return
            except Exception as exc:
                last_error = exc
                failed_document = self.document
                failed_word_app = self.word_app
                self.document = None
                self.word_app = None
                try:
                    if failed_document is not None:
                        failed_document.Close(SaveChanges=False)
                except Exception:
                    pass
                try:
                    if failed_word_app is not None:
                        failed_word_app.Quit(SaveChanges=False)
                except Exception:
                    pass
                if attempt < 3:
                    time.sleep(attempt)

        raise last_error

    def close(self, save_changes: bool = False) -> None:
        if self._released_to_user:
            self.document = None
            self.word_app = None
            self._slots = []
            self._current_slot_index = 0
            self._initial_slot_count = 0
            self._released_to_user = False
            return

        document = self.document
        word_app = self.word_app
        self.document = None
        self.word_app = None
        self._slots = []
        self._current_slot_index = 0
        self._initial_slot_count = 0
        self._released_to_user = False

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

    def prepare_template_document(self) -> None:
        if self.document is None:
            raise RuntimeError("Word no estÃƒÆ’Ã‚Â¡ abierto.")

        self._slots = self._discover_slots()
        if not self._slots:
            self._neutralize_untrusted_template_placeholders()
        self._current_slot_index = 0
        self._initial_slot_count = len(self._slots)

        if self._slots:
            first = self._slots[0]
            self.last_layout_details = (
                f"plantilla real con {len(self._slots)} slot(s) existente(s); "
                f"primer slot {first['name']} mide "
                f"{self._points_to_cm(first['width']):.2f}x{self._points_to_cm(first['height']):.2f} cm; "
                f"etiqueta objetivo {WORD_LABEL_WIDTH_CM:.2f}x{WORD_LABEL_HEIGHT_CM:.2f} cm"
            )
        else:
            self.last_layout_details = (
                "plantilla real sin slots controlados confiables; se ignoran placeholders/grupos amplios "
                "y se crearÃƒÆ’Ã‚Â¡n pÃƒÆ’Ã‚Â¡ginas/slots dinÃƒÆ’Ã‚Â¡micos "
                f"de {WORD_LABEL_WIDTH_CM:.2f}x{WORD_LABEL_HEIGHT_CM:.2f} cm cuando hagan falta"
            )

    def get_slot_count(self) -> int:
        return len(self._slots)

    def get_initial_slot_count(self) -> int:
        return self._initial_slot_count

    def get_label_layout_details(self) -> str:
        return self.last_layout_details

    def prepare_placeholder_document(self) -> None:
        if self.document is None:
            raise RuntimeError("Word no estÃƒÆ’Ã‚Â¡ abierto.")

        self.last_layout_details = (
            "plantilla Word con placeholders de texto normal <img1>...<img27>; "
            "se reemplaza cada placeholder por una imagen inline real y se limpian los no usados"
        )

    def replace_image_placeholder(self, placeholder_number: int, image_path: str | Path) -> VisualValidationResult:
        if self.document is None:
            raise RuntimeError("Word no estÃƒÆ’Ã‚Â¡ abierto.")
        if placeholder_number < 1 or placeholder_number > self.MAX_PLACEHOLDER_IMAGES:
            raise RuntimeError(
                f"Placeholder fuera de rango: <img{placeholder_number}>. "
                f"La plantilla soporta <img1>...<img{self.MAX_PLACEHOLDER_IMAGES}>."
            )

        placeholder = self._placeholder_text(placeholder_number)
        placeholder_range = self._find_exact_text_range(placeholder)
        if placeholder_range is None:
            raise RuntimeError(
                f"No se encontrÃƒÆ’Ã‚Â³ el placeholder obligatorio {placeholder} en la copia de la plantilla Word. "
                "VerificÃƒÆ’Ã‚Â¡ que exista como texto normal visible y que no haya espacios internos."
            )

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
            details="placeholder de texto normal reemplazado por InlineShapes.AddPicture",
            shape_name="",
            page_number=page_number,
            created_slot=False,
        )

    def clear_unused_image_placeholders(self, used_count: int, max_count: int | None = None) -> list[str]:
        if self.document is None:
            raise RuntimeError("Word no estÃƒÆ’Ã‚Â¡ abierto.")

        max_count = max_count or self.MAX_PLACEHOLDER_IMAGES
        cleaned: list[str] = []
        for placeholder_number in range(used_count + 1, max_count + 1):
            placeholder = self._placeholder_text(placeholder_number)
            self._replace_all_exact_text(placeholder, "")
            cleaned.append(placeholder)
        return cleaned

    def validate_embedded_image_count(self, expected_count: int) -> ImageCountValidationResult:
        if self.document is None:
            raise RuntimeError("Word no estÃƒÆ’Ã‚Â¡ abierto.")
        if expected_count < 0:
            raise RuntimeError("El conteo esperado de imÃƒÆ’Ã‚Â¡genes no puede ser negativo.")

        detected_count = self.count_visible_images()
        missing_count = max(expected_count - detected_count, 0)
        extra_count = max(detected_count - expected_count, 0)
        first_missing = detected_count + 1
        missing_positions = list(range(first_missing, expected_count + 1)) if missing_count else []
        return ImageCountValidationResult(
            expected_count=expected_count,
            detected_count=detected_count,
            missing_count=missing_count,
            extra_count=extra_count,
            missing_positions=missing_positions,
            repaired_positions=[],
        )

    def count_visible_images(self) -> int:
        if self.document is None:
            raise RuntimeError("Word no estÃƒÆ’Ã‚Â¡ abierto.")

        self._repaginate()
        return self._count_visible_inline_images() + self._count_visible_floating_images()

    def append_label_image_on_new_page(self, placeholder_number: int, image_path: str | Path) -> VisualValidationResult:
        if self.document is None:
            raise RuntimeError("Word no estÃƒÆ’Ã‚Â¡ abierto.")
        if placeholder_number < 1 or placeholder_number > self.MAX_PLACEHOLDER_IMAGES:
            raise RuntimeError(
                f"Placeholder fuera de rango: <img{placeholder_number}>. "
                f"La plantilla soporta <img1>...<img{self.MAX_PLACEHOLDER_IMAGES}>."
            )

        expected_width = WORD_LABEL_WIDTH_CM * self.POINTS_PER_CM
        expected_height = WORD_LABEL_HEIGHT_CM * self.POINTS_PER_CM
        end_range = self.document.Range(self.document.Content.End - 1, self.document.Content.End - 1)
        end_range.Collapse(0)  # wdCollapseEnd
        end_range.InsertBreak(7)  # wdPageBreak
        end_range.Collapse(0)

        try:
            picture = self.document.InlineShapes.AddPicture(
                FileName=str(Path(image_path).resolve()),
                LinkToFile=False,
                SaveWithDocument=True,
                Range=end_range,
            )
        except Exception as exc:
            raise RuntimeError(f"Word no pudo reponer <img{placeholder_number}> en una nueva pÃƒÆ’Ã‚Â¡gina: {exc}") from exc

        try:
            picture.LockAspectRatio = 0
            picture.Width = expected_width
            picture.Height = expected_height
        except Exception:
            pass

        self._repaginate()
        return VisualValidationResult(
            position=placeholder_number,
            slot_name=self._placeholder_text(placeholder_number),
            expected_width_cm=WORD_LABEL_WIDTH_CM,
            expected_height_cm=WORD_LABEL_HEIGHT_CM,
            applied_width_cm=self._points_to_cm(getattr(picture, "Width", expected_width)),
            applied_height_cm=self._points_to_cm(getattr(picture, "Height", expected_height)),
            container_width_cm=None,
            container_height_cm=None,
            adjusted=True,
            details="imagen repuesta en pÃƒÆ’Ã‚Â¡gina nueva despuÃƒÆ’Ã‚Â©s de validar conteo final",
            shape_name="",
            page_number=self._range_page_number(picture.Range),
            created_slot=True,
        )

    def cleanup_blank_pages(
        self,
        expected_last_image_page: int | None = None,
        validation_passes: int = 3,
        cleaned_placeholders: list[str] | None = None,
    ) -> BlankPageCleanupResult:
        if self.document is None:
            raise RuntimeError("Word no estÃƒÆ’Ã‚Â¡ abierto.")

        cleaned_placeholders = cleaned_placeholders or []
        self._clear_all_image_placeholders()
        self._repaginate()
        pages_before = self.get_page_count()
        removed_pages: list[int] = []

        passes_done = 0
        for _ in range(max(1, validation_passes)):
            passes_done += 1
            removed_in_pass = False
            self._clear_all_image_placeholders()
            self._repaginate()
            current_pages = self.get_page_count()

            for page_number in range(current_pages, 0, -1):
                page_range = self._page_range(page_number)
                if page_range is None:
                    continue

                has_image = self._range_has_useful_image(page_range, page_number)
                is_after_last_image = expected_last_image_page is not None and page_number > expected_last_image_page
                removable_empty_tail = not has_image and self._range_is_blank_or_placeholder_only(page_range)

                if has_image:
                    continue
                if not (is_after_last_image or removable_empty_tail):
                    continue

                if self._delete_blank_page_range(page_range, page_number):
                    removed_pages.append(page_number)
                    removed_in_pass = True
                    self._repaginate()

            if not removed_in_pass:
                break

        pages_after = self.get_page_count()
        return BlankPageCleanupResult(
            pages_before=pages_before,
            pages_after=pages_after,
            removed_pages=len(removed_pages),
            removed_page_numbers=sorted(removed_pages),
            cleaned_placeholders=cleaned_placeholders,
            remaining_placeholders=self.find_remaining_image_placeholders(),
            expected_last_image_page=expected_last_image_page,
            validation_passes=passes_done,
        )

    def find_remaining_image_placeholders(self) -> list[str]:
        if self.document is None:
            raise RuntimeError("Word no estÃƒÆ’Ã‚Â¡ abierto.")

        remaining: list[str] = []
        for placeholder_number in range(1, self.MAX_PLACEHOLDER_IMAGES + 1):
            placeholder = self._placeholder_text(placeholder_number)
            if self._contains_exact_text(placeholder):
                remaining.append(placeholder)
        return remaining

    def get_page_count(self) -> int:
        if self.document is None:
            raise RuntimeError("Word no estÃƒÆ’Ã‚Â¡ abierto.")
        self._repaginate()
        try:
            return int(self.document.ComputeStatistics(2))  # wdStatisticPages
        except Exception:
            try:
                return int(self.document.Range().Information(4))  # wdNumberOfPagesInDocument
            except Exception:
                return 1

    def _placeholder_text(self, placeholder_number: int) -> str:
        return f"<img{placeholder_number}>"

    def _clear_all_image_placeholders(self) -> None:
        for placeholder_number in range(1, self.MAX_PLACEHOLDER_IMAGES + 1):
            self._replace_all_exact_text(self._placeholder_text(placeholder_number), "")

    def _repaginate(self) -> None:
        try:
            self.document.Repaginate()
        except Exception:
            pass

    def _page_range(self, page_number: int):
        try:
            start_range = self.document.GoTo(What=1, Which=1, Count=page_number)  # wdGoToPage/wdGoToAbsolute
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

    def _range_has_useful_image(self, page_range, page_number: int) -> bool:
        try:
            if int(page_range.InlineShapes.Count) > 0:
                return True
        except Exception:
            pass

        useful_picture_types = {11, 13}  # msoLinkedPicture, msoPicture
        try:
            for index in range(1, self.document.Shapes.Count + 1):
                shape = self.document.Shapes.Item(index)
                try:
                    if self._anchor_page_number(shape.Anchor) != page_number:
                        continue
                    if int(getattr(shape, "Type", 0) or 0) in useful_picture_types:
                        return True
                except Exception:
                    continue
        except Exception:
            pass
        return False

    def _count_visible_inline_images(self) -> int:
        picture_types = {3, 4}  # wdInlineShapePicture, wdInlineShapeLinkedPicture
        count = 0
        try:
            for index in range(1, self.document.InlineShapes.Count + 1):
                inline_shape = self.document.InlineShapes.Item(index)
                try:
                    inline_type = int(getattr(inline_shape, "Type", 0) or 0)
                    width = float(getattr(inline_shape, "Width", 0) or 0)
                    height = float(getattr(inline_shape, "Height", 0) or 0)
                    if inline_type in picture_types and width > 0 and height > 0:
                        count += 1
                except Exception:
                    continue
        except Exception:
            pass
        return count

    def _count_visible_floating_images(self) -> int:
        useful_picture_types = {11, 13}  # msoLinkedPicture, msoPicture
        count = 0
        try:
            for index in range(1, self.document.Shapes.Count + 1):
                shape = self.document.Shapes.Item(index)
                try:
                    shape_type = int(getattr(shape, "Type", 0) or 0)
                    visible = int(getattr(shape, "Visible", -1) or 0)
                    width = float(getattr(shape, "Width", 0) or 0)
                    height = float(getattr(shape, "Height", 0) or 0)
                    if shape_type in useful_picture_types and visible != 0 and width > 0 and height > 0:
                        count += 1
                except Exception:
                    continue
        except Exception:
            pass
        return count

    def _range_is_blank(self, page_range) -> bool:
        text = str(getattr(page_range, "Text", "") or "")
        ignored_chars = " \t\r\n\f\v\a\x07\u00a0"
        return text.strip(ignored_chars) == ""

    def _range_is_blank_or_placeholder_only(self, page_range) -> bool:
        text = str(getattr(page_range, "Text", "") or "")
        for placeholder_number in range(1, self.MAX_PLACEHOLDER_IMAGES + 1):
            text = text.replace(self._placeholder_text(placeholder_number), "")
        ignored_chars = " \t\r\n\f\v\a\x07\u00a0"
        return text.strip(ignored_chars) == ""

    def _delete_blank_page_range(self, page_range, page_number: int) -> bool:
        try:
            page_count = self.get_page_count()
            if page_count <= 1:
                page_range.Text = ""
            else:
                page_range = self._include_preceding_page_break(page_range, page_number)
                page_range.Delete()
            return True
        except Exception:
            if page_number != self.get_page_count():
                return False
            try:
                page_range.Text = ""
                return True
            except Exception:
                return False

    def _include_preceding_page_break(self, page_range, page_number: int):
        if page_number <= 1:
            return page_range
        try:
            start = int(page_range.Start)
            if start <= 0:
                return page_range
            previous_char = self.document.Range(start - 1, start).Text
            if previous_char == "\f":
                return self.document.Range(start - 1, int(page_range.End))
        except Exception:
            pass
        return page_range

    def _find_exact_text_range(self, text: str):
        for story_range in self._iter_story_ranges():
            search_range = story_range.Duplicate
            find = search_range.Find
            find.ClearFormatting()
            find.Text = text
            find.Forward = True
            find.Wrap = 0  # wdFindStop
            find.MatchCase = True
            find.MatchWholeWord = False
            find.MatchWildcards = False
            if find.Execute():
                return search_range.Duplicate
        return None

    def _replace_all_exact_text(self, text: str, replacement: str) -> None:
        for story_range in self._iter_story_ranges():
            search_range = story_range.Duplicate
            find = search_range.Find
            find.ClearFormatting()
            find.Replacement.ClearFormatting()
            find.Text = text
            find.Replacement.Text = replacement
            find.Forward = True
            find.Wrap = 0  # wdFindStop
            find.MatchCase = True
            find.MatchWholeWord = False
            find.MatchWildcards = False
            find.Execute(Replace=2)  # wdReplaceAll

    def _contains_exact_text(self, text: str) -> bool:
        for story_range in self._iter_story_ranges():
            try:
                if text in str(getattr(story_range, "Text", "") or ""):
                    return True
            except Exception:
                continue
        return False

    def _iter_story_ranges(self):
        for index in range(1, self.document.StoryRanges.Count + 1):
            current = self.document.StoryRanges.Item(index)
            while current is not None:
                yield current
                try:
                    current = current.NextStoryRange
                except Exception:
                    current = None

    def _range_page_number(self, word_range) -> int:
        try:
            return int(word_range.Information(3))
        except Exception:
            return 1

    def insert_next_label_image_file(self, image_path: str | Path) -> VisualValidationResult:
        position = self._current_slot_index
        if position >= len(self._slots):
            self.add_label_page_slot()

        result = self.insert_label_image_file(image_path, position)
        self._current_slot_index += 1
        return result

    def insert_label_image_file(self, image_path: str | Path, position: int) -> VisualValidationResult:
        if self.document is None:
            raise RuntimeError("Word no estÃƒÆ’Ã‚Â¡ abierto.")
        if position >= len(self._slots):
            self.add_label_page_slot()

        slot = self._slots[position]
        self._neutralize_slot_placeholder(slot)

        expected_width = WORD_LABEL_WIDTH_CM * self.POINTS_PER_CM
        expected_height = WORD_LABEL_HEIGHT_CM * self.POINTS_PER_CM
        left = slot["left"] + max((slot["width"] - expected_width) / 2, 0)
        top = slot["top"] + max((slot["height"] - expected_height) / 2, 0)

        try:
            picture = self.document.Shapes.AddPicture(
                FileName=str(Path(image_path).resolve()),
                LinkToFile=False,
                SaveWithDocument=True,
                Left=left,
                Top=top,
                Width=expected_width,
                Height=expected_height,
                Anchor=slot["anchor"],
            )
        except Exception as exc:
            raise RuntimeError(f"Word no pudo insertar la imagen real en el slot {slot['name']}: {exc}") from exc

        picture.WrapFormat.Type = 3
        try:
            picture.LockAspectRatio = 0
            picture.Width = expected_width
            picture.Height = expected_height
        except Exception:
            pass
        try:
            picture.Line.Visible = 0
        except Exception:
            pass

        try:
            picture.ZOrder(0)  # bring to front
        except Exception:
            pass

        return VisualValidationResult(
            position=position + 1,
            slot_name=slot["name"],
            expected_width_cm=WORD_LABEL_WIDTH_CM,
            expected_height_cm=WORD_LABEL_HEIGHT_CM,
            applied_width_cm=self._points_to_cm(expected_width),
            applied_height_cm=self._points_to_cm(expected_height),
            container_width_cm=self._points_to_cm(slot["width"]),
            container_height_cm=self._points_to_cm(slot["height"]),
            adjusted=True,
            details="imagen real embebida con AddPicture dentro de un slot controlado",
            shape_name=getattr(picture, "Name", ""),
            page_number=slot["page_number"],
            created_slot=bool(slot.get("created", False)),
        )

    def add_label_page_slot(self) -> dict:
        if self.document is None:
            raise RuntimeError("Word no estÃƒÆ’Ã‚Â¡ abierto.")

        geometry = self._slot_geometry_template()
        end_range = self.document.Range(self.document.Content.End - 1, self.document.Content.End - 1)
        end_range.Collapse(0)  # wdCollapseEnd
        if self._slots:
            end_range.InsertBreak(7)  # wdPageBreak
            end_range.Collapse(0)
        end_range.InsertParagraphAfter()
        end_range.Collapse(0)

        slot_number = len(self._slots) + 1
        shape = self.document.Shapes.AddShape(
            Type=1,
            Left=geometry["left"],
            Top=geometry["top"],
            Width=geometry["width"],
            Height=geometry["height"],
            Anchor=end_range,
        )
        shape.Name = f"ATL_LabelSlot_{slot_number:03d}"
        shape.WrapFormat.Type = 3
        try:
            shape.Fill.Visible = 0
            shape.Line.Visible = 0
        except Exception:
            pass

        slot = self._shape_to_slot(shape, slot_number, created=True)
        self._slots.append(slot)
        return slot

    def _discover_slots(self) -> list[dict]:
        slots: list[dict] = []
        for index in range(1, self.document.Shapes.Count + 1):
            shape = self.document.Shapes.Item(index)
            if not self._is_label_slot_candidate(shape):
                continue
            slots.append(self._shape_to_slot(shape, index, created=False))

        slots.sort(key=lambda item: (item["page_number"], item["top"], item["left"]))
        return self._one_slot_per_page(slots)

    def _one_slot_per_page(self, slots: list[dict]) -> list[dict]:
        page_numbers: set[int] = set()
        controlled_slots: list[dict] = []
        for slot in slots:
            page_number = slot["page_number"]
            if page_number in page_numbers:
                continue
            page_numbers.add(page_number)
            controlled_slots.append(slot)
        return controlled_slots

    def _shape_to_slot(self, shape, index: int, created: bool) -> dict:
        anchor = shape.Anchor
        return {
            "name": getattr(shape, "Name", f"Shape {index}"),
            "left": float(getattr(shape, "Left", 0)),
            "top": float(getattr(shape, "Top", 0)),
            "width": float(getattr(shape, "Width", 0)),
            "height": float(getattr(shape, "Height", 0)),
            "anchor": anchor,
            "shape_type": int(getattr(shape, "Type", 0) or 0),
            "page_number": self._anchor_page_number(anchor),
            "created": created,
        }

    def _neutralize_slot_placeholder(self, slot: dict) -> None:
        shape = self._find_shape_by_name(slot["name"])
        if shape is None:
            return

        try:
            shape.Fill.Visible = 0
        except Exception:
            pass

        try:
            shape.Line.Visible = 0
        except Exception:
            pass

        try:
            shape.TextFrame.TextRange.Text = ""
        except Exception:
            pass

        try:
            shape.TextFrame2.TextRange.Text = ""
        except Exception:
            pass

        try:
            if int(getattr(shape, "Type", 0) or 0) == 6:  # grupo
                shape.Delete()
                return
        except Exception:
            pass

        try:
            shape.Visible = 0
        except Exception:
            pass

    def _find_shape_by_name(self, shape_name: str):
        for index in range(1, self.document.Shapes.Count + 1):
            shape = self.document.Shapes.Item(index)
            if getattr(shape, "Name", "") == shape_name:
                return shape
        return None

    def _is_label_slot_candidate(self, shape) -> bool:
        name = str(getattr(shape, "Name", "") or "")
        if not name.startswith(self.TRUSTED_SLOT_PREFIX):
            return False

        width = float(getattr(shape, "Width", 0) or 0)
        height = float(getattr(shape, "Height", 0) or 0)
        if width <= 0 or height <= 0:
            return False

        expected_width = WORD_LABEL_WIDTH_CM * self.POINTS_PER_CM
        expected_height = WORD_LABEL_HEIGHT_CM * self.POINTS_PER_CM
        min_width = expected_width * (1 - self.SLOT_SIZE_TOLERANCE)
        max_width = expected_width * (1 + self.SLOT_SIZE_TOLERANCE)
        min_height = expected_height * (1 - self.SLOT_SIZE_TOLERANCE)
        max_height = expected_height * (1 + self.SLOT_SIZE_TOLERANCE)
        shape_type = int(getattr(shape, "Type", 0) or 0)
        return shape_type != 6 and min_width <= width <= max_width and min_height <= height <= max_height

    def _neutralize_untrusted_template_placeholders(self) -> None:
        for index in range(self.document.Shapes.Count, 0, -1):
            shape = self.document.Shapes.Item(index)
            if self._is_untrusted_template_placeholder(shape):
                self._neutralize_shape(shape)

    def _is_untrusted_template_placeholder(self, shape) -> bool:
        width = float(getattr(shape, "Width", 0) or 0)
        height = float(getattr(shape, "Height", 0) or 0)
        if width <= 0 or height <= 0:
            return False

        min_width = WORD_LABEL_WIDTH_CM * self.POINTS_PER_CM * 0.75
        min_height = WORD_LABEL_HEIGHT_CM * self.POINTS_PER_CM * 0.75
        shape_type = int(getattr(shape, "Type", 0) or 0)
        return shape_type == 6 or (width >= min_width and height >= min_height)

    def _neutralize_shape(self, shape) -> None:
        try:
            shape.Fill.Visible = 0
        except Exception:
            pass

        try:
            shape.Line.Visible = 0
        except Exception:
            pass

        try:
            shape.TextFrame.TextRange.Text = ""
        except Exception:
            pass

        try:
            shape.TextFrame2.TextRange.Text = ""
        except Exception:
            pass

        try:
            if int(getattr(shape, "Type", 0) or 0) == 6:  # grupo no confiable
                shape.Delete()
                return
        except Exception:
            pass

        try:
            shape.Visible = 0
        except Exception:
            pass

    def _slot_geometry_template(self) -> dict:
        if self._slots:
            first = self._slots[0]
            return {
                "left": first["left"],
                "top": first["top"],
                "width": first["width"],
                "height": first["height"],
            }

        return {
            "left": float(self.document.PageSetup.LeftMargin),
            "top": float(self.document.PageSetup.TopMargin),
            "width": WORD_LABEL_WIDTH_CM * self.POINTS_PER_CM,
            "height": WORD_LABEL_HEIGHT_CM * self.POINTS_PER_CM,
        }

    def _anchor_page_number(self, anchor) -> int:
        try:
            return int(anchor.Information(3))
        except Exception:
            return 1

    def print_document(self, printer_name: str) -> None:
        current_printer = self.word_app.ActivePrinter
        self.word_app.ActivePrinter = printer_name
        try:
            self.document.PrintOut(Background=False)
        finally:
            self.word_app.ActivePrinter = current_printer

    def print_label(self, printer_name: str, shape_name: str, page_number: int | None = None) -> None:
        if self.word_app is None or self.document is None:
            raise RuntimeError("Word no estÃƒÆ’Ã‚Â¡ abierto.")

        current_printer = self.word_app.ActivePrinter
        self.word_app.ActivePrinter = printer_name
        try:
            if shape_name:
                try:
                    self.document.Shapes.Item(shape_name).Select()
                except Exception:
                    pass
            if page_number is not None:
                try:
                    self.word_app.Selection.GoTo(What=1, Which=1, Count=page_number)
                except Exception:
                    pass
            self.document.PrintOut(Background=False, Range=2)
        finally:
            self.word_app.ActivePrinter = current_printer

    def save_document_copy(self, output_path: str) -> None:
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        self.document.SaveAs2(str(Path(output_path).resolve()))

    def show_to_user(self) -> None:
        if self.word_app is None or self.document is None:
            raise RuntimeError("Word no estÃƒÆ’Ã‚Â¡ abierto.")
        self.word_app.Visible = True
        self.word_app.DisplayAlerts = -1
        self.document.Activate()

    def release_to_user(self) -> None:
        self._released_to_user = True
        self.document = None
        self.word_app = None

    def export_pdf(self, output_path: str) -> None:
        wd_format_pdf = 17
        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        self.document.SaveAs2(str(Path(output_path).resolve()), FileFormat=wd_format_pdf)

    def _points_to_cm(self, value: float | None) -> float | None:
        if value is None:
            return None
        return round(float(value) / self.POINTS_PER_CM, 3)
