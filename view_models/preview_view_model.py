from collections.abc import Callable

from PySide6.QtCore import QObject, Signal
from PySide6.QtGui import QImage

from view_models.label_item_view_model import LabelItemViewModel


class PreviewViewModel(QObject):
    previewReady = Signal()
    printStarted = Signal()
    printCompleted = Signal()
    printFailed = Signal(str)
    redoRequested = Signal()

    def __init__(
        self,
        print_callback: Callable[[list[LabelItemViewModel]], None] | None = None,
        preview_image_callback: Callable[[LabelItemViewModel], QImage] | None = None,
        parent=None,
    ):
        super().__init__(parent)
        self._label_items: list[LabelItemViewModel] = []
        self._current_page_index = 0
        self._print_callback = print_callback
        self._preview_image_callback = preview_image_callback

    @property
    def label_items(self):
        return self._label_items

    @property
    def current_page_index(self):
        return self._current_page_index

    def set_items(self, items: list[LabelItemViewModel]):
        self._label_items = items
        self._current_page_index = 0
        self.previewReady.emit()

    def next_page(self):
        if self._current_page_index < len(self._label_items) - 1:
            self._current_page_index += 1

    def previous_page(self):
        if self._current_page_index > 0:
            self._current_page_index -= 1

    def preview_images(self) -> list[QImage]:
        if self._preview_image_callback is None:
            return []
        return [self._preview_image_callback(item) for item in self._label_items]

    def confirm(self):
        self.printStarted.emit()
        try:
            if self._print_callback is not None:
                self._print_callback(self._label_items)
        except Exception as exc:
            self.printFailed.emit(str(exc))
            return
        self.printCompleted.emit()

    def redo(self):
        self._label_items = []
        self._current_page_index = 0
        self.redoRequested.emit()
