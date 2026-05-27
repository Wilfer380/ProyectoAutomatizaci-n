from collections.abc import Callable

from PySide6.QtCore import QObject, Signal


class PreviewViewModel(QObject):
    previewReady = Signal()
    printStarted = Signal()
    printCompleted = Signal()
    redoRequested = Signal()

    def __init__(self, print_callback: Callable[[list], None] | None = None, parent=None):
        super().__init__(parent)
        self._label_items = []
        self._current_page_index = 0
        self._print_callback = print_callback

    @property
    def label_items(self):
        return self._label_items

    @property
    def current_page_index(self):
        return self._current_page_index

    def set_items(self, items):
        self._label_items = items
        self._current_page_index = 0
        self.previewReady.emit()

    def next_page(self):
        if self._current_page_index < len(self._label_items) - 1:
            self._current_page_index += 1

    def previous_page(self):
        if self._current_page_index > 0:
            self._current_page_index -= 1

    def confirm(self):
        self.printStarted.emit()
        if self._print_callback is not None:
            self._print_callback(self._label_items)
        self.printCompleted.emit()

    def redo(self):
        self._label_items = []
        self._current_page_index = 0
        self.redoRequested.emit()
