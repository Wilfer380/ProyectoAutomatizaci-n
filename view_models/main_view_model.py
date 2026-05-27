from PySide6.QtCore import QObject, Signal, QThread
from services.excel_service import ExcelService
from view_models.label_item_view_model import LabelItemViewModel


class ExtractionWorker(QThread):
    dataReady = Signal(list)
    error = Signal(str)

    def __init__(self, file_path):
        super().__init__()
        self.file_path = file_path

    def run(self):
        try:
            service = ExcelService()
            records = service.extract_data(self.file_path)
            self.dataReady.emit(records)
        except Exception as e:
            self.error.emit(str(e))


class MainViewModel(QObject):
    fileSelected = Signal(str)
    processingStarted = Signal()
    processingFinished = Signal(list)
    progressChanged = Signal(int)
    errorOccurred = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._selected_file_path = ""
        self._is_processing = False
        self._progress_value = 0
        self.worker = None
        self._pending_items = []

    @property
    def selected_file_path(self):
        return self._selected_file_path

    @property
    def is_processing(self):
        return self._is_processing

    @property
    def progress_value(self):
        return self._progress_value

    def select_file(self, file_path):
        self._selected_file_path = file_path
        self.fileSelected.emit(file_path)

    def set_progress(self, value):
        self._progress_value = value
        self.progressChanged.emit(value)

    def process_file(self):
        if not self._selected_file_path or self._is_processing:
            return

        self._is_processing = True
        self.set_progress(0)
        self.processingStarted.emit()

        # Start worker
        self.worker = ExtractionWorker(self._selected_file_path)
        self.worker.setParent(self)
        self.worker.dataReady.connect(self._on_worker_data_ready)
        self.worker.error.connect(self._on_worker_error)
        self.worker.finished.connect(self._on_worker_finished)
        self.worker.finished.connect(self.worker.deleteLater)

        # Simulate some progress before we know it's done
        self.set_progress(50)

        self.worker.start()

    def _on_worker_data_ready(self, records):
        items = []
        for record in records:
            vm = LabelItemViewModel()
            vm.asset_id = record.asset_id
            vm.asset_name = record.asset_name
            vm.section = record.section
            vm.image_data = record.image
            items.append(vm)

        self._pending_items = items

    def _on_worker_finished(self):
        self._is_processing = False
        if self._pending_items:
            self.set_progress(100)
            self.processingFinished.emit(self._pending_items)
        self._pending_items = []
        self.worker = None

    def _on_worker_error(self, err_msg):
        self._pending_items = []
        self._is_processing = False
        self.errorOccurred.emit(err_msg)
