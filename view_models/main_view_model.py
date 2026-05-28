from models.asset_record import AssetRecord
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
    filtersLoaded = Signal(list)
    recordCountChanged = Signal(int)
    errorOccurred = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._selected_file_path = ""
        self._is_processing = False
        self._progress_value = 0
        self.worker = None
        self._pending_items = []
        self._records: list[AssetRecord] = []
        self._selected_filters: list[str] = []
        self._selected_records: list[AssetRecord] = []
        self._selected_records_by_filter: dict[str, list[AssetRecord]] = {}
        self._selection_configured = False

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
        self._clear_selection()
        self.fileSelected.emit(file_path)
        self._load_records_and_filters(file_path)

    @property
    def selected_filters(self):
        return self._selected_filters

    def set_selected_filters(self, filters: list[str]):
        self._selection_configured = True
        self._selected_filters = filters
        self._selected_records = []

    def set_selected_records(self, records: list[AssetRecord]):
        self._selection_configured = True
        self._selected_records = records

    def set_selected_records_by_filter(
        self, selected_records_by_filter: dict[str, list[AssetRecord]]
    ) -> None:
        self._selected_records_by_filter = selected_records_by_filter

    def records_for_filter(self, filter_name: str) -> list[AssetRecord]:
        return [record for record in self._records if record.section == filter_name]

    def records_by_filter(self) -> dict[str, list[AssetRecord]]:
        grouped: dict[str, list[AssetRecord]] = {}
        for record in self._records:
            grouped.setdefault(record.section, []).append(record)
        return grouped

    def selected_records_by_filter(self) -> dict[str, list[AssetRecord]]:
        return self._selected_records_by_filter

    def _clear_selection(self) -> None:
        self._selected_filters = []
        self._selected_records = []
        self._selected_records_by_filter = {}
        self._selection_configured = False

    def _load_records_and_filters(self, file_path: str) -> None:
        if not file_path:
            return
        try:
            self._records = ExcelService().extract_data(file_path)
            filters = sorted(
                {record.section for record in self._records if record.section}
            )
            self.filtersLoaded.emit(filters)
            self.recordCountChanged.emit(len(self._records))
        except Exception as exc:
            self._records = []
            self.filtersLoaded.emit([])
            self.recordCountChanged.emit(0)
            self.errorOccurred.emit(str(exc))

    def set_progress(self, value):
        self._progress_value = value
        self.progressChanged.emit(value)

    def process_file(self):
        if self._is_processing:
            return

        if self._records:
            if self._selection_configured and self._selected_records:
                self._emit_records_for_preview(self._selected_records)
                return
            self.errorOccurred.emit(
                "Seleccioná al menos un filtro o una etiqueta antes de generar."
            )
            return

        if not self._selected_file_path:
            return

        self._is_processing = True
        self.set_progress(0)
        self.processingStarted.emit()

        self.worker = ExtractionWorker(self._selected_file_path)
        self.worker.setParent(self)
        self.worker.dataReady.connect(self._on_worker_data_ready)
        self.worker.error.connect(self._on_worker_error)
        self.worker.finished.connect(self._on_worker_finished)
        self.worker.finished.connect(self.worker.deleteLater)

        # Simulate some progress before we know it's done
        self.set_progress(50)

        self.worker.start()

    def _records_to_items(self, records: list[AssetRecord]) -> list[LabelItemViewModel]:
        items = []
        for record in records:
            vm = LabelItemViewModel()
            vm.asset_id = record.asset_id
            vm.asset_name = record.asset_name
            vm.section = record.section
            vm.image_data = record.image
            items.append(vm)
        return items

    def _emit_records_for_preview(self, records: list[AssetRecord]) -> None:
        self.processingStarted.emit()
        self.set_progress(100)
        self.processingFinished.emit(self._records_to_items(records))

    def _on_worker_data_ready(self, records):
        self._records = records
        self.recordCountChanged.emit(len(records))
        self._pending_items = self._records_to_items(records)

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
