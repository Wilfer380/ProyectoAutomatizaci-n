from __future__ import annotations

import os
from threading import Event
from pathlib import Path
from typing import Iterable

from PySide6.QtCore import QObject, QThread, QTimer, Signal

from models.app_settings import AppSettings
from services.excel_service import ExcelService
from services.print_service import PrintService
from services.process_service import ManualAdjustmentCancelled, ProcessService
from services.worker_client import WorkerClient
from services.validation_service import ValidationError, ValidationService
from services.word_service import WordService
from ui.main_window import MainWindow
from utils.error_guard import ErrorGuard, SafeExecutor
from utils.config import ConfigManager
from utils.constants import PROCESS_STATUS_COMPLETED, PROCESS_STATUS_ERROR, PROCESS_STATUS_RUNNING


class ProcessWorker(QObject):
    finished = Signal()
    cancelled = Signal(str)
    failed = Signal(str)
    log_emitted = Signal(str)
    progress_changed = Signal(int)
    status_changed = Signal(str)
    record_count_changed = Signal(int)
    manual_adjust_requested = Signal(int, int, str, object)

    def __init__(self, process_service: ProcessService, settings: AppSettings, logger, simulate: bool = False) -> None:
        super().__init__()
        self.process_service = process_service
        self.settings = settings
        self.logger = logger
        self.simulate = simulate
        self._manual_adjust_event = Event()
        self._manual_adjust_response = "cancelar"

    def run(self) -> None:
        try:
            self.process_service.run(
                excel_path=self.settings.excel_path,
                word_path=self.settings.word_template_path,
                selected_filter=self.settings.selected_filter,
                printer_name=self.settings.printer_name,
                simulate=self.simulate,
                working_directory=self.settings.working_directory,
                log_callback=self.log_emitted.emit,
                progress_callback=self.progress_changed.emit,
                status_callback=self.status_changed.emit,
                record_count_callback=self.record_count_changed.emit,
                manual_adjust_callback=self.pause_for_manual_adjust,
                output_directory=self.settings.output_directory,
            )
            self.finished.emit()
        except ManualAdjustmentCancelled as cancelled:
            self.logger.info("Proceso cancelado por revisión manual: %s", cancelled)
            self.cancelled.emit(str(cancelled))
        except Exception as error:
            self.logger.exception("Error no controlado en ProcessWorker.")
            self.failed.emit(str(error))

    def pause_for_manual_adjust(self, block_index: int, total_blocks: int, document_path: str, baseline_mtime_ns: int) -> str:
        self._manual_adjust_event.clear()
        self._manual_adjust_response = "cancelar"
        self.manual_adjust_requested.emit(block_index, total_blocks, document_path, baseline_mtime_ns)
        self._manual_adjust_event.wait()
        return self._manual_adjust_response

    def set_manual_adjust_response(self, response: str) -> None:
        self._manual_adjust_response = response
        self._manual_adjust_event.set()


class FilterLoadWorker(QObject):
    finished = Signal(list)
    failed = Signal(str)

    def __init__(self, process_service: ProcessService, excel_path: str, logger) -> None:
        super().__init__()
        self.process_service = process_service
        self.excel_path = excel_path
        self.logger = logger

    def run(self) -> None:
        try:
            filters = self.process_service.load_filters(self.excel_path)
            self.finished.emit(filters)
        except Exception as error:
            self.logger.exception("Error no controlado en FilterLoadWorker.")
            self.failed.emit(str(error))


class MainController:
    def __init__(self, window: MainWindow, config_manager: ConfigManager, logger) -> None:
        self.window = window
        self.config_manager = config_manager
        self.logger = logger
        self.settings = self.config_manager.load()

        self.validation_service = ValidationService()
        self.print_service = PrintService()
        self.process_service = ProcessService(
            excel_service=ExcelService(),
            word_service=WordService(),
            print_service=self.print_service,
            validation_service=self.validation_service,
            logger=logger,
        )

        self.worker_client = WorkerClient(logger, self.window)
        self.safe_executor = SafeExecutor(logger, lambda message: self.window.show_error("Proceso", message))

        self.worker_thread: QThread | None = None
        self.worker: ProcessWorker | None = None
        self.filter_thread: QThread | None = None
        self.filter_worker: FilterLoadWorker | None = None
        self._active_worker_operation = ""
        self._paused_waiting_for_block_review = False
        self._current_block_index: int | None = None
        self._current_total_blocks: int | None = None
        self._current_block_file = ""
        self._current_block_baseline_mtime_ns = 0

        self._bind_events()
        self._bind_worker_events()
        self._load_settings_into_view()
        self.logger.info("Controlador principal inicializado.")

    def cleanup_on_exit(self) -> None:
        try:
            self.logger.info("Cierre de aplicación solicitado.")
            if self.worker_client.is_running():
                self.logger.info("Hay un proceso secundario activo al cerrar la aplicación.")
            self.worker_client.shutdown()

            excel_service = self.process_service.excel_service
            if excel_service.excel_app is not None or excel_service.workbook is not None:
                self.logger.info("Se detectó una instancia propia de Excel abierta al cerrar la aplicación.")
            else:
                self.logger.info("No había una instancia propia de Excel abierta al cerrar la aplicación.")
            self.process_service.excel_service.close(save_changes=False)
            self.process_service.word_service.close(save_changes=False)
            self.logger.info("Limpieza de recursos al cerrar la aplicación completada.")
        except Exception:
            self.logger.exception("Error durante la limpieza de cierre de la aplicación.")

    def _safe_ui_action(self, action_name: str, callback, *, user_message: str = "Se produjo un error inesperado.") -> None:
        try:
            callback()
        except ValidationError as error:
            self.logger.warning("%s falló por validación: %s", action_name, error)
            self.window.show_error("Validación", str(error))
            self.window.append_log(f"Error: {error}")
        except Exception as error:
            self.logger.exception("Error inesperado en %s.", action_name)
            self.safe_executor.execute(
                action_name,
                lambda: (_ for _ in ()).throw(error),
                user_message=user_message or ErrorGuard.friendly_message(error),
            )
            self.window.append_log(f"Error inesperado: {error}")

    def _bind_events(self) -> None:
        self.window.select_excel_requested.connect(self.on_select_excel)
        self.window.select_word_requested.connect(self.on_select_word)
        self.window.excel_path_changed.connect(self.on_excel_path_changed)
        self.window.word_path_changed.connect(self.on_word_path_changed)
        self.window.refresh_filters_requested.connect(self.on_refresh_filters)
        self.window.start_process_requested.connect(self.on_start_process)
        self.window.start_simulation_requested.connect(self.on_start_simulation)
        self.window.cancel_process_requested.connect(self.on_cancel_process)
        self.window.printer_config_requested.connect(self.on_printer_config)
        self.window.filter_combo.currentTextChanged.connect(self.on_filter_changed)

    def _bind_worker_events(self) -> None:
        self.worker_client.log_received.connect(self._append_process_log)
        self.worker_client.status_received.connect(self.window.set_status)
        self.worker_client.progress_received.connect(self.window.set_progress)
        self.worker_client.record_count_received.connect(self.window.set_record_count)
        self.worker_client.filters_received.connect(self._on_filters_loaded_from_worker)
        self.worker_client.manual_review_requested.connect(self._on_manual_adjust_requested)
        self.worker_client.cancelled.connect(self._on_worker_cancelled)
        self.worker_client.finished.connect(self._on_worker_finished)
        self.worker_client.failed.connect(self._on_worker_failed)

    def _load_settings_into_view(self) -> None:
        self.settings.excel_path = ""
        self.settings.word_template_path = ""
        self.window.set_excel_path("")
        self.window.set_word_path("")
        self.window.set_printer_name(self.settings.printer_name)
        self.window.set_status("Listo para iniciar.")

    def on_select_excel(self) -> None:
        self._safe_ui_action("on_select_excel", self._on_select_excel)

    def _on_select_excel(self) -> None:
        path = self.window.choose_excel_file(self.settings.working_directory)
        if not path:
            return

        self.settings.excel_path = path
        self.settings.working_directory = str(Path(path).parent)
        self.window.set_excel_path(path)
        self.window.set_start_enabled(False)
        self._restore_cached_filters_for_excel(path)
        self._save_settings()
        self.window.append_log(f"Archivo Excel seleccionado: {path}")
        self.logger.info("Archivo Excel seleccionado: %s", path)
        self.on_refresh_filters()

    def on_select_word(self) -> None:
        self._safe_ui_action("on_select_word", self._on_select_word)

    def _on_select_word(self) -> None:
        path = self.window.choose_word_file(self.settings.working_directory)
        if not path:
            return

        self.settings.word_template_path = path
        self.window.set_word_path(path)
        self._save_settings()
        self.window.append_log(f"Plantilla Word seleccionada: {path}")
        self.logger.info("Plantilla Word seleccionada: %s", path)

    def on_excel_path_changed(self, path: str) -> None:
        self._safe_ui_action("on_excel_path_changed", lambda: self._on_excel_path_changed(path))

    def _on_excel_path_changed(self, path: str) -> None:
        if not path:
            self.settings.excel_path = ""
            self.window.set_excel_path("")
            self._save_settings()
            return

        self.settings.excel_path = path
        parent = Path(path).parent
        if str(parent):
            self.settings.working_directory = str(parent)
        self.window.set_excel_path(path)
        self.window.set_start_enabled(False)
        self._restore_cached_filters_for_excel(path)
        self._save_settings()
        self.window.append_log(f"Ruta Excel actualizada manualmente: {path}")
        self.logger.info("Ruta Excel actualizada manualmente: %s", path)

    def on_word_path_changed(self, path: str) -> None:
        self._safe_ui_action("on_word_path_changed", lambda: self._on_word_path_changed(path))

    def _on_word_path_changed(self, path: str) -> None:
        if not path:
            self.settings.word_template_path = ""
            self.window.set_word_path("")
            self._save_settings()
            return

        self.settings.word_template_path = path
        self.window.set_word_path(path)
        self._save_settings()
        self.window.append_log(f"Ruta Word actualizada manualmente: {path}")
        self.logger.info("Ruta Word actualizada manualmente: %s", path)

    def on_refresh_filters(self) -> None:
        self._safe_ui_action("on_refresh_filters", self._on_refresh_filters)

    def _on_refresh_filters(self) -> None:
        if self.worker_client.is_running():
            self.window.append_log("La carga de filtros ya está en curso.")
            self.logger.info("Solicitud de carga de filtros ignorada: ya hay una carga en curso.")
            return

        self.logger.info("Validando Excel antes de cargar filtros: %s", self.settings.excel_path)
        self.validation_service.validate_excel_file(self.settings.excel_path)

        self.window.set_status("Cargando filtros desde Excel...")
        self.window.refresh_filters_button.setEnabled(False)
        self._active_worker_operation = "load-filters"
        self.worker_client.start_load_filters(self.settings.excel_path)

    def _on_filters_loaded(self, filters: list[str]) -> None:
        try:
            self._cache_filters_for_current_excel(filters)
            self.window.set_filters(filters)
            if self.settings.selected_filter:
                self.window.set_selected_filter(self.settings.selected_filter)
            if not self.window.selected_filter() and filters:
                self.window.set_selected_filter(filters[0])
            self._save_settings()
            self.window.refresh_filters_button.setEnabled(True)
            self.window.filter_combo.setEnabled(True)

            if not filters:
                self.window.set_start_enabled(False)
                self.window.set_status("No se encontraron filtros válidos.")
                self.window.append_log("No se encontraron filtros válidos en el Excel.")
                self.logger.warning("No se encontraron filtros válidos en el Excel.")
                self.window.show_error("Filtros", "No se encontraron filtros válidos en el Excel.")
                return

            self.window.append_log(f"Filtros cargados: {len(filters)}")
            self.logger.info("Filtros cargados correctamente. Cantidad: %s", len(filters))
            self.window.set_status("Filtros cargados correctamente.")
            self.window.set_start_enabled(True)
        except Exception:
            self.logger.exception("Error inesperado al aplicar filtros cargados en la UI.")
            self.window.show_error("Filtros", "No se pudieron aplicar los filtros cargados.")

    def _on_filters_loaded_from_worker(self, filters: list[str]) -> None:
        self._on_filters_loaded(filters)

    def _on_filters_failed(self, message: str) -> None:
        try:
            self.logger.error("Error al cargar filtros: %s", message)
            self.window.show_error("Validación", message)
            self.window.append_log(f"Error: {message}")
            self.window.set_status("Error al cargar filtros.")
            self.window.filter_combo.setEnabled(True)
            self.window.refresh_filters_button.setEnabled(True)
            self.window.set_start_enabled(False)
        except Exception:
            self.logger.exception("Error inesperado al manejar la falla de filtros.")

    def on_filter_changed(self, value: str) -> None:
        self._safe_ui_action("on_filter_changed", lambda: self._on_filter_changed(value))

    def _on_filter_changed(self, value: str) -> None:
        self.settings.selected_filter = value.strip()
        self._save_settings()
        if self.settings.selected_filter:
            self.logger.info("Filtro seleccionado: %s", self.settings.selected_filter)

    def on_printer_config(self) -> None:
        self._safe_ui_action("on_printer_config", self._on_printer_config, user_message="No se pudo validar la impresora.")

    def _on_printer_config(self) -> None:
        printers = self.print_service.get_installed_printers()
        if not printers:
            self.logger.error("No se pudieron detectar impresoras instaladas.")
            self.window.show_error("Impresora", "No se pudieron detectar impresoras instaladas.")
            return

        if self.settings.printer_name not in printers:
            self.logger.error("Impresora requerida no instalada: %s", self.settings.printer_name)
            self.window.show_error("Impresora", f"La impresora '{self.settings.printer_name}' no está instalada.")
            return

        self.window.show_info("Impresora", f"Impresora detectada correctamente: {self.settings.printer_name}")
        self.window.append_log(f"Impresora detectada: {self.settings.printer_name}")
        self.logger.info("Impresora detectada correctamente: %s", self.settings.printer_name)

    def on_start_process(self) -> None:
        self._safe_ui_action("on_start_process", self._on_start_process, user_message="No se pudo iniciar el proceso.")

    def _on_start_process(self) -> None:
        if self._paused_waiting_for_block_review:
            self._continue_after_manual_review()
            return
        self._start_process(simulate=False)

    def on_start_simulation(self) -> None:
        self._safe_ui_action("on_start_simulation", lambda: self._start_process(simulate=True), user_message="No se pudo iniciar la simulación.")

    def on_cancel_process(self) -> None:
        self._safe_ui_action("on_cancel_process", self._on_cancel_process, user_message="No se pudo cancelar el proceso.")

    def _on_cancel_process(self) -> None:
        if not self._paused_waiting_for_block_review:
            return

        block = self._current_block_index or 0
        total = self._current_total_blocks or 0
        self.window.append_log(f"Bloque {block}/{total}: cancelación solicitada por el usuario.")
        self.logger.info("Bloque %s/%s: cancelación solicitada por el usuario.", block, total)
        self.window.set_status("Cancelando proceso...")
        self._leave_manual_review_mode()
        self.window.set_busy(True)
        self.worker_client.send_manual_review_response("cancelar")

    def _start_process(self, simulate: bool) -> None:
        if self.worker_client.is_running():
            self.window.append_log("El proceso ya está en curso.")
            self.logger.info("Solicitud de inicio ignorada: proceso ya en curso.")
            return

        try:
            self.logger.info("Validando datos antes de iniciar proceso. Simulación=%s", simulate)
            self.validation_service.validate_excel_file(self.settings.excel_path)
            self.validation_service.validate_word_file(self.settings.word_template_path)
            self.settings.selected_filter = self.window.selected_filter()
            self.validation_service.validate_selected_filter(self.settings.selected_filter)
            if not simulate:
                self.validation_service.validate_printer_installed(self.settings.printer_name)
            self._save_settings()
        except ValidationError as error:
            self.logger.warning("Validación fallida antes de iniciar proceso: %s", error)
            self.window.show_error("Validación", str(error))
            self.window.append_log(f"Error: {error}")
            return
        except Exception as error:
            self.logger.exception("Error inesperado durante la validación previa al proceso.")
            self.window.show_error("Proceso", "No se pudo iniciar el proceso por un error inesperado.")
            self.window.append_log(f"Error: {error}")
            return

        self.logger.info(
            "Inicio de %s. Excel=%s Word=%s Filtro=%s",
            "simulación" if simulate else "generación por bloques",
            self.settings.excel_path,
            self.settings.word_template_path,
            self.settings.selected_filter,
        )

        self.window.clear_logs()
        self.window.set_progress(0)
        self.window.set_busy(True)
        self.window.set_status("Simulación en curso..." if simulate else PROCESS_STATUS_RUNNING)
        self._active_worker_operation = "run-process"
        self.worker_client.start_process(self.settings, simulate=simulate)

    def _on_manual_adjust_requested(self, block_index: int, total_blocks: int, document_path: str, baseline_mtime_ns: object) -> None:
        try:
            self._paused_waiting_for_block_review = True
            self._current_block_index = block_index
            self._current_total_blocks = total_blocks
            self._current_block_file = document_path
            self._current_block_baseline_mtime_ns = int(baseline_mtime_ns)
            self.window.set_manual_review_mode(True, block_index, total_blocks, document_path)
            if document_path:
                QTimer.singleShot(300, lambda path=document_path: self._open_manual_review_document(path))
            self.logger.info("Bloque %s/%s listo para revisión manual: %s", block_index, total_blocks, document_path)
        except Exception:
            self.logger.exception("Error inesperado al activar revisión manual del bloque.")
            self.window.show_error("Proceso", "No se pudo abrir la revisión manual del bloque.")

    def _open_manual_review_document(self, document_path: str) -> None:
        try:
            review_service = WordService()
            review_service.open(document_path, visible=True)
            review_service.show_to_user()
            review_service.release_to_user()
            self.logger.info("Documento Word abierto con COM para revisión manual: %s", document_path)
        except Exception:
            self.logger.exception("No se pudo abrir el documento Word para revisión manual: %s", document_path)
            try:
                os.startfile(document_path)
            except Exception:
                self.logger.exception("Tampoco se pudo abrir el documento con la asociación del sistema: %s", document_path)

    def _continue_after_manual_review(self) -> None:
        block = self._current_block_index or 0
        total = self._current_total_blocks or 0
        document_path = self._current_block_file

        if not self.window.file_was_saved_after(document_path, self._current_block_baseline_mtime_ns):
            self.window.append_log(f"Bloque {block}/{total}: no se detectó guardado/modificación posterior.")
            self.logger.warning("Bloque %s/%s: no se detectó guardado posterior en %s", block, total, document_path)
            if not self.window.confirm_continue_without_save(document_path):
                self.window.set_status(
                    f"Bloque {block}/{total} listo. Guardá en Word o presioná Continuar para forzar."
                )
                return
            self.window.append_log(f"Bloque {block}/{total}: continuación forzada sin guardado detectado.")
            self.logger.warning("Bloque %s/%s: continuación forzada sin guardado detectado.", block, total)

        self.window.append_log(f"Bloque {block}/{total}: continuación al siguiente bloque.")
        self.logger.info("Bloque %s/%s: continuación al siguiente bloque.", block, total)
        self.window.set_status("Continuando con el siguiente bloque...")
        self._leave_manual_review_mode()
        self.window.set_busy(True)
        self.worker_client.send_manual_review_response("continuar")

    def _leave_manual_review_mode(self, update_window: bool = True) -> None:
        self._paused_waiting_for_block_review = False
        if update_window:
            self.window.set_manual_review_mode(False)
        self._current_block_index = None
        self._current_total_blocks = None
        self._current_block_file = ""
        self._current_block_baseline_mtime_ns = 0

    def _on_worker_finished(self) -> None:
        if self._active_worker_operation == "load-filters":
            self._active_worker_operation = ""
            self.window.filter_combo.setEnabled(True)
            self.window.refresh_filters_button.setEnabled(True)
            self.window.set_status("Filtros cargados correctamente.")
            return

        try:
            self._active_worker_operation = ""
            QTimer.singleShot(0, self._reset_window_after_process_finished)
        except Exception:
            self.logger.exception("No se pudo programar el cierre exitoso del proceso.")

    def _on_worker_cancelled(self, message: str) -> None:
        try:
            self._active_worker_operation = ""
            QTimer.singleShot(0, lambda: self._reset_window_after_process_cancelled(message))
        except Exception:
            self.logger.exception("No se pudo programar el cierre por cancelación.")

    def _reset_window_after_process_finished(self) -> None:
        try:
            self._leave_manual_review_mode(update_window=False)
            self.window.reset_after_process_finished("Listo para una nueva impresión.")
            self.logger.info("Proceso finalizado correctamente y UI restablecida.")
        except RuntimeError as error:
            self.logger.warning("Reset final omitido porque la ventana ya no está disponible: %s", error)
        except Exception:
            self.logger.exception("Error no crítico durante el reset final de la UI.")

    def _reset_window_after_process_cancelled(self, message: str) -> None:
        try:
            self._leave_manual_review_mode(update_window=False)
            self.window.reset_after_process_finished("Proceso cancelado. Listo para una nueva impresión.")
            self.window.append_log(message)
            self.logger.info("Proceso cancelado: %s", message)
        except RuntimeError as error:
            self.logger.warning("Reset de cancelación omitido porque la ventana ya no está disponible: %s", error)
        except Exception:
            self.logger.exception("Error no crítico durante el reset de cancelación de la UI.")

    def _on_worker_failed(self, message: str) -> None:
        try:
            active_operation = self._active_worker_operation
            self._active_worker_operation = ""
            self._leave_manual_review_mode()
            if active_operation != "load-filters":
                self.window.set_busy(False)
            self.window.set_status(PROCESS_STATUS_ERROR)
            self.window.append_log(f"Error: {message}")
            self.logger.error("Proceso finalizado con error: %s", message)
            title = "Filtros" if active_operation == "load-filters" else "Proceso"
            self.window.show_error(title, message)
            self.window.refresh_filters_button.setEnabled(True)
            self.window.filter_combo.setEnabled(True)
        except Exception:
            self.logger.exception("Error inesperado al mostrar la falla del proceso.")

    def _append_process_log(self, message: str) -> None:
        QTimer.singleShot(0, lambda msg=message: self._append_process_log_ui(msg))

    def _append_process_log_ui(self, message: str) -> None:
        try:
            self.logger.info(message)
            self.window.append_log(message)
        except Exception:
            self.logger.exception("Error inesperado al agregar log a la UI.")

    def _save_settings(self) -> None:
        self.config_manager.save(self.settings)

    def _excel_cache_key(self, excel_path: str) -> str:
        path = Path(excel_path)
        try:
            stat = path.stat()
            return f"{path.resolve()}|{stat.st_mtime_ns}|{stat.st_size}"
        except OSError:
            return str(path.resolve())

    def _restore_cached_filters_for_excel(self, excel_path: str) -> None:
        cached = self.settings.filter_cache.get(self._excel_cache_key(excel_path), [])
        self.window.set_filters(cached)
        if self.settings.selected_filter:
            self.window.set_selected_filter(self.settings.selected_filter)
        if not self.window.selected_filter() and cached:
            self.window.set_selected_filter(cached[0])

    def _cache_filters_for_current_excel(self, filters: Iterable[str]) -> None:
        excel_path = self.settings.excel_path.strip()
        if not excel_path:
            return

        self.settings.filter_cache[self._excel_cache_key(excel_path)] = [str(item) for item in filters]
