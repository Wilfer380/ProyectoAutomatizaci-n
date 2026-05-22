from __future__ import annotations

import json
from typing import Any

from PySide6.QtCore import QObject, QProcess, Signal

from models.app_settings import AppSettings
from utils.constants import TARGET_PRINTER_NAME
from utils.runtime import build_worker_command


class WorkerClient(QObject):
    log_received = Signal(str)
    status_received = Signal(str)
    progress_received = Signal(int)
    record_count_received = Signal(int)
    filters_received = Signal(list)
    manual_review_requested = Signal(int, int, str, object)
    cancelled = Signal(str)
    finished = Signal()
    failed = Signal(str)

    def __init__(self, logger, parent: QObject | None = None) -> None:
        super().__init__(parent)
        self.logger = logger
        self._process: QProcess | None = None
        self._stdout_buffer = ""
        self._stderr_buffer = ""
        self._operation = ""
        self._completed = False
        self._failed = False

    def is_running(self) -> bool:
        return self._process is not None and self._process.state() != QProcess.NotRunning

    def start_load_filters(self, excel_path: str) -> None:
        self._start_worker("load-filters", ["--excel", excel_path])

    def start_process(self, settings: AppSettings, simulate: bool = False) -> None:
        args = [
            "--excel",
            settings.excel_path,
            "--word",
            settings.word_template_path,
            "--filter",
            settings.selected_filter,
            "--printer",
            settings.printer_name or TARGET_PRINTER_NAME,
            "--working-directory",
            settings.working_directory,
            "--output-directory",
            settings.output_directory,
        ]
        if simulate:
            args.append("--simulate")
        self._start_worker("run-process", args)

    def send_manual_review_response(self, response: str) -> None:
        if self._process is None or self._process.state() == QProcess.NotRunning:
            return
        payload = json.dumps({"type": "manual_review_response", "response": response}, ensure_ascii=False).encode("utf-8")
        self._process.write(payload + b"\n")

    def _start_worker(self, action: str, args: list[str]) -> None:
        if self.is_running():
            raise RuntimeError("El proceso secundario ya está en ejecución.")

        self._stdout_buffer = ""
        self._stderr_buffer = ""
        self._operation = action
        self._completed = False
        self._failed = False

        process = QProcess(self)
        process.setProcessChannelMode(QProcess.SeparateChannels)
        environment = process.processEnvironment()
        environment.insert("PYTHONUTF8", "1")
        environment.insert("PYTHONIOENCODING", "utf-8")
        process.setProcessEnvironment(environment)
        process.readyReadStandardOutput.connect(self._handle_stdout)
        process.readyReadStandardError.connect(self._handle_stderr)
        process.errorOccurred.connect(self._handle_error)
        process.finished.connect(self._handle_finished)
        self._process = process

        command = build_worker_command()
        program, program_args = command[0], command[1:]
        self.logger.info("Iniciando worker: accion=%s programa=%s args=%s", action, program, program_args + ["--worker", "--worker-action", action] + args)
        process.start(program, program_args + ["--worker", "--worker-action", action] + args)
        if not process.waitForStarted(5000):
            self._process = None
            raise RuntimeError("No se pudo iniciar el proceso secundario.")

    def _handle_stdout(self) -> None:
        if self._process is None:
            return
        chunk = bytes(self._process.readAllStandardOutput()).decode("utf-8", errors="replace")
        self._stdout_buffer += chunk
        while "\n" in self._stdout_buffer:
            line, self._stdout_buffer = self._stdout_buffer.split("\n", 1)
            self._process_stdout_line(line.strip())

    def _process_stdout_line(self, line: str) -> None:
        if not line:
            return
        try:
            payload = json.loads(line)
        except json.JSONDecodeError:
            self._protocol_error(f"Salida inválida del worker: {line}")
            return
        if not isinstance(payload, dict):
            self._protocol_error("Salida inválida del worker: formato no reconocido.")
            return

        event = str(payload.get("type", "")).strip()
        if event == "log":
            message = str(payload.get("message", ""))
            self.logger.info("[worker] %s", message)
            self.log_received.emit(message)
            return

        if event == "status":
            self.status_received.emit(str(payload.get("message", "")))
            return

        if event == "progress":
            self.progress_received.emit(int(payload.get("value", 0)))
            return

        if event == "record_count":
            self.record_count_received.emit(int(payload.get("value", 0)))
            return

        if event == "manual_review_requested":
            self.manual_review_requested.emit(
                int(payload.get("block_index", 0)),
                int(payload.get("total_blocks", 0)),
                str(payload.get("document_path", "")),
                int(payload.get("baseline_mtime_ns", 0)),
            )
            return

        if event == "result":
            if self._operation == "load-filters":
                filters = payload.get("filters", [])
                if not isinstance(filters, list):
                    self._protocol_error("El worker devolvió filtros con formato inválido.")
                    return
                self.filters_received.emit([str(item) for item in filters])
            self._completed = True
            self.finished.emit()
            return

        if event == "cancelled":
            self._completed = True
            self.cancelled.emit(str(payload.get("message", "Proceso cancelado por el usuario.")))
            return

        if event == "error":
            self._failed = True
            details = str(payload.get("details", "")).strip()
            message = str(payload.get("message", "Error desconocido del worker.")).strip()
            if details:
                self.logger.error("Worker error: %s\n%s", message, details)
            else:
                self.logger.error("Worker error: %s", message)
            self.failed.emit(message)
            return

        self._protocol_error(f"Evento de worker no reconocido: {event or '<vacío>'}")

    def _handle_stderr(self) -> None:
        if self._process is None:
            return
        output = bytes(self._process.readAllStandardError()).decode("utf-8", errors="replace").strip()
        if not output:
            return
        self._stderr_buffer += output + "\n"
        self.logger.warning("[worker stderr] %s", output)
        self.log_received.emit(output)

    def _handle_error(self, error: QProcess.ProcessError) -> None:
        if self._failed:
            return
        self._failed = True
        if error == QProcess.FailedToStart:
            message = "No se pudo iniciar el proceso secundario."
        elif error == QProcess.Crashed:
            message = "El proceso secundario terminó inesperadamente."
        else:
            message = "Ocurrió un error en el proceso secundario."
        if self._stderr_buffer.strip():
            self.logger.error("Worker process error: %s\nÚltimo stderr:\n%s", message, self._stderr_buffer.rstrip())
        else:
            self.logger.error("Worker process error: %s", message)
        self.failed.emit(message)

    def _handle_finished(self, exit_code: int, exit_status: QProcess.ExitStatus) -> None:
        if self._process is None:
            return

        if not self._completed and not self._failed:
            if exit_status == QProcess.CrashExit:
                message = f"El proceso secundario se cerró inesperadamente (código {exit_code})."
            elif exit_code != 0:
                message = f"El proceso secundario finalizó con código {exit_code}."
            else:
                message = "El proceso secundario devolvió una salida incompleta."
            self._failed = True
            if self._stderr_buffer.strip():
                self.logger.error("%s\nÚltimo stderr:\n%s", message, self._stderr_buffer.rstrip())
            else:
                self.logger.error(message)
            self.failed.emit(message)

        self._process.deleteLater()
        self._process = None
        self._stdout_buffer = ""
        self._stderr_buffer = ""
        self._operation = ""

    def _protocol_error(self, message: str) -> None:
        if self._failed:
            return
        self._failed = True
        self.logger.error(message)
        self.failed.emit("La respuesta del proceso secundario fue inválida.")
        if self._process is not None:
            self._process.kill()

    def shutdown(self) -> None:
        process = self._process
        if process is None:
            self.logger.info("Cierre del worker: no había proceso secundario activo.")
            return

        try:
            if process.state() != QProcess.NotRunning:
                self.logger.info("Cierre de aplicación: deteniendo worker propio (%s).", self._operation or "sin operación")
                process.terminate()
                if not process.waitForFinished(3000):
                    self.logger.warning("El worker no terminó a tiempo; forzando cierre del proceso propio.")
                    process.kill()
                    process.waitForFinished(3000)
            else:
                self.logger.info("Cierre del worker: el proceso secundario ya estaba detenido.")
        except Exception:
            self.logger.exception("Error al cerrar el proceso secundario propio durante el apagado de la app.")
