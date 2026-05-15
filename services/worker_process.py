from __future__ import annotations

import argparse
import json
import sys
import traceback
from dataclasses import dataclass
from typing import Any

from models.app_settings import AppSettings
from services.excel_service import ExcelService
from services.print_service import PrintService
from services.process_service import ManualAdjustmentCancelled, ProcessService
from services.validation_service import ValidationError, ValidationService
from services.word_service import WordService
from utils.logger import log_runtime_context, setup_worker_logger


@dataclass(slots=True)
class WorkerResult:
    event: str
    payload: dict[str, Any]


def _emit(event: str, **payload: Any) -> None:
    message = {"type": event, **payload}
    data = (json.dumps(message, ensure_ascii=False) + "\n").encode("utf-8")
    sys.stdout.buffer.write(data)
    sys.stdout.buffer.flush()


def _configure_stdio() -> None:
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        if stream is None:
            continue
        try:
            stream.reconfigure(encoding="utf-8", errors="strict", newline="\n")
        except Exception:
            pass


def _read_command() -> dict[str, Any] | None:
    while True:
        line = sys.stdin.readline()
        if not line:
            return None
        line = line.strip()
        if not line:
            continue
        try:
            data = json.loads(line)
        except json.JSONDecodeError:
            continue
        if isinstance(data, dict):
            return data


def _build_process_service(logger) -> ProcessService:
    return ProcessService(
        excel_service=ExcelService(),
        word_service=WordService(),
        print_service=PrintService(),
        validation_service=ValidationService(),
        logger=logger,
    )


def _run_load_filters(logger, excel_path: str) -> int:
    process_service = _build_process_service(logger)
    _emit("status", message="Cargando filtros desde Excel...")
    try:
        filters = process_service.load_filters(excel_path)
        _emit("result", status="ok", filters=filters)
        return 0
    except Exception as error:
        logger.exception("Worker falló al cargar filtros.")
        _emit("error", message=str(error), details=traceback.format_exc())
        return 1


def _run_process(logger, settings: AppSettings, simulate: bool) -> int:
    process_service = _build_process_service(logger)

    def log_callback(message: str) -> None:
        _emit("log", message=message)

    def progress_callback(value: int) -> None:
        _emit("progress", value=int(value))

    def status_callback(message: str) -> None:
        _emit("status", message=message)

    def record_count_callback(count: int) -> None:
        _emit("record_count", value=int(count))

    def manual_adjust_callback(block_index: int, total_blocks: int, document_path: str, baseline_mtime_ns: int) -> str:
        _emit(
            "manual_review_requested",
            block_index=int(block_index),
            total_blocks=int(total_blocks),
            document_path=document_path,
            baseline_mtime_ns=int(baseline_mtime_ns),
        )
        while True:
            command = _read_command()
            if command is None:
                return "cancelar"
            if command.get("type") != "manual_review_response":
                continue
            response = str(command.get("response", "")).strip().lower()
            if response in {"continuar", "cancelar"}:
                return response

    try:
        process_service.run(
            excel_path=settings.excel_path,
            word_path=settings.word_template_path,
            selected_filter=settings.selected_filter,
            printer_name=settings.printer_name,
            simulate=simulate,
            log_callback=log_callback,
            progress_callback=progress_callback,
            status_callback=status_callback,
            record_count_callback=record_count_callback,
            manual_adjust_callback=manual_adjust_callback,
            working_directory=settings.working_directory,
            output_directory=settings.output_directory,
        )
        _emit("result", status="ok")
        return 0
    except ManualAdjustmentCancelled as error:
        logger.info("Worker cancelado por el usuario: %s", error)
        _emit("cancelled", message=str(error))
        return 0
    except ValidationError as error:
        logger.exception("Worker terminó con error de validación.")
        _emit("error", message=str(error), details=traceback.format_exc())
        return 1
    except Exception as error:
        logger.exception("Worker terminó con error inesperado.")
        _emit("error", message=str(error), details=traceback.format_exc())
        return 1


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(add_help=True)
    parser.add_argument("--worker", action="store_true")
    parser.add_argument("--worker-action", choices=("load-filters", "run-process"), default="")
    parser.add_argument("--excel", default="")
    parser.add_argument("--word", default="")
    parser.add_argument("--filter", dest="selected_filter", default="")
    parser.add_argument("--printer", default="")
    parser.add_argument("--simulate", action="store_true")
    parser.add_argument("--working-directory", default="")
    parser.add_argument("--output-directory", default="")
    return parser


def worker_main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args, _ = parser.parse_known_args(argv)
    _configure_stdio()
    logger = setup_worker_logger()
    log_runtime_context(logger)

    if args.worker_action == "load-filters":
        if not args.excel:
            _emit("error", message="Falta el archivo Excel para cargar filtros.")
            return 1
        return _run_load_filters(logger, args.excel)

    if args.worker_action == "run-process":
        settings = AppSettings(
            excel_path=args.excel,
            word_template_path=args.word,
            selected_filter=args.selected_filter,
            printer_name=args.printer,
            working_directory=args.working_directory,
            output_directory=args.output_directory,
        )
        return _run_process(logger, settings, args.simulate)

    _emit("error", message="Acción de worker no reconocida.")
    return 1
