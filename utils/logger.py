from __future__ import annotations

import faulthandler
import logging
import platform
import sys
import threading
from datetime import datetime
from logging import Logger
from pathlib import Path
from types import TracebackType
from typing import Callable

from utils.constants import APP_NAME, LOGS_DIR_NAME


UserMessageCallback = Callable[[str], None]


def get_application_base_dir() -> Path:
    """Return the executable/install folder when frozen, or the project root in development."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def get_logs_dir(base_dir: Path | None = None) -> Path:
    root_dir = Path(base_dir) if base_dir is not None else get_application_base_dir()
    logs_dir = root_dir / LOGS_DIR_NAME
    logs_dir.mkdir(parents=True, exist_ok=True)
    return logs_dir


def get_log_file_path(base_dir: Path | None = None) -> Path:
    return get_logs_dir(base_dir) / f"app_{datetime.now():%Y-%m-%d}.log"


def setup_logger(base_dir: Path | None = None) -> Logger:
    logs_dir = get_logs_dir(base_dir)
    log_file = logs_dir / f"app_{datetime.now():%Y-%m-%d}.log"

    logger = logging.getLogger(APP_NAME)
    logger.setLevel(logging.INFO)
    for handler in list(logger.handlers):
        logger.removeHandler(handler)
        handler.close()

    formatter = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] [%(name)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(formatter)

    try:
        faulthandler.enable(file=file_handler.stream, all_threads=True)
    except Exception:
        pass

    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(stream_handler)
    logger.propagate = False

    logging.captureWarnings(True)
    warnings_logger = logging.getLogger("py.warnings")
    warnings_logger.handlers.clear()
    warnings_logger.addHandler(file_handler)
    warnings_logger.setLevel(logging.WARNING)
    warnings_logger.propagate = False

    logger.info("Logger configurado. Archivo activo: %s", log_file)

    return logger


def setup_worker_logger(base_dir: Path | None = None) -> Logger:
    logs_dir = get_logs_dir(base_dir)
    log_file = logs_dir / f"worker_{datetime.now():%Y-%m-%d}.log"

    logger = logging.getLogger(f"{APP_NAME}.worker")
    logger.setLevel(logging.INFO)
    for handler in list(logger.handlers):
        logger.removeHandler(handler)
        handler.close()

    formatter = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] [%(name)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(formatter)

    try:
        faulthandler.enable(file=file_handler.stream, all_threads=True)
    except Exception:
        pass

    logger.addHandler(file_handler)
    logger.propagate = False
    logger.info("Logger del worker configurado. Archivo activo: %s", log_file)

    return logger


def install_global_exception_handlers(
    logger: Logger,
    user_message_callback: UserMessageCallback | None = None,
) -> None:
    """Install safe handlers for uncaught Python, thread and Qt messages."""

    def notify_user(message: str) -> None:
        if user_message_callback is None:
            return
        try:
            user_message_callback(message)
        except Exception:
            logger.exception("No se pudo mostrar el mensaje amigable de error al usuario.")

    def handle_exception(exc_type: type[BaseException], exc_value: BaseException, exc_traceback: TracebackType | None) -> None:
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        logger.critical("Excepción global no controlada.", exc_info=(exc_type, exc_value, exc_traceback))
        notify_user("Ocurrió un error inesperado. El detalle técnico fue guardado en la carpeta logs.")

    def handle_thread_exception(args: threading.ExceptHookArgs) -> None:
        logger.critical(
            "Excepción no controlada en hilo %s.",
            getattr(args.thread, "name", "desconocido"),
            exc_info=(args.exc_type, args.exc_value, args.exc_traceback),
        )
        notify_user("Ocurrió un error inesperado en segundo plano. El detalle técnico fue guardado en la carpeta logs.")

    sys.excepthook = handle_exception
    if hasattr(threading, "excepthook"):
        threading.excepthook = handle_thread_exception

    try:
        from PySide6.QtCore import QtMsgType, qInstallMessageHandler

        def qt_message_handler(mode: QtMsgType, context, message: str) -> None:
            context_text = ""
            if getattr(context, "file", None):
                context_text = f" ({context.file}:{context.line})"
            if mode == QtMsgType.QtDebugMsg:
                logger.debug("Qt: %s%s", message, context_text)
            elif mode == QtMsgType.QtInfoMsg:
                logger.info("Qt: %s%s", message, context_text)
            elif mode == QtMsgType.QtWarningMsg:
                logger.warning("Qt: %s%s", message, context_text)
            elif mode == QtMsgType.QtCriticalMsg:
                logger.error("Qt: %s%s", message, context_text)
            elif mode == QtMsgType.QtFatalMsg:
                logger.critical("Qt fatal: %s%s", message, context_text)
            else:
                logger.warning("Qt: %s%s", message, context_text)

        qInstallMessageHandler(qt_message_handler)
    except Exception:
        logger.exception("No se pudo instalar el handler de mensajes Qt.")


def log_runtime_context(logger: Logger) -> None:
    logger.info("Inicio de aplicación: %s", APP_NAME)
    logger.info("Ruta base de ejecución: %s", get_application_base_dir())
    logger.info("Directorio de logs: %s", get_logs_dir())
    logger.info("Sistema operativo: %s", platform.platform())
    logger.info("Python: %s", sys.version.replace("\n", " "))
