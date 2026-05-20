from __future__ import annotations

from dataclasses import dataclass
import traceback
from logging import Logger
from pathlib import Path
from typing import Any, Callable, TypeVar


T = TypeVar("T")


@dataclass(slots=True)
class ErrorContext:
    step: str
    excel_path: str = ""
    word_path: str = ""
    temp_path: str = ""
    process_name: str = ""


class ErrorGuard:
    OFFICE_OFFLINE_HINT = (
        "No se pudo completar el proceso porque Microsoft Word o Excel presentó un problema. "
        "Verifica que Office esté instalado, activado y que no tenga ventanas pendientes de inicio de sesión. "
        "El detalle técnico fue guardado en la carpeta logs."
    )
    GENERIC_HINT = (
        "Hubo un problema al generar el documento. La aplicación no se cerrará. "
        "El detalle técnico fue guardado en la carpeta logs."
    )

    @staticmethod
    def friendly_message(exc: BaseException) -> str:
        text = f"{exc}".lower()
        office_markers = (
            "microsoft word",
            "microsoft excel",
            "word ha detectado un problema",
            "word no pudo desencadenar el evento",
            "inicia sesión",
            "sign in",
            "office",
            "0x800706ba",
            "0x8001010d",
            "0x80010108",
            "0x800a13e9",
            "com_error",
        )
        file_markers = (
            "permission denied",
            "acceso denegado",
            "file not found",
            "no such file",
            "could not find",
            "path not found",
        )

        if any(marker in text for marker in office_markers):
            return ErrorGuard.OFFICE_OFFLINE_HINT
        if any(marker in text for marker in file_markers):
            return "No se pudo abrir o crear un archivo requerido. Verifica la ruta, permisos y que no esté bloqueado. El detalle técnico fue guardado en la carpeta logs."
        return ErrorGuard.GENERIC_HINT

    @staticmethod
    def format_exception_details(exc: BaseException) -> str:
        return "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))


class SafeExecutor:
    def __init__(self, logger: Logger, user_message_callback: Callable[[str], None] | None = None) -> None:
        self.logger = logger
        self.user_message_callback = user_message_callback

    def execute(
        self,
        process_name: str,
        func: Callable[[], T],
        *,
        context: ErrorContext | None = None,
        user_message: str | None = None,
        on_error: Callable[[BaseException, str], None] | None = None,
        default: T | None = None,
    ) -> T | None:
        try:
            return func()
        except Exception as exc:
            self.logger.exception("Error controlado en %s.", process_name)
            if context is not None:
                self.logger.error(
                    "Contexto de error [%s]: excel=%s word=%s temp=%s",
                    context.step,
                    context.excel_path or "n/d",
                    context.word_path or "n/d",
                    context.temp_path or "n/d",
                )
            message = user_message or ErrorGuard.friendly_message(exc)
            if self.user_message_callback is not None:
                try:
                    self.user_message_callback(message)
                except Exception:
                    self.logger.exception("No se pudo mostrar el mensaje amigable al usuario.")
            if on_error is not None:
                try:
                    on_error(exc, message)
                except Exception:
                    self.logger.exception("No se pudo ejecutar el callback de error seguro.")
            return default
