from __future__ import annotations

from collections.abc import Callable, Iterable
from dataclasses import dataclass

from PySide6.QtPrintSupport import QPrinterInfo

from utils.constants import TARGET_PRINTER_NAME


class PrinterDriverMissingError(RuntimeError):
    def __init__(self, printer_name: str = TARGET_PRINTER_NAME) -> None:
        self.printer_name = printer_name
        super().__init__(missing_driver_message(printer_name))


def missing_driver_message(printer_name: str = TARGET_PRINTER_NAME) -> str:
    return (
        f"Controlador {printer_name} no detectado.\n\n"
        "Instalá el controlador/Printer Utility de SATO WS4 en este computador "
        "y verificá que Windows muestre la impresora con ese nombre exacto.\n\n"
        "Configuración esperada para producción: cola 'SATO WS408', etiqueta "
        "48 mm x 23 mm, resolución 203 DPI, sin escalado automático y sensor "
        "de etiquetas calibrado.\n\n"
        "Si no tenés permisos o el instalador no funciona, contactá a TI o al "
        "departamento de informática para que instalen y configuren la impresora."
    )


@dataclass(frozen=True)
class PrinterDriverStatus:
    printer_name: str
    installed: bool
    available_printers: tuple[str, ...]
    message: str = ""


def available_printer_names() -> tuple[str, ...]:
    return tuple(QPrinterInfo.availablePrinterNames())


def check_printer_driver(
    printer_name: str = TARGET_PRINTER_NAME,
    *,
    printer_names_provider: Callable[[], Iterable[str]] | None = None,
) -> PrinterDriverStatus:
    provider = printer_names_provider or available_printer_names
    available = tuple(str(name) for name in provider())
    installed = printer_name in available
    return PrinterDriverStatus(
        printer_name=printer_name,
        installed=installed,
        available_printers=available,
        message="" if installed else missing_driver_message(printer_name),
    )


def ensure_printer_driver(
    printer_name: str = TARGET_PRINTER_NAME,
    *,
    printer_names_provider: Callable[[], Iterable[str]] | None = None,
) -> None:
    status = check_printer_driver(
        printer_name,
        printer_names_provider=printer_names_provider,
    )
    if not status.installed:
        raise PrinterDriverMissingError(printer_name)
