from __future__ import annotations

from utils.constants import TARGET_PRINTER_NAME

try:
    import win32print
except ImportError:  # pragma: no cover
    win32print = None


class PrintService:
    def get_installed_printers(self) -> list[str]:
        if win32print is None:
            return []

        return [
            info[2]
            for info in win32print.EnumPrinters(
                win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            )
        ]

    def set_default_printer(self, printer_name: str = TARGET_PRINTER_NAME) -> None:
        if win32print is None:
            raise RuntimeError("win32print no está disponible.")
        win32print.SetDefaultPrinter(printer_name)
