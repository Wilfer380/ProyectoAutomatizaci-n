from __future__ import annotations

import subprocess
from collections.abc import Callable, Sequence

from utils.constants import TARGET_PRINTER_NAME

DRIVER_GUIDANCE = (
    f"Controlador {TARGET_PRINTER_NAME} no detectado.\n\n"
    "Para imprimir etiquetas reales, este computador necesita el controlador/Printer Utility "
    "de SATO WS4 instalado y la impresora registrada en Windows con el nombre exacto "
    f"'{TARGET_PRINTER_NAME}'.\n\n"
    "Si no tenés permisos de instalación o el instalador del controlador falla, contactá a TI "
    "o al departamento de informática."
)


def build_get_printer_command(printer_name: str = TARGET_PRINTER_NAME) -> list[str]:
    escaped = printer_name.replace("'", "''")
    return [
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        f"if (Get-Printer -Name '{escaped}' -ErrorAction SilentlyContinue) {{ exit 0 }} else {{ exit 1 }}",
    ]


def is_printer_installed_for_installer(
    printer_name: str = TARGET_PRINTER_NAME,
    *,
    runner: Callable[..., subprocess.CompletedProcess] = subprocess.run,
) -> bool:
    try:
        result = runner(
            build_get_printer_command(printer_name),
            capture_output=True,
            text=True,
            timeout=10,
            check=False,
        )
    except Exception:
        return False
    return result.returncode == 0


def format_available_driver_artifacts(paths: Sequence[str]) -> str:
    if not paths:
        return DRIVER_GUIDANCE
    return (
        DRIVER_GUIDANCE
        + "\n\nInstaladores detectados:\n"
        + "\n".join(f"- {path}" for path in paths)
    )
