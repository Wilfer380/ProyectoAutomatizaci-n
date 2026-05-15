from __future__ import annotations

import subprocess
import sys
from pathlib import Path
from tkinter import Tk, messagebox

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from deploy.paths_config_generadoretiquetassap import APP_DISPLAY_NAME, APP_EXE_NAME, INSTALL_APP_DIR


def _show_error(message: str) -> None:
    root = Tk()
    root.withdraw()
    messagebox.showerror(APP_DISPLAY_NAME, message)
    root.destroy()


def main() -> int:
    app_exe = INSTALL_APP_DIR / APP_EXE_NAME
    if not app_exe.exists():
        _show_error(
            f"No se encontró la aplicación instalada.\n\n"
            f"Ruta esperada:\n{app_exe}\n\n"
            "Ejecutá nuevamente el instalador desde el paquete de entrega."
        )
        return 1

    subprocess.Popen([str(app_exe)], cwd=str(INSTALL_APP_DIR))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
