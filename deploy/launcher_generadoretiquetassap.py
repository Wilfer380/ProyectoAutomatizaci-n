from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path
from tkinter import Tk, messagebox

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from deploy.paths_config_generadoretiquetassap import (
    APP_DISPLAY_NAME,
    APP_EXE_NAME,
    INSTALL_APP_DIR,
    INSTALL_ROOT,
    INSTALLER_EXE_NAME,
    LATEST_RELEASE_DIR,
)


def _show_error(message: str) -> None:
    root = Tk()
    root.withdraw()
    messagebox.showerror(APP_DISPLAY_NAME, message)
    root.destroy()


def _read_version(path: Path) -> tuple[int, ...]:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        version = str(data.get("version", "0")).strip()
    except Exception:
        version = "0"
    parts = []
    for part in version.split("."):
        try:
            parts.append(int(part))
        except ValueError:
            parts.append(0)
    return tuple(parts) if parts else (0,)


def _version_text(path: Path) -> str:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return str(data.get("version", "0")).strip() or "0"
    except Exception:
        return "0"


def _launch_installer(source_root: Path) -> int:
    installer_exe = source_root / INSTALLER_EXE_NAME
    if not installer_exe.exists():
        _show_error(
            f"No se encontró el actualizador de la nueva versión.\n\nRuta esperada:\n{installer_exe}"
        )
        return 1
    subprocess.Popen([str(installer_exe)], cwd=str(source_root))
    return 0


def main() -> int:
    app_exe = INSTALL_APP_DIR / APP_EXE_NAME
    latest_version_file = LATEST_RELEASE_DIR / "version.json"
    installed_version_file = INSTALL_ROOT / "version.json"

    if latest_version_file.exists():
        latest_version = _read_version(latest_version_file)
        installed_version = _read_version(installed_version_file) if installed_version_file.exists() else (0,)
        if latest_version > installed_version:
            root = Tk()
            root.withdraw()
            response = messagebox.askyesno(
                APP_DISPLAY_NAME,
                f"Hay una nueva versión disponible ({_version_text(latest_version_file)}).\n\n¿Querés actualizar ahora?",
            )
            root.destroy()
            if response:
                return _launch_installer(LATEST_RELEASE_DIR)

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
