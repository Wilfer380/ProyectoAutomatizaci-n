from __future__ import annotations

import os
import shutil
import subprocess
import sys
import traceback
from pathlib import Path
from tkinter import BOTH, END, LEFT, RIGHT, X, Button, Frame, Label, Text, Tk, messagebox, ttk

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from deploy.paths_config_generadoretiquetassap import (
    APP_DISPLAY_NAME,
    APP_PAYLOAD_DIR_NAME,
    ICON_ICO,
    INSTALL_APP_DIR,
    INSTALL_ROOT,
    INSTALLER_EXE_NAME,
    LAUNCHER_EXE_NAME,
)
from utils.runtime import get_user_home_dir


WEG_BLUE = "#003E7E"
WEG_YELLOW = "#F8C200"
WEG_PALE = "#D8E3F0"


def runtime_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def bundled_resource_path(*parts: str) -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS).joinpath(*parts)
    return Path(__file__).resolve().parents[1].joinpath(*parts)


class InstallerWindow:
    def __init__(self) -> None:
        self.root = Tk()
        self.root.title(f"Instalador {APP_DISPLAY_NAME}")
        self.root.geometry("720x520")
        self.root.minsize(680, 500)
        self.root.configure(bg="white")

        icon = bundled_resource_path("assets", "app_icon.ico")
        if icon.exists():
            try:
                self.root.iconbitmap(str(icon))
            except Exception:
                pass

        self.source_dir = runtime_base_dir()
        self.payload_dir = self.source_dir / APP_PAYLOAD_DIR_NAME
        self.launcher_source = self.source_dir / LAUNCHER_EXE_NAME

        self._build_ui()

    def _build_ui(self) -> None:
        header = Frame(self.root, bg=WEG_BLUE, padx=18, pady=14)
        header.pack(fill=X)
        Label(
            header,
            text=f"Instalador {APP_DISPLAY_NAME}",
            fg="white",
            bg=WEG_BLUE,
            font=("Segoe UI", 18, "bold"),
        ).pack(anchor="w")
        Label(
            header,
            text="Instalación portable local para el generador de etiquetas SAP",
            fg="white",
            bg=WEG_BLUE,
            font=("Segoe UI", 10),
        ).pack(anchor="w")

        body = Frame(self.root, bg="white", padx=18, pady=16)
        body.pack(fill=BOTH, expand=True)

        callout = Frame(body, bg=WEG_PALE, highlightbackground=WEG_YELLOW, highlightthickness=3, padx=14, pady=12)
        callout.pack(fill=X)
        Label(callout, text="Se instalará en:", bg=WEG_PALE, fg=WEG_BLUE, font=("Segoe UI", 11, "bold")).pack(anchor="w")
        Label(callout, text=str(INSTALL_ROOT), bg=WEG_PALE, fg=WEG_BLUE, font=("Consolas", 15, "bold")).pack(anchor="w", pady=(4, 0))
        Label(callout, text=f"Origen: {self.source_dir}", bg=WEG_PALE, fg="#555555", font=("Segoe UI", 9)).pack(anchor="w", pady=(8, 0))

        self.progress = ttk.Progressbar(body, mode="determinate", maximum=100)
        self.progress.pack(fill=X, pady=(16, 8))

        self.status = Label(body, text="Listo para instalar.", bg="white", fg="#333333", font=("Segoe UI", 10, "bold"))
        self.status.pack(anchor="w")

        self.log = Text(body, height=12, wrap="word", bg="#111111", fg="#E6E6E6", insertbackground="white")
        self.log.pack(fill=BOTH, expand=True, pady=(10, 0))

        buttons = Frame(body, bg="white")
        buttons.pack(fill=X, pady=(12, 0))
        Button(buttons, text="Cancelar", command=self.root.destroy, width=14).pack(side=RIGHT, padx=(8, 0))
        self.install_button = Button(buttons, text="▶ Instalar", command=self.install, width=16, bg=WEG_BLUE, fg="white")
        self.install_button.pack(side=RIGHT)

    def _log(self, message: str) -> None:
        self.log.insert(END, message + "\n")
        self.log.see(END)
        self.root.update_idletasks()

    def _set_progress(self, value: int, message: str) -> None:
        self.progress["value"] = value
        self.status.configure(text=message)
        self._log(message)

    def _validate_source(self) -> None:
        if not self.payload_dir.exists():
            raise FileNotFoundError(f"No se encontró {self.payload_dir}")
        if not self.launcher_source.exists():
            raise FileNotFoundError(f"No se encontró {self.launcher_source}")

    def _copy_payload(self) -> None:
        if INSTALL_APP_DIR.exists():
            shutil.rmtree(INSTALL_APP_DIR)
        INSTALL_APP_DIR.parent.mkdir(parents=True, exist_ok=True)
        shutil.copytree(self.payload_dir, INSTALL_APP_DIR)

    def _copy_launcher(self) -> Path:
        INSTALL_ROOT.mkdir(parents=True, exist_ok=True)
        target = INSTALL_ROOT / LAUNCHER_EXE_NAME
        if target.exists():
            target.unlink()
        shutil.copy2(self.launcher_source, target)
        return target

    def _copy_version_file(self) -> None:
        source_version = self.source_dir / "version.json"
        if not source_version.exists():
            return
        INSTALL_ROOT.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source_version, INSTALL_ROOT / "version.json")

    def _launch_installed_app(self, launcher_target: Path) -> None:
        subprocess.Popen([str(launcher_target)], cwd=str(INSTALL_ROOT))

    def _create_desktop_shortcut(self, launcher_target: Path) -> None:
        desktop = get_user_home_dir() / "Desktop"
        shortcut_path = desktop / f"{APP_DISPLAY_NAME}.lnk"
        try:
            import win32com.client

            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortcut(str(shortcut_path))
            shortcut.TargetPath = str(launcher_target)
            shortcut.WorkingDirectory = str(INSTALL_ROOT)
            shortcut.IconLocation = str(launcher_target)
            shortcut.Description = APP_DISPLAY_NAME
            shortcut.Save()
        except Exception as exc:
            self._log(f"No se pudo crear el acceso directo .lnk ({exc}). Se crea acceso .url de respaldo.")
            shortcut_path = desktop / f"{APP_DISPLAY_NAME}.url"
            shortcut_path.write_text(f"[InternetShortcut]\nURL=file:///{launcher_target.as_posix()}\n", encoding="utf-8")

    def install(self) -> None:
        self.install_button.configure(state="disabled")
        try:
            self._set_progress(5, "Validando paquete de instalación...")
            self._validate_source()
            self._set_progress(25, f"Creando carpeta destino: {INSTALL_ROOT}")
            INSTALL_ROOT.mkdir(parents=True, exist_ok=True)
            self._set_progress(45, f"Copiando aplicación a: {INSTALL_APP_DIR}")
            self._copy_payload()
            self._set_progress(75, "Copiando launcher...")
            launcher_target = self._copy_launcher()
            self._set_progress(82, "Guardando versión instalada...")
            self._copy_version_file()
            self._set_progress(90, "Creando acceso directo en el escritorio...")
            self._create_desktop_shortcut(launcher_target)
            self._set_progress(100, "Instalación completada correctamente.")
            self._log("Abriendo la aplicación instalada...")
            self.status.configure(text="Instalación completada. Abriendo la aplicación...", fg="#166534")
            self.root.update_idletasks()
            self._launch_installed_app(launcher_target)
            self.root.after(300, self.root.destroy)
        except Exception as exc:
            self._log(traceback.format_exc())
            self.status.configure(text="La instalación falló.", fg="#B00020")
            messagebox.showerror(APP_DISPLAY_NAME, f"No se pudo instalar.\n\n{exc}")
            self.install_button.configure(state="normal")

    def run(self) -> None:
        self.root.mainloop()


def main() -> int:
    InstallerWindow().run()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
