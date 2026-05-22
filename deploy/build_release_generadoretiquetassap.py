from __future__ import annotations

import json
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from deploy.paths_config_generadoretiquetassap import (
    APP_DISPLAY_NAME,
    APP_EXE_NAME,
    APP_NAME,
    APP_PAYLOAD_DIR_NAME,
    DIST_RELEASE_DIR,
    INSTALL_ROOT,
    INSTALLER_EXE_NAME,
    LAUNCHER_EXE_NAME,
    LATEST_RELEASE_DIR,
    PROJECT_ROOT,
    RELEASE_SHORT_NAME,
    RELEASE_DIR,
    RELEASE_VERSION,
    RELEASE_ZIP,
)


def run(command: list[str]) -> None:
    print("+", " ".join(command), flush=True)
    subprocess.run(command, cwd=PROJECT_ROOT, check=True)


def require_pyinstaller() -> None:
    try:
        import PyInstaller  # noqa: F401
    except Exception as exc:
        raise SystemExit(
            "PyInstaller no está instalado en este Python. Ejecutá: pip install pyinstaller"
        ) from exc


def clean_build_outputs() -> None:
    for relative in ("build", "dist"):
        target = PROJECT_ROOT / relative
        if target.exists():
            shutil.rmtree(target)
    if DIST_RELEASE_DIR.exists():
        for child in DIST_RELEASE_DIR.iterdir():
            if child.is_dir() and child.name.startswith(f"{RELEASE_SHORT_NAME}_"):
                shutil.rmtree(child)
            elif child.is_file() and child.suffix.lower() == ".zip" and child.stem.startswith(f"{RELEASE_SHORT_NAME}_"):
                child.unlink()
    if RELEASE_DIR.exists():
        shutil.rmtree(RELEASE_DIR)
    if LATEST_RELEASE_DIR.exists():
        shutil.rmtree(LATEST_RELEASE_DIR)
    if RELEASE_ZIP.exists():
        RELEASE_ZIP.unlink()
    DIST_RELEASE_DIR.mkdir(parents=True, exist_ok=True)


def build_executables() -> None:
    specs = [
        "GeneradorEtiquetasSAP.spec",
        "Launcher_GeneradorEtiquetasSAP.spec",
        "Installer_GeneradorEtiquetasSAP.spec",
    ]
    for spec in specs:
        run([sys.executable, "-m", "PyInstaller", "--noconfirm", spec])


def write_leeme(target: Path) -> None:
    content = f"""{APP_DISPLAY_NAME} - paquete de instalación portable
Versión: {RELEASE_VERSION}

INSTRUCCIONES
1. Extraiga este ZIP en una carpeta local, por ejemplo Descargas.
2. Ejecute {INSTALLER_EXE_NAME} con doble clic.
3. El instalador copiará la aplicación a:
   {INSTALL_ROOT}
4. Se creará un acceso directo en el escritorio llamado {APP_DISPLAY_NAME}.

REQUISITOS DEL EQUIPO
- Windows.
- Microsoft Excel y Microsoft Word instalados.
- Impresora SATO WS408 instalada con ese nombre exacto para impresión real.

CONTENIDO DEL PAQUETE
- {INSTALLER_EXE_NAME}: instalador gráfico.
- {LAUNCHER_EXE_NAME}: lanzador que queda instalado en {INSTALL_ROOT}.
- {APP_PAYLOAD_DIR_NAME}: aplicación principal y dependencias.
- version.json: versión del paquete.

NOTA
Este paquete no usa TTS/voz. Si el antivirus bloquea la ejecución por ser un EXE interno sin firma,
solicite autorización de TI o agregue la excepción correspondiente.
"""
    target.write_text(content, encoding="utf-8")


def write_version(target: Path) -> None:
    payload = {
        "app": APP_NAME,
        "display_name": APP_DISPLAY_NAME,
        "version": RELEASE_VERSION,
        "install_root": str(INSTALL_ROOT),
        "distribution_mode": "zip-portable-with-installer",
    }
    target.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")


def assemble_release() -> None:
    RELEASE_DIR.mkdir(parents=True, exist_ok=True)
    payload_dir = RELEASE_DIR / APP_PAYLOAD_DIR_NAME
    latest_dir = LATEST_RELEASE_DIR
    app_dist_dir = PROJECT_ROOT / "dist" / APP_NAME
    launcher_exe = PROJECT_ROOT / "dist" / LAUNCHER_EXE_NAME
    installer_exe = PROJECT_ROOT / "dist" / INSTALLER_EXE_NAME

    required = [app_dist_dir / APP_EXE_NAME, launcher_exe, installer_exe]
    missing = [str(path) for path in required if not path.exists()]
    if missing:
        raise FileNotFoundError("No se encontraron artefactos esperados: " + ", ".join(missing))

    shutil.copytree(app_dist_dir, payload_dir)
    shutil.copy2(launcher_exe, RELEASE_DIR / LAUNCHER_EXE_NAME)
    shutil.copy2(installer_exe, RELEASE_DIR / INSTALLER_EXE_NAME)
    write_leeme(RELEASE_DIR / "LEEME.txt")
    write_version(RELEASE_DIR / "version.json")

    if latest_dir.exists():
        shutil.rmtree(latest_dir)
    shutil.copytree(RELEASE_DIR, latest_dir)


def zip_release() -> None:
    with zipfile.ZipFile(RELEASE_ZIP, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for file_path in RELEASE_DIR.rglob("*"):
            if file_path.is_file():
                archive.write(file_path, file_path.relative_to(RELEASE_DIR.parent))


def verify_release() -> None:
    checks = [
        RELEASE_ZIP,
        RELEASE_DIR / INSTALLER_EXE_NAME,
        RELEASE_DIR / LAUNCHER_EXE_NAME,
        RELEASE_DIR / APP_PAYLOAD_DIR_NAME / APP_EXE_NAME,
        RELEASE_DIR / APP_PAYLOAD_DIR_NAME / "_internal",
        RELEASE_DIR / "LEEME.txt",
        RELEASE_DIR / "version.json",
        LATEST_RELEASE_DIR / "version.json",
        LATEST_RELEASE_DIR / APP_PAYLOAD_DIR_NAME / APP_EXE_NAME,
    ]
    missing = [str(path) for path in checks if not path.exists()]
    if missing:
        raise FileNotFoundError("Release incompleto: " + ", ".join(missing))
    print("\nRelease generado correctamente:")
    print(f"- Carpeta: {RELEASE_DIR}")
    print(f"- ZIP: {RELEASE_ZIP}")


def main() -> int:
    require_pyinstaller()
    clean_build_outputs()
    build_executables()
    assemble_release()
    zip_release()
    verify_release()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
