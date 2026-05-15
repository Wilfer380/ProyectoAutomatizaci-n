from __future__ import annotations

from pathlib import Path


APP_NAME = "GeneradorEtiquetasSAP"
APP_DISPLAY_NAME = "GeneradorEtiquetasSAP"
APP_EXE_NAME = "GeneradorEtiquetasSAP.exe"
LAUNCHER_EXE_NAME = "Launcher_GeneradorEtiquetasSAP.exe"
INSTALLER_EXE_NAME = "Installer_GeneradorEtiquetasSAP.exe"
RELEASE_SHORT_NAME = "GESAP"

RELEASE_VERSION = "0.05.14.4"

INSTALL_ROOT = Path(r"C:\PDI_APP\GeneradorEtiquetasSAP")
INSTALL_APP_DIR = INSTALL_ROOT / "app"
APP_PAYLOAD_DIR_NAME = "app"

PROJECT_ROOT = Path(__file__).resolve().parents[1]
ASSETS_DIR = PROJECT_ROOT / "assets"
ICON_ICO = ASSETS_DIR / "app_icon.ico"
ICON_PNG = ASSETS_DIR / "app_icon.png"
ICON_SVG = ASSETS_DIR / "app_icon.svg"

DIST_RELEASE_DIR = PROJECT_ROOT / "dist_release"
RELEASE_DIR_NAME = f"{RELEASE_SHORT_NAME}_{RELEASE_VERSION}"
RELEASE_DIR = DIST_RELEASE_DIR / RELEASE_DIR_NAME
RELEASE_ZIP = DIST_RELEASE_DIR / f"{RELEASE_DIR_NAME}.zip"
