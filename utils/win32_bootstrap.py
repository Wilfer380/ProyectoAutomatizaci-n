from __future__ import annotations

import importlib
import os
import sys
import tempfile
from pathlib import Path


def ensure_pywin32_gen_path() -> Path:
    base_dir = Path(
        os.environ.get("LOCALAPPDATA")
        or os.environ.get("TEMP")
        or os.environ.get("TMP")
        or tempfile.gettempdir()
    ) / "GeneradorEtiquetasSAP" / "gen_py"
    version_dir = base_dir / f"{sys.version_info.major}.{sys.version_info.minor}"
    version_dir.mkdir(parents=True, exist_ok=True)
    init_file = version_dir / "__init__.py"
    if not init_file.exists():
        init_file.write_text("", encoding="utf-8")
    return version_dir


def load_pywin32():
    gen_path = ensure_pywin32_gen_path()
    import win32com

    win32com.__gen_path__ = str(gen_path)
    pythoncom = importlib.import_module("pythoncom")
    win32 = importlib.import_module("win32com.client")
    return pythoncom, win32