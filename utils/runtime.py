from __future__ import annotations

import os
import sys
from pathlib import Path


def get_main_entrypoint() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve()
    return Path(__file__).resolve().parents[1] / "main.py"


def build_worker_command() -> list[str]:
    entrypoint = get_main_entrypoint()
    debug_enabled = os.environ.get("AUTOMATIZACION_SAP_DEBUG", "").strip().lower() in {"1", "true", "yes", "on"}
    if getattr(sys, "frozen", False):
        command = [str(entrypoint)]
    else:
        command = [sys.executable, str(entrypoint)]
    if debug_enabled:
        command.append("--debug")
    return command


def get_user_home_dir() -> Path:
    userprofile = os.environ.get("USERPROFILE")
    if userprofile:
        return Path(userprofile)

    home_drive = os.environ.get("HOMEDRIVE", "")
    home_path = os.environ.get("HOMEPATH", "")
    if home_drive and home_path:
        return Path(home_drive + home_path)

    return Path.cwd()
