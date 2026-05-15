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
    if getattr(sys, "frozen", False):
        return [str(entrypoint)]
    return [sys.executable, str(entrypoint)]


def get_user_home_dir() -> Path:
    userprofile = os.environ.get("USERPROFILE")
    if userprofile:
        return Path(userprofile)

    home_drive = os.environ.get("HOMEDRIVE", "")
    home_path = os.environ.get("HOMEPATH", "")
    if home_drive and home_path:
        return Path(home_drive + home_path)

    return Path.cwd()
