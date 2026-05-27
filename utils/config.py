import json
from dataclasses import asdict
from pathlib import Path

from models.app_settings import AppSettings
from utils.constants import APP_SETTINGS_FILE, LOCAL_APPDATA_ROOT


class ConfigManager:
    def __init__(self, base_dir: Path | None = None) -> None:
        self.base_dir = base_dir or LOCAL_APPDATA_ROOT
        self.settings_path = self.base_dir / APP_SETTINGS_FILE

    def load(self) -> AppSettings:
        if not self.settings_path.exists():
            return AppSettings()

        data = json.loads(self.settings_path.read_text(encoding="utf-8"))
        data["excel_path"] = ""
        data.pop("word_template_path", None)
        data.pop("save_word_copies", None)
        return AppSettings.from_dict(data)

    def save(self, settings: AppSettings) -> None:
        self.base_dir.mkdir(parents=True, exist_ok=True)
        self.settings_path.write_text(
            json.dumps(asdict(settings), indent=2, ensure_ascii=False),
            encoding="utf-8",
        )
