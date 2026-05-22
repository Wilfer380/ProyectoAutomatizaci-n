from dataclasses import dataclass, field

from utils.constants import TARGET_PRINTER_NAME


@dataclass(slots=True)
class AppSettings:
    excel_path: str = ""
    word_template_path: str = ""
    selected_filter: str = ""
    filter_cache: dict[str, list[str]] = field(default_factory=dict)
    printer_name: str = TARGET_PRINTER_NAME
    working_directory: str = ""
    save_word_copies: bool = False
    output_directory: str = ""

    @classmethod
    def from_dict(cls, data: dict) -> "AppSettings":
        return cls(
            excel_path=data.get("excel_path", ""),
            word_template_path=data.get("word_template_path", ""),
            selected_filter=data.get("selected_filter", ""),
            filter_cache={str(key): [str(item) for item in value] for key, value in dict(data.get("filter_cache", {})).items()},
            printer_name=data.get("printer_name", TARGET_PRINTER_NAME),
            working_directory=data.get("working_directory", ""),
            save_word_copies=bool(data.get("save_word_copies", False)),
            output_directory=data.get("output_directory", ""),
        )
