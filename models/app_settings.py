from dataclasses import dataclass, field

from utils.constants import TARGET_PRINTER_NAME


@dataclass(slots=True)
class AppSettings:
    excel_path: str = ""
    selected_filter: str = ""
    filter_cache: dict[str, list[str]] = field(default_factory=dict)
    printer_name: str = TARGET_PRINTER_NAME
    working_directory: str = ""
    output_directory: str = ""

    @classmethod
    def from_dict(cls, data: dict) -> "AppSettings":
        return cls(
            excel_path=data.get("excel_path", ""),
            selected_filter=data.get("selected_filter", ""),
            filter_cache={str(key): [str(item) for item in value] for key, value in dict(data.get("filter_cache", {})).items()},
            printer_name=data.get("printer_name", TARGET_PRINTER_NAME),
            working_directory=data.get("working_directory", ""),
            output_directory=data.get("output_directory", ""),
        )
