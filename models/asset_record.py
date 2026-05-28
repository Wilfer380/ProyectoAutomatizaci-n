from dataclasses import dataclass
from typing import Any


@dataclass(slots=True)
class AssetRecord:
    row_index: int
    asset_id: str
    asset_name: str
    section: str
    image: Any | None = None
