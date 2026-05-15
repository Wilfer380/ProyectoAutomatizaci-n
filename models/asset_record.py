from dataclasses import dataclass


@dataclass(slots=True)
class AssetRecord:
    row_index: int
    asset_id: str
    asset_name: str
    section: str
