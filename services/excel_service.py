import logging
import openpyxl
from PySide6.QtGui import QImage
from pathlib import Path

from models.asset_record import AssetRecord
from utils.constants import APP_NAME, EXCEL_SHEET_SOURCE, SOURCE_HEADERS


class ExcelService:
    def __init__(self) -> None:
        self.logger = logging.getLogger(APP_NAME)

    def extract_data(self, filepath: str | Path) -> list[AssetRecord]:
        """
        Extracts text and mapped anchored images from the Excel file using openpyxl.
        Returns a list of AssetRecord.
        """
        self.logger.info("Extracting data from %s", filepath)
        wb = openpyxl.load_workbook(filepath, data_only=True)

        if EXCEL_SHEET_SOURCE not in wb.sheetnames:
            # Fallback to active sheet if Hoja1 is missing
            ws = wb.active
        else:
            ws = wb[EXCEL_SHEET_SOURCE]

        # Find column indices for headers
        headers = {}
        for col_idx in range(1, ws.max_column + 1):
            cell_value = ws.cell(row=1, column=col_idx).value
            if cell_value:
                headers[str(cell_value).strip()] = col_idx

        # We need these columns based on SOURCE_HEADERS
        # Example SOURCE_HEADERS: {"asset_id": "Activo fijo", "asset_name": "Denominación del activo fijo", "section": "Seccion"}

        # We will fallback to 1, 2, 3 if headers aren't found for the simple test
        col_id = headers.get(SOURCE_HEADERS.get("asset_id", "ID"), 1)
        col_name = headers.get(SOURCE_HEADERS.get("asset_name", "Name"), 2)
        col_section = headers.get(SOURCE_HEADERS.get("section", "Section"), 3)

        # Map images by row
        images_by_row: dict[int, QImage] = {}
        if hasattr(ws, "_images"):
            for img in ws._images:
                # anchor is a OneCellAnchor or TwoCellAnchor
                anchor = img.anchor
                if anchor and hasattr(anchor, "_from"):
                    # openpyxl uses 0-based indexing for row/col in anchors
                    row_idx = anchor._from.row + 1

                    try:
                        # Extract image bytes
                        img_bytes = img._data()
                        qimage = QImage.fromData(img_bytes)
                        if not qimage.isNull():
                            images_by_row[row_idx] = qimage
                    except Exception as e:
                        self.logger.warning(
                            "Failed to extract image at row %s: %s", row_idx, e
                        )

        records = []
        for row_idx in range(2, ws.max_row + 1):
            asset_id = ws.cell(row=row_idx, column=col_id).value

            # Skip empty rows (at least ID must be present)
            if not asset_id:
                continue

            asset_name = ws.cell(row=row_idx, column=col_name).value or ""
            section = ws.cell(row=row_idx, column=col_section).value or ""

            image = images_by_row.get(row_idx)

            record = AssetRecord(
                row_index=row_idx,
                asset_id=str(asset_id),
                asset_name=str(asset_name),
                section=str(section),
                image=image,
            )
            records.append(record)

        wb.close()
        return records
