from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook
from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from models.asset_record import AssetRecord
from services.excel_service import ExcelService


class ExcelServiceHeadlessTests(unittest.TestCase):
    def test_loads_filters_without_excel_com(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "source.xlsx"
            self._write_workbook(path)

            service = ExcelService()
            service.open(str(path))

            self.assertEqual(["Hoja1"], service.get_sheet_names())
            self.assertEqual(
                ["Activo fijo", "Denominación del activo fijo", "Seccion"],
                service.get_headers(),
            )
            self.assertEqual(["B1_WTC", "B2_WTC"], service.get_filters())

    def test_renders_label_image_from_record_data(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "label.png"
            service = ExcelService()
            service.write_block_to_label_sheet(
                [
                    AssetRecord(
                        row_index=2,
                        asset_id="101478",
                        asset_name="SISTEMA DE AIRE ACONDICIONADO",
                        section="Administracion_WTC",
                    )
                ]
            )

            result = service.export_label_shape_image(0, output)

            self.assertTrue(output.exists())
            with Image.open(output) as image:
                self.assertEqual((412, 210), image.size)
            self.assertEqual(1, result.position)

    @staticmethod
    def _write_workbook(path: Path) -> None:
        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "Hoja1"
        worksheet.append(["Activo fijo", "Denominación del activo fijo", "Seccion"])
        worksheet.append(["101478", "SISTEMA DE AIRE ACONDICIONADO", "B1_WTC"])
        worksheet.append(["101479", "MOTOR PRINCIPAL", "B2_WTC"])
        worksheet.append(["101480", "TABLERO DE CONTROL", "B1_WTC"])
        workbook.save(path)


if __name__ == "__main__":
    unittest.main()
