from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import openpyxl

from services.validation_service import ValidationError, ValidationService


class ValidationServiceTests(unittest.TestCase):
    def test_validate_excel_file_accepts_xlsx(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "input.xlsx"
            openpyxl.Workbook().save(path)

            result = ValidationService().validate_excel_file(str(path))

            self.assertEqual(path, result)

    def test_validate_excel_file_rejects_xls(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "input.xls"
            path.write_bytes(b"legacy xls placeholder")

            with self.assertRaises(ValidationError):
                ValidationService().validate_excel_file(str(path))


if __name__ == "__main__":
    unittest.main()
