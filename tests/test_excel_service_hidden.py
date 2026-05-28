from __future__ import annotations

import inspect
import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import services.excel_service as excel_service_module
from services.excel_service import ExcelService


class ExcelServiceHeadlessSourceTests(unittest.TestCase):
    def test_excel_service_source_no_longer_depends_on_excel_com(self) -> None:
        source = inspect.getsource(excel_service_module)
        self.assertIn("load_workbook(", source)
        for token in (
            'DispatchEx("Excel.Application")',
            "CopyPicture(",
            "CoInitialize",
            "win32com",
            "pythoncom",
        ):
            self.assertNotIn(token, source)

    def test_open_rejects_legacy_xls_for_headless_mode(self) -> None:
        service = ExcelService()
        with self.assertRaises(RuntimeError):
            service.open(str(Path("book.xls")))


if __name__ == "__main__":
    unittest.main()
