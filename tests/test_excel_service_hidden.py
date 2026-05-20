from __future__ import annotations

import logging
import sys
import types
import unittest
from pathlib import Path
from unittest.mock import patch

from services.excel_service import ExcelService


class FakePythonCom(types.ModuleType):
    def __init__(self) -> None:
        super().__init__("pythoncom")
        self.initialized = 0

    def CoInitialize(self) -> None:
        self.initialized += 1


class FakeWorkbook:
    def __init__(self) -> None:
        self.open_calls: list[dict[str, object]] = []

    def Close(self, SaveChanges: bool) -> None:  # noqa: N802
        return None


class FakeWorkbooks:
    def __init__(self, workbook: FakeWorkbook) -> None:
        self.workbook = workbook

    def Open(self, **kwargs) -> FakeWorkbook:  # noqa: N802
        self.workbook.open_calls.append(kwargs)
        return self.workbook


class FakeExcel:
    def __init__(self, workbook: FakeWorkbook) -> None:
        self.Workbooks = FakeWorkbooks(workbook)
        self.Visible = True
        self.DisplayAlerts = True
        self.ScreenUpdating = True
        self.Interactive = True
        self.EnableEvents = True
        self.AskToUpdateLinks = True


class FakeWin32Client(types.ModuleType):
    def __init__(self, excel: FakeExcel) -> None:
        super().__init__("win32com.client")
        self.excel = excel
        self.dispatch_calls: list[str] = []

    def DispatchEx(self, name: str) -> FakeExcel:  # noqa: N802
        self.dispatch_calls.append(name)
        return self.excel


class ExcelServiceHiddenTests(unittest.TestCase):
    def test_open_uses_private_hidden_excel_instance(self) -> None:
        workbook = FakeWorkbook()
        excel = FakeExcel(workbook)
        pythoncom = FakePythonCom()
        win32_client = FakeWin32Client(excel)
        win32 = types.ModuleType("win32com")
        win32.client = win32_client

        with patch.dict(
            sys.modules,
            {
                "pythoncom": pythoncom,
                "win32com": win32,
                "win32com.client": win32_client,
            },
        ):
            service = ExcelService()
            service.open(str(Path("book.xlsx")), visible=False)

        self.assertEqual(["Excel.Application"], win32_client.dispatch_calls)
        self.assertEqual(1, pythoncom.initialized)
        self.assertFalse(excel.Visible)
        self.assertFalse(excel.DisplayAlerts)
        self.assertFalse(excel.ScreenUpdating)
        self.assertFalse(excel.Interactive)
        self.assertFalse(excel.EnableEvents)
        self.assertFalse(excel.AskToUpdateLinks)
        self.assertTrue(service._owns_excel_app)
        self.assertIsNotNone(service.workbook)


if __name__ == "__main__":
    unittest.main()
