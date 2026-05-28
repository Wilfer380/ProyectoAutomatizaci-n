from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from models.asset_record import AssetRecord
from services.process_service import ProcessService


class DummyLogger:
    def __init__(self) -> None:
        self.info_calls: list[str] = []
        self.warning_calls: list[str] = []
        self.exception_calls: list[str] = []

    def info(self, message: str, *args, **kwargs) -> None:
        self.info_calls.append(self._format(message, *args))

    def warning(self, message: str, *args, **kwargs) -> None:
        self.warning_calls.append(self._format(message, *args))

    def exception(self, message: str, *args, **kwargs) -> None:
        self.exception_calls.append(self._format(message, *args))

    @staticmethod
    def _format(message: str, *args) -> str:
        return message % args if args else message


class DummyExcelService:
    def __init__(self, failing_block_index: int | None = None, failing_positions: set[int] | None = None) -> None:
        self.workbook = object()
        self.failing_block_index = failing_block_index
        self.failing_positions = failing_positions or set()
        self.current_block_index = 0

    def open(self, workbook_path: str, visible: bool = False) -> None:
        self.workbook = object()

    def close(self, save_changes: bool = False) -> None:
        return None

    def write_block_to_label_sheet(self, block) -> None:
        self.current_block_index += 1

    def get_generated_assets(self, count: int) -> list[str]:
        return [f"A{index:03d}" for index in range(1, count + 1)]

    def export_label_shape_image(self, position: int, image_path: Path, target_px: tuple[int, int]):
        if (self.current_block_index == self.failing_block_index and position == 0) or position in self.failing_positions:
            raise RuntimeError("falla controlada exportando imagen")
        return SimpleNamespace(output_path=str(image_path), group_name="Grupo SAP", target_size_px=target_px)


class DummyWordService:
    def __init__(self) -> None:
        self.generated_paths: list[str] = []
        self.close_calls = 0

    def close(self, save_changes: bool = False) -> None:
        self.close_calls += 1

    def validate_embedded_image_count(self, expected_count: int):
        return SimpleNamespace(detected_count=expected_count, expected_count=expected_count)

    def build_document_from_template(self, template_path: str, output_path: str, image_paths, placeholder_count: int) -> None:
        path = Path(output_path)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.touch()
        self.generated_paths.append(output_path)


class DummyValidationService:
    def validate_assets_match(self, source_assets, generated_assets) -> None:
        return None


class ProcessServiceRecoveryTests(unittest.TestCase):
    def test_print_blocks_continues_after_recoverable_block_failure(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            temp_root = Path(tmpdir)
            word_path = temp_root / "template.docx"
            excel_path = temp_root / "source.xlsx"
            word_path.write_text("template", encoding="utf-8")
            excel_path.write_text("workbook", encoding="utf-8")
            output_dir = temp_root / "output"
            output_dir.mkdir()
            workspace_dir = temp_root / "workspace"
            workspace_dir.mkdir()

            logger = DummyLogger()
            excel_service = DummyExcelService(failing_block_index=1)
            word_service = DummyWordService()
            print_service = SimpleNamespace(set_default_printer=lambda _printer: None)
            service = ProcessService(excel_service, word_service, print_service, DummyValidationService(), logger)

            records = [self._record(index) for index in range(1, 29)]
            blocks = ProcessService._plan_blocks(records)
            progress_values: list[int] = []

            service._print_blocks(
                word_path=word_path,
                runtime_excel_path=excel_path,
                blocks=blocks,
                selected_filter="Filtro SAP",
                output_dir=output_dir,
                workspace_dir=workspace_dir,
                printer_name="Printer",
                log=lambda _message: None,
                status=lambda _message: None,
                progress=progress_values.append,
                manual_adjust=lambda _block, _total, _path, _mtime: "continuar",
            )

            self.assertTrue(any("bloque 1/2" in entry.lower() for entry in logger.exception_calls))
            self.assertGreaterEqual(len(word_service.generated_paths), 1)
            self.assertEqual(100, progress_values[-1])

    def test_simulate_blocks_continues_after_recoverable_block_failure(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            temp_root = Path(tmpdir)
            word_path = temp_root / "template.docx"
            excel_path = temp_root / "source.xlsx"
            word_path.write_text("template", encoding="utf-8")
            excel_path.write_text("workbook", encoding="utf-8")
            simulation_dir = temp_root / "sim"
            simulation_dir.mkdir()

            logger = DummyLogger()
            excel_service = DummyExcelService(failing_block_index=1)
            word_service = DummyWordService()
            print_service = SimpleNamespace(set_default_printer=lambda _printer: None)
            service = ProcessService(excel_service, word_service, print_service, DummyValidationService(), logger)

            records = [self._record(index) for index in range(1, 29)]
            blocks = ProcessService._plan_blocks(records)
            progress_values: list[int] = []

            service._simulate_blocks(
                word_path=word_path,
                runtime_excel_path=excel_path,
                blocks=blocks,
                selected_filter="Filtro SAP",
                simulation_dir=simulation_dir,
                log=lambda _message: None,
                status=lambda _message: None,
                progress=progress_values.append,
                manual_adjust=lambda _block, _total, _path, _mtime: "continuar",
            )

            self.assertTrue(any("bloque 1/2" in entry.lower() for entry in logger.exception_calls))
            self.assertGreaterEqual(len(word_service.generated_paths), 1)
            self.assertEqual(100, progress_values[-1])

    def test_prepare_block_images_continues_after_single_image_export_failure(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            temp_root = Path(tmpdir)
            excel_service = DummyExcelService(failing_positions={1})
            service = ProcessService(
                excel_service,
                DummyWordService(),
                SimpleNamespace(set_default_printer=lambda _printer: None),
                DummyValidationService(),
                DummyLogger(),
            )
            image_dir = temp_root / "images"
            block = [self._record(index) for index in range(1, 4)]

            image_paths = service._prepare_block_images(
                block=block,
                block_index=1,
                total_blocks=1,
                selected_filter="Filtro SAP",
                image_dir=image_dir,
                log=lambda _message: None,
                progress=None,
                progress_base=0,
                total_labels=len(block),
            )

            self.assertEqual(3, len(image_paths))
            self.assertTrue(image_paths[0].name.endswith("label_001.png"))
            self.assertTrue(image_paths[1].name.endswith("label_002.png"))
            self.assertTrue(image_paths[2].name.endswith("label_003.png"))

    @staticmethod
    def _record(index: int) -> AssetRecord:
        return AssetRecord(
            row_index=index,
            asset_id=f"A{index:03d}",
            asset_name=f"Asset {index}",
            section="SAP",
        )


if __name__ == "__main__":
    unittest.main()
