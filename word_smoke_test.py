from __future__ import annotations

import argparse
import tempfile
from pathlib import Path

from services.excel_service import ExcelService
from services.validation_service import ValidationService
from services.word_service import WordService
from utils.constants import SAFE_TEMP_ROOT
from utils.error_guard import ErrorGuard
from utils.logger import setup_logger


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Prueba aislada del flujo Excel -> imagen -> Word.")
    parser.add_argument("--excel", required=True, help="Ruta del Excel real a probar")
    parser.add_argument("--word", required=True, help="Ruta de la plantilla Word real a probar")
    parser.add_argument("--filter", required=True, help="Filtro real a usar para extraer un registro")
    parser.add_argument("--sheet", default="Etiqueta provisional", help="Hoja de salida de Excel")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    logger = setup_logger(debug=True)
    validator = ValidationService()
    excel_service = ExcelService()
    word_service = WordService()

    excel_path = validator.validate_excel_file(args.excel)
    word_path = validator.validate_word_file(args.word)
    temp_root = validator.validate_directory_writable(SAFE_TEMP_ROOT)
    smoke_dir = Path(tempfile.mkdtemp(prefix="word_smoke_", dir=str(temp_root)))
    temp_excel_path = smoke_dir / excel_path.name
    temp_word_path = smoke_dir / word_path.name
    image_path = smoke_dir / "label_001.png"
    output_docx = smoke_dir / f"{word_path.stem}_smoke_output.docx"

    logger.info("Smoke test Word: excel=%s word=%s temp=%s", excel_path, word_path, smoke_dir)

    try:
        temp_excel_path.write_bytes(excel_path.read_bytes())
        temp_word_path.write_bytes(word_path.read_bytes())

        excel_service.open(str(temp_excel_path), visible=False)
        records = excel_service.get_filtered_records(args.filter)
        if not records:
            raise RuntimeError(f"No hay registros para el filtro '{args.filter}'.")

        excel_service.write_block_to_label_sheet([records[0]])
        excel_service.export_label_shape_image(0, image_path)
        logger.info("Excel OK: imagen exportada en %s", image_path)

        word_service.build_document_without_com(str(temp_word_path), str(output_docx), [image_path], 1)
        logger.info("Word sin COM OK: documento guardado en %s", output_docx)

        print(f"OK: Documento generado sin COM en {output_docx}")
        return 0
    except Exception as exc:
        logger.exception("Smoke test de Word falló.")
        print(ErrorGuard.friendly_message(exc))
        return 1
    finally:
        try:
            excel_service.close(save_changes=False)
        except Exception:
            logger.exception("Fallo al cerrar Excel en smoke test.")


if __name__ == "__main__":
    raise SystemExit(main())
