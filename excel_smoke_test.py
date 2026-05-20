from __future__ import annotations

import argparse
import tempfile
from datetime import datetime
from pathlib import Path

from services.excel_service import ExcelService
from services.validation_service import ValidationService
from utils.constants import SAFE_TEMP_ROOT
from utils.error_guard import ErrorGuard
from utils.logger import setup_logger


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Prueba aislada de apertura y escritura en Excel COM.")
    parser.add_argument("--excel", required=True, help="Ruta del archivo Excel real a probar")
    parser.add_argument("--sheet", default="Etiqueta provisional", help="Hoja donde escribir el dato de prueba")
    parser.add_argument("--cell", default="J2", help="Celda donde escribir el dato de prueba")
    parser.add_argument("--value", default=f"SMOKE_{datetime.now():%Y%m%d_%H%M%S}", help="Valor de prueba")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    logger = setup_logger(debug=True)
    validator = ValidationService()
    excel_service = ExcelService()

    excel_path = validator.validate_excel_file(args.excel)
    temp_root = validator.validate_directory_writable(SAFE_TEMP_ROOT)
    smoke_dir = Path(tempfile.mkdtemp(prefix="excel_smoke_", dir=str(temp_root)))
    temp_excel_path = smoke_dir / excel_path.name

    logger.info("Smoke test Excel: origen=%s temp=%s", excel_path, temp_excel_path)

    try:
        temp_excel_path.write_bytes(excel_path.read_bytes())
        excel_service.open(str(temp_excel_path), visible=False)
        logger.info("Excel abierto correctamente en smoke test.")

        worksheet = excel_service.workbook.Worksheets(args.sheet)
        worksheet.Range(args.cell).Value = args.value
        excel_service.excel_app.CalculateFull()
        excel_service.workbook.Save()

        read_back = worksheet.Range(args.cell).Value
        logger.info("Valor escrito en %s!%s = %s", args.sheet, args.cell, read_back)
        print(f"OK: {args.sheet}!{args.cell} = {read_back}")
        return 0
    except Exception as exc:
        logger.exception("Smoke test de Excel falló.")
        print(ErrorGuard.friendly_message(exc))
        return 1
    finally:
        try:
            excel_service.close(save_changes=False)
        except Exception:
            logger.exception("Fallo al cerrar Excel en smoke test.")


if __name__ == "__main__":
    raise SystemExit(main())
