from __future__ import annotations

import tempfile
from pathlib import Path

from utils.constants import (
    EXCEL_SHEET_LABEL,
    EXCEL_SHEET_SOURCE,
    SOURCE_HEADERS,
    TARGET_PRINTER_NAME,
)
from services.driver_check import PrinterDriverMissingError, ensure_printer_driver
from utils.normalization import normalize_excel_scalar


class ValidationError(Exception):
    pass


class ValidationService:
    def _validate_file_access(self, path: Path, *, readable: bool = True) -> None:
        try:
            if readable:
                with path.open("rb"):
                    pass
            else:
                with path.open("r+b"):
                    pass
        except PermissionError as exc:
            raise ValidationError(
                f"No hay permisos suficientes para acceder a '{path}'."
            ) from exc
        except OSError as exc:
            raise ValidationError(
                f"El archivo '{path}' no se puede abrir en este momento. Puede estar bloqueado por otra aplicación."
            ) from exc

    def validate_excel_file(self, excel_path: str) -> Path:
        path = Path(excel_path)
        if not path.exists():
            raise ValidationError("El archivo Excel no existe.")
        if path.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
            raise ValidationError("El archivo Excel no tiene una extensión válida.")
        self._validate_file_access(path)
        return path

    def validate_word_file(self, word_path: str) -> Path:
        raise ValidationError(
            "La plantilla Word ya no se usa. Seleccioná solo el archivo Excel."
        )

    def validate_directory_writable(self, directory: str | Path) -> Path:
        path = Path(directory)
        path.mkdir(parents=True, exist_ok=True)
        try:
            with tempfile.NamedTemporaryFile(
                prefix="automatizacion_sap_check_", dir=str(path), delete=True
            ):
                pass
        except PermissionError as exc:
            raise ValidationError(
                f"No se puede escribir en la carpeta '{path}'."
            ) from exc
        except OSError as exc:
            raise ValidationError(
                f"No se pudo validar la carpeta de trabajo '{path}'."
            ) from exc
        return path

    def validate_selected_filter(self, selected_filter: str) -> str:
        value = selected_filter.strip()
        if not value:
            raise ValidationError("Debe seleccionar un filtro antes de comenzar.")
        return value

    def validate_required_sheets(self, sheet_names: list[str]) -> None:
        missing = [
            name
            for name in (EXCEL_SHEET_SOURCE, EXCEL_SHEET_LABEL)
            if name not in sheet_names
        ]
        if missing:
            raise ValidationError(
                f"No se encontraron las hojas requeridas: {', '.join(missing)}."
            )

    def validate_required_headers(self, headers: list[str]) -> None:
        missing = [
            header for header in SOURCE_HEADERS.values() if header not in headers
        ]
        if missing:
            raise ValidationError(
                f"No se encontraron las columnas requeridas: {', '.join(missing)}."
            )

    def validate_filter_records(self, records_count: int, selected_filter: str) -> None:
        if records_count <= 0:
            raise ValidationError(
                f"El filtro seleccionado '{selected_filter}' no tiene registros."
            )

    def validate_block_size(self, records_count: int, block_size: int) -> None:
        if records_count > block_size:
            raise ValidationError(
                f"El bloque actual supera el máximo permitido de {block_size} registros."
            )

    def validate_assets_match(
        self, source_assets: list[str], generated_assets: list[str]
    ) -> None:
        normalized_source = [normalize_excel_scalar(asset) for asset in source_assets]
        normalized_generated = [
            normalize_excel_scalar(asset) for asset in generated_assets
        ]

        if normalized_source == normalized_generated:
            return

        mismatches: list[str] = []
        for index, (source, generated, source_norm, generated_norm) in enumerate(
            zip(
                source_assets,
                generated_assets,
                normalized_source,
                normalized_generated,
                strict=False,
            ),
            start=1,
        ):
            if source_norm != generated_norm:
                mismatches.append(
                    f"Fila {index}: origen='{source}' normalizado='{source_norm}' salida='{generated}' normalizado='{generated_norm}'"
                )

        if len(normalized_source) != len(normalized_generated):
            mismatches.append(
                f"Cantidad distinta de activos: origen={len(normalized_source)} salida={len(normalized_generated)}"
            )

        raise ValidationError(
            "La validación de activos falló. Diferencias encontradas: "
            + "; ".join(mismatches)
        )

    def validate_printer_installed(
        self, printer_name: str = TARGET_PRINTER_NAME
    ) -> None:
        try:
            ensure_printer_driver(printer_name)
        except PrinterDriverMissingError as exc:
            raise ValidationError(str(exc)) from exc
