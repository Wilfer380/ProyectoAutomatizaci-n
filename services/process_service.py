from __future__ import annotations

import subprocess
import shutil
import tempfile
from datetime import datetime
from collections.abc import Callable
from pathlib import Path

from models.asset_record import AssetRecord
from services.excel_service import ExcelService
from services.print_service import PrintService
from services.validation_service import ValidationError, ValidationService
from services.word_service import WordService
from utils.constants import BLOCK_SIZE, LABEL_IMAGE_HEIGHT_PX, LABEL_IMAGE_WIDTH_PX, SAFE_TEMP_ROOT, TARGET_PRINTER_NAME, TEMP_IMAGES_DIR_NAME
from utils.error_guard import ErrorGuard


LogCallback = Callable[[str], None]
ProgressCallback = Callable[[int], None]
StatusCallback = Callable[[str], None]
RecordCountCallback = Callable[[int], None]
ManualAdjustCallback = Callable[[int, int, str, int], str]


class ManualAdjustmentCancelled(Exception):
    pass


class ProcessService:
    def __init__(
        self,
        excel_service: ExcelService,
        word_service: WordService,
        print_service: PrintService,
        validation_service: ValidationService,
        logger,
    ) -> None:
        self.excel_service = excel_service
        self.word_service = word_service
        self.print_service = print_service
        self.validation_service = validation_service
        self.logger = logger

    def load_filters(self, excel_path: str) -> list[str]:
        self.logger.info("Iniciando carga de filtros desde Excel: %s", excel_path)
        self.validation_service.validate_excel_file(excel_path)
        try:
            self.excel_service.open(excel_path, visible=False)
            self.validation_service.validate_required_sheets(self.excel_service.get_sheet_names())
            self.validation_service.validate_required_headers(self.excel_service.get_headers())
            filters = self.excel_service.get_filters()
            self.logger.info("Carga de filtros finalizada. Cantidad: %s", len(filters))
            return filters
        except Exception as exc:
            self.logger.exception("Error durante la carga de filtros. Excel=%s", excel_path)
            raise ValidationError(f"No se pudo cargar los filtros desde el Excel: {excel_path}") from exc
        finally:
            self.excel_service.close(save_changes=False)

    def run(
        self,
        excel_path: str,
        word_path: str,
        selected_filter: str,
        printer_name: str = TARGET_PRINTER_NAME,
        simulate: bool = False,
        log_callback: LogCallback | None = None,
        progress_callback: ProgressCallback | None = None,
        status_callback: StatusCallback | None = None,
        record_count_callback: RecordCountCallback | None = None,
        manual_adjust_callback: ManualAdjustCallback | None = None,
        working_directory: str | None = None,
        output_directory: str | None = None,
    ) -> None:
        log = log_callback or (lambda message: self.logger.info(message))
        progress = progress_callback or (lambda _: None)
        status = status_callback or (lambda _: None)
        record_count = record_count_callback or (lambda _: None)

        excel_path_obj = self.validation_service.validate_excel_file(excel_path)
        word_path_obj = self.validation_service.validate_word_file(word_path)
        selected_filter = self.validation_service.validate_selected_filter(selected_filter)
        if not simulate:
            self.validation_service.validate_printer_installed(printer_name)

        output_dir = Path(output_directory or working_directory or excel_path_obj.parent)
        simulation_dir = self._build_simulation_dir(output_dir, selected_filter) if simulate else None

        self.logger.info(
            "Proceso iniciado. SimulaciÃ³n=%s Excel=%s Word=%s Filtro=%s Output=%s",
            simulate,
            excel_path_obj,
            word_path_obj,
            selected_filter,
            output_dir,
        )

        log(f"Archivo Excel fuente: {excel_path_obj}")
        log(f"Plantilla Word fuente: {word_path_obj}")
        log(f"Filtro seleccionado: {selected_filter}")
        if simulate:
            log(f"Modo prueba visual activo. Evidencias en: {simulation_dir}")
        else:
            log(f"Impresora objetivo: {printer_name}")
        status("Preparando copias temporales de trabajo...")

        temp_root = self.validation_service.validate_directory_writable(SAFE_TEMP_ROOT)
        temp_dir = tempfile.mkdtemp(prefix="automatizacion_sap_", dir=str(temp_root))
        temp_path = Path(temp_dir)
        try:
            runtime_excel_path = temp_path / excel_path_obj.name
            try:
                shutil.copy2(excel_path_obj, runtime_excel_path)
            except Exception as exc:
                self.logger.exception("Error copiando Excel temporal. Excel=%s Temp=%s", excel_path_obj, runtime_excel_path)
                raise ValidationError(f"No se pudo crear la copia temporal de Excel en '{runtime_excel_path}'.") from exc
            log(f"Excel temporal de trabajo: {runtime_excel_path}")

            try:
                self.excel_service.open(str(runtime_excel_path), visible=False)
            except Exception as exc:
                self.logger.exception("Error abriendo Excel temporal. Excel=%s", runtime_excel_path)
                raise ValidationError(ErrorGuard.friendly_message(exc)) from exc
            log(f"Excel temporal de trabajo abierto correctamente: {runtime_excel_path}")
            self.validation_service.validate_required_sheets(self.excel_service.get_sheet_names())
            self.validation_service.validate_required_headers(self.excel_service.get_headers())

            identifier_audit = self.excel_service.normalize_identifier_column()
            if identifier_audit:
                sample = "; ".join(
                    f"fila {row}: '{raw}' -> '{normalized}'" for row, raw, normalized in identifier_audit[:5]
                )
                log(
                    f"AuditorÃ­a columna A: {len(identifier_audit)} valor(es) normalizados a texto. "
                    f"Ejemplos: {sample}"
                )

            status("Leyendo base de datos...")
            try:
                records = self.excel_service.get_filtered_records(selected_filter)
            except Exception as exc:
                self.logger.exception(
                    "Error filtrando registros. Filtro=%s Excel=%s",
                    selected_filter,
                    runtime_excel_path,
                )
                raise ValidationError(f"No se pudieron obtener registros para el filtro '{selected_filter}'.") from exc
            self.validation_service.validate_filter_records(len(records), selected_filter)
            record_count(len(records))
            log(f"Cantidad de registros encontrados: {len(records)}")

            blocks = [records[start : start + BLOCK_SIZE] for start in range(0, len(records), BLOCK_SIZE)]
            for block in blocks:
                self.validation_service.validate_block_size(len(block), BLOCK_SIZE)

            if simulate:
                self._simulate_blocks(
                    word_path_obj,
                    runtime_excel_path,
                    blocks,
                    selected_filter,
                    simulation_dir,
                    log,
                    status,
                    progress,
                    manual_adjust_callback,
                )
            else:
                real_output_dir = self._build_real_output_dir(output_dir, selected_filter)
                log(f"Ruta de plantillas generadas: {real_output_dir}")
                self._print_blocks(
                    word_path_obj,
                    runtime_excel_path,
                    blocks,
                    selected_filter,
                    real_output_dir,
                    temp_path,
                    printer_name,
                    log,
                    status,
                    progress,
                    manual_adjust_callback,
                )

            status("Proceso finalizado.")
            self.logger.info("Último bloque completado. Progreso total alcanzado: 100%%.")
            log(f"Proceso completado correctamente: {len(blocks)} bloque(s), {len(records)} registro(s).")
            self.logger.info("Proceso finalizado correctamente.")
        except ManualAdjustmentCancelled:
            self.logger.info("Proceso cancelado por el usuario durante revisiÃ³n manual.")
            raise
        except ValidationError:
            self.logger.exception("Error de validaciÃ³n durante el proceso.")
            raise
        except Exception as error:
            self.logger.exception("Error durante el proceso.")
            raise ValidationError(str(error)) from error
        finally:
            self.excel_service.close(save_changes=False)
            shutil.rmtree(temp_path, ignore_errors=True)

    def _simulate_blocks(
        self,
        word_path: Path,
        runtime_excel_path: Path,
        blocks: list[list[AssetRecord]],
        selected_filter: str,
        simulation_dir: Path | None,
        log: LogCallback,
        status: StatusCallback,
        progress: ProgressCallback,
        manual_adjust: ManualAdjustCallback | None,
    ) -> None:
        if simulation_dir is None:
            raise ValidationError("No se pudo crear la carpeta de simulaci?n.")

        total_blocks = len(blocks)
        total_labels = sum(len(block) for block in blocks)
        manual_adjust = manual_adjust or (lambda _block, _total, _path, _mtime: "continuar")

        for block_index, block in enumerate(blocks, start=1):
            runtime_word_path = simulation_dir / self._block_document_name(selected_filter, block_index)
            word_opened = False
            try:
                status(f"Simulando bloque {block_index} de {total_blocks}")
                image_dir = simulation_dir / TEMP_IMAGES_DIR_NAME / f"bloque_{block_index:03d}"
                log(
                    f"Inicio de generaci?n por bloques: simulaci?n bloque {block_index}/{total_blocks}, cantidad {len(block)}. "
                    f"Plantilla origen: {word_path}; salida: {runtime_word_path}"
                )

                if self.excel_service.workbook is None:
                    self.excel_service.open(str(runtime_excel_path), visible=False)
                image_paths = self._prepare_block_images(
                    block,
                    block_index,
                    total_blocks,
                    selected_filter,
                    image_dir,
                    log,
                    progress,
                    (block_index - 1) * BLOCK_SIZE,
                    total_labels,
                )
                shutil.copy2(word_path, runtime_word_path)
                self.word_service.open(str(runtime_word_path), visible=False)
                word_opened = True
                self.word_service.prepare_placeholder_document()
                for position, image_path in enumerate(image_paths, start=1):
                    self.word_service.replace_image_placeholder(position, image_path)

                cleaned_placeholders = self.word_service.clear_unused_image_placeholders(len(block))
                image_validation = self.word_service.validate_embedded_image_count(len(block))
                self.word_service.save_document_copy(str(runtime_word_path))
                baseline_mtime_ns = self._safe_mtime_ns(runtime_word_path)
                self.word_service.show_to_user()
                self.word_service.release_to_user()
                log(
                    f"Bloque {block_index}/{total_blocks}: documento Word COM listo para revisi?n manual. "
                    f"Im?genes detectadas: {image_validation.detected_count}/{image_validation.expected_count}."
                )
                log(f"Plantilla generada correctamente con COM: {runtime_word_path}")
            except Exception:
                self.logger.exception(
                    "Error generando simulación del bloque %s/%s. Filtro=%s Excel=%s Word=%s",
                    block_index,
                    total_blocks,
                    selected_filter,
                    runtime_excel_path,
                    runtime_word_path,
                )
                raise

            try:
                status(
                    f"Ajust?/imprim?/guard? el bloque {block_index}/{total_blocks} en Word. "
                    "Cuando hayas guardado, presion? Continuar."
                )
                log(f"Bloque {block_index}/{total_blocks}: documento abierto para revisi?n visual: {runtime_word_path}")
                action = manual_adjust(block_index, total_blocks, str(runtime_word_path), baseline_mtime_ns).strip().lower()
                if action == "cancelar":
                    raise ManualAdjustmentCancelled("Proceso cancelado por el usuario durante la revisi?n del bloque.")

                saved = self._was_saved_after(runtime_word_path, baseline_mtime_ns)
                log(f"Bloque {block_index}/{total_blocks}: guardado detectado={'s?' if saved else 'no' }.")
                log(f"Bloque {block_index}/{total_blocks}: se contin?a al siguiente bloque.")
            finally:
                if word_opened:
                    self.word_service.close(save_changes=False)

            progress(int((block_index / total_blocks) * 95))

        progress(100)
        self.logger.info("Último bloque completado. Progreso total alcanzado: 100%%.")

    def _print_blocks(
        self,
        word_path: Path,
        runtime_excel_path: Path,
        blocks: list[list[AssetRecord]],
        selected_filter: str,
        output_dir: Path,
        workspace_dir: Path,
        printer_name: str,
        log: LogCallback,
        status: StatusCallback,
        progress: ProgressCallback,
        manual_adjust: ManualAdjustCallback | None,
    ) -> None:
        manual_adjust = manual_adjust or (lambda _block, _total, _path, _mtime: "continuar")
        total_blocks = len(blocks)
        total_labels = sum(len(block) for block in blocks)
        processed_labels = 0
        for block_index, block in enumerate(blocks, start=1):
            block_word_path = output_dir / self._block_document_name(word_path.stem, block_index)
            status(f"Generando bloque {block_index} de {total_blocks}")
            log(f"Inicio de generaci?n por bloques: bloque {block_index}/{total_blocks}, cantidad {len(block)}. Copia: {block_word_path}")

            image_dir = output_dir / TEMP_IMAGES_DIR_NAME / f"bloque_{block_index:03d}"
            log(
                f"Transici?n Excel?Word por COM: bloque {block_index}/{total_blocks} listo. "
                f"Temp workspace activo={workspace_dir.exists()} plantilla Word origen={word_path.exists()}."
            )
            if self.excel_service.workbook is None:
                self.excel_service.open(str(runtime_excel_path), visible=False)
            image_paths = self._prepare_block_images(
                block,
                block_index,
                total_blocks,
                selected_filter,
                image_dir,
                log,
                progress,
                processed_labels,
                total_labels,
            )
            word_opened = False
            try:
                shutil.copy2(word_path, block_word_path)
                self.word_service.open(str(block_word_path), visible=False)
                word_opened = True
                self.word_service.prepare_placeholder_document()
                for position, image_path in enumerate(image_paths, start=1):
                    self.word_service.replace_image_placeholder(position, image_path)

                cleaned_placeholders = self.word_service.clear_unused_image_placeholders(len(block))
                image_validation = self.word_service.validate_embedded_image_count(len(block))
                self.word_service.save_document_copy(str(block_word_path))
                log(f"Bloque {block_index}/{total_blocks}: documento Word COM guardado en {block_word_path}.")

                baseline_mtime_ns = self._safe_mtime_ns(block_word_path)
                self.print_service.set_default_printer(printer_name)
                self.word_service.print_document(printer_name)
                log(
                    f"Bloque {block_index}/{total_blocks}: impresi?n solicitada con Word COM usando '{printer_name}'. "
                    f"Im?genes detectadas: {image_validation.detected_count}/{image_validation.expected_count}."
                )
            except Exception as exc:
                self.logger.warning(
                    "Word COM falló en el bloque %s/%s; usando fallback sin COM. Filtro=%s Documento=%s",
                    block_index,
                    total_blocks,
                    selected_filter,
                    block_word_path,
                    exc_info=True,
                )
                if word_opened:
                    self.word_service.close(save_changes=False)
                    word_opened = False
                self.word_service.build_document_without_com(
                    str(word_path),
                    str(block_word_path),
                    image_paths,
                    len(block),
                )
                baseline_mtime_ns = self._safe_mtime_ns(block_word_path)
                log(
                    f"Bloque {block_index}/{total_blocks}: documento generado con fallback sin COM."
                )

            try:
                status(
                    f"Ajust?/imprim?/guard? el bloque {block_index}/{total_blocks} en Word. "
                    "Cuando hayas guardado, presion? Continuar."
                )
                log(f"Bloque {block_index}/{total_blocks}: documento abierto para revisi?n manual: {block_word_path}")
                action = manual_adjust(block_index, total_blocks, str(block_word_path), baseline_mtime_ns).strip().lower()
                if action == "cancelar":
                    raise ManualAdjustmentCancelled("Proceso cancelado por el usuario durante la revisi?n del bloque.")

                saved = self._was_saved_after(block_word_path, baseline_mtime_ns)
                log(f"Bloque {block_index}/{total_blocks}: guardado detectado={'s?' if saved else 'no' }.")
                log(f"Bloque {block_index}/{total_blocks}: se contin?a al siguiente bloque.")
            finally:
                if word_opened:
                    self.word_service.close(save_changes=False)

            processed_labels += len(block)
            progress(int((processed_labels / total_labels) * 100))
        self.logger.info("Último bloque completado. Progreso total alcanzado: 100%%.")
        log(f"Fin de generaci?n por bloques. Total bloques: {total_blocks}; etiquetas: {total_labels}.")

    def _prepare_block_images(
        self,
        block: list[AssetRecord],
        block_index: int,
        total_blocks: int,
        selected_filter: str,
        image_dir: Path,
        log: LogCallback,
        progress: ProgressCallback | None = None,
        progress_base: int = 0,
        total_labels: int | None = None,
    ) -> list[Path]:
        try:
            self.excel_service.write_block_to_label_sheet(block)
            generated_assets = self.excel_service.get_generated_assets(len(block))
        except Exception:
            self.logger.exception(
                "Error preparando datos del bloque %s/%s. Filtro=%s",
                block_index,
                total_blocks,
                selected_filter,
            )
            raise
        source_assets = [record.asset_id for record in block]
        self.validation_service.validate_assets_match(source_assets, generated_assets)

        image_dir.mkdir(parents=True, exist_ok=True)
        log(
            f"Bloque {block_index}/{total_blocks}: se usarÃ¡n placeholders <img1>...<img{len(block)}> "
            "y luego se limpiarÃ¡n los placeholders restantes hasta <img27>."
        )

        image_paths: list[Path] = []
        for position, record in enumerate(block):
            image_path = image_dir / f"label_{position + 1:03d}.png"
            try:
                exported = self.excel_service.export_label_shape_image(
                    position,
                    image_path,
                    target_px=(LABEL_IMAGE_WIDTH_PX, LABEL_IMAGE_HEIGHT_PX),
                )
            except Exception:
                self.logger.exception(
                    "Error exportando imagen de etiqueta. Filtro=%s Bloque=%s/%s Posición=%s",
                    selected_filter,
                    block_index,
                    total_blocks,
                    position + 1,
                )
                raise
            image_paths.append(Path(exported.output_path))
            log(
                f"Bloque {block_index}/{total_blocks}: etiqueta {position + 1}/{len(block)} exportada desde "
                f"{exported.group_name}; tamaÃ±o final {exported.target_size_px[0]}x{exported.target_size_px[1]} px; "
                f"archivo {exported.output_path}."
            )
            if progress is not None and total_labels:
                progress(int(((progress_base + position + 1) / total_labels) * 100))

        return image_paths

    def _build_simulation_dir(self, output_dir: Path, selected_filter: str) -> Path:
        safe_filter = "".join(char if char.isalnum() or char in ("-", "_") else "_" for char in selected_filter).strip("_")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        simulation_dir = output_dir / "simulaciones" / f"{timestamp}_{safe_filter or 'filtro'}"
        simulation_dir.mkdir(parents=True, exist_ok=True)
        return simulation_dir

    def _build_real_output_dir(self, output_dir: Path, selected_filter: str) -> Path:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        real_output_dir = output_dir / "bloques_impresion" / f"{timestamp}_{self._safe_filter(selected_filter)}"
        real_output_dir.mkdir(parents=True, exist_ok=True)
        return real_output_dir

    def _safe_mtime_ns(self, path: Path) -> int:
        try:
            return path.stat().st_mtime_ns
        except OSError:
            return 0

    def _was_saved_after(self, path: Path, baseline_mtime_ns: int) -> bool:
        return self._safe_mtime_ns(path) > baseline_mtime_ns

    def _safe_filter(self, selected_filter: str) -> str:
        return "".join(char if char.isalnum() or char in ("-", "_") else "_" for char in selected_filter).strip("_") or "filtro"

    def _block_document_name(self, selected_filter: str, block_index: int) -> str:
        return f"{self._safe_filter(selected_filter)}_bloque_{block_index:03d}.docx"

    def _format_cm(self, value: float | None) -> str:
        return "n/d" if value is None else f"{value:.2f}"

