from __future__ import annotations

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

    def _build_word_document(
        self,
        template_path: Path,
        output_path: Path,
        image_paths: list[Path],
        placeholder_count: int,
        *,
        block_index: int,
        total_blocks: int,
        log: LogCallback,
    ) -> None:
        try:
            self.word_service.build_document_from_template(
                str(template_path),
                str(output_path),
                image_paths,
                placeholder_count,
            )
            log(f"Bloque {block_index}/{total_blocks}: documento Word generado con plantilla controlada.")
            return
        except Exception as exc:
            self.logger.exception(
                "Falló la generación Word con COM para el bloque %s/%s. Se intentará fallback sin COM.",
                block_index,
                total_blocks,
            )
            log(
                f"Bloque {block_index}/{total_blocks}: Word COM falló al armar el documento "
                f"({ErrorGuard.friendly_message(exc)}). Se intentará fallback sin COM."
            )

        self.word_service.build_document_without_com(
            str(template_path),
            str(output_path),
            image_paths,
            placeholder_count,
        )
        log(f"Bloque {block_index}/{total_blocks}: documento Word generado con fallback sin COM.")

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
        self.validation_service.validate_word_template_placeholders(word_path_obj)
        selected_filter = self.validation_service.validate_selected_filter(selected_filter)
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
            log("Impresión automática deshabilitada. Los documentos quedarán listos para revisión e impresión manual.")
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

            blocks = self._plan_blocks(records)
            for block in blocks:
                self.validation_service.validate_block_size(len(block), BLOCK_SIZE)

            recoverable_failures = 0
            if simulate:
                recoverable_failures = self._simulate_blocks(
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
                recoverable_failures = self._print_blocks(
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
            if recoverable_failures:
                self.logger.warning(
                    "Proceso finalizado con %s bloque(s) omitido(s) por errores recuperables.",
                    recoverable_failures,
                )
                log(
                    f"Proceso completado con {recoverable_failures} bloque(s) omitido(s) por errores recuperables. "
                    f"Se procesaron {len(blocks)} bloque(s) y {len(records)} registro(s)."
                )
            else:
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
    ) -> int:
        if simulation_dir is None:
            raise ValidationError("No se pudo crear la carpeta de simulaci?n.")

        total_blocks = len(blocks)
        total_labels = sum(len(block) for block in blocks)
        manual_adjust = manual_adjust or (lambda _block, _total, _path, _mtime: "continuar")
        recoverable_failures = 0
        last_progress = 0

        def report_progress(value: int) -> None:
            nonlocal last_progress
            value = int(value)
            if value <= last_progress:
                return
            last_progress = value
            progress(value)

        for block_index, block in enumerate(blocks, start=1):
            runtime_word_path = simulation_dir / self._block_document_name(selected_filter, block_index)
            word_opened = False
            block_failed = False
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
                    report_progress,
                    (block_index - 1) * BLOCK_SIZE,
                    total_labels,
                )
                shutil.copy2(word_path, runtime_word_path)
                self._build_word_document(
                    word_path,
                    runtime_word_path,
                    image_paths,
                    len(block),
                    block_index=block_index,
                    total_blocks=total_blocks,
                    log=log,
                )
                image_validation = self.word_service.validate_embedded_image_count(len(block))
                baseline_mtime_ns = self._safe_mtime_ns(runtime_word_path)
                log(
                    f"Bloque {block_index}/{total_blocks}: documento Word listo para revisi?n manual. "
                    f"Im?genes detectadas: {image_validation.detected_count}/{image_validation.expected_count}."
                )
                log(f"Plantilla generada correctamente: {runtime_word_path}")
            except ValidationError:
                raise
            except ManualAdjustmentCancelled:
                raise
            except Exception as exc:
                block_failed = True
                recoverable_failures += 1
                self._log_recoverable_block_failure(
                    action="simulación",
                    block_index=block_index,
                    total_blocks=total_blocks,
                    selected_filter=selected_filter,
                    document_path=runtime_word_path,
                    exception=exc,
                    log=log,
                    status=status,
                )

            if not block_failed:
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
            if block_failed:
                report_progress(int((block_index / total_blocks) * 95))
            else:
                report_progress(int((block_index / total_blocks) * 95))

        report_progress(100)
        self.logger.info("Último bloque completado. Progreso total alcanzado: 100%%.")
        return recoverable_failures

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
    ) -> int:
        manual_adjust = manual_adjust or (lambda _block, _total, _path, _mtime: "continuar")
        total_blocks = len(blocks)
        total_labels = sum(len(block) for block in blocks)
        processed_labels = 0
        recoverable_failures = 0
        last_progress = 0

        def report_progress(value: int) -> None:
            nonlocal last_progress
            value = int(value)
            if value <= last_progress:
                return
            last_progress = value
            progress(value)

        for block_index, block in enumerate(blocks, start=1):
            block_word_path = output_dir / self._block_document_name(word_path.stem, block_index)
            status(f"Generando bloque {block_index} de {total_blocks}")
            log(f"Inicio de generaci?n por bloques: bloque {block_index}/{total_blocks}, cantidad {len(block)}. Copia: {block_word_path}")

            image_dir = output_dir / TEMP_IMAGES_DIR_NAME / f"bloque_{block_index:03d}"
            log(
                f"Transición Excel headless a Word plantilla: bloque {block_index}/{total_blocks} listo. "
                f"Temp workspace activo={workspace_dir.exists()} plantilla Word origen={word_path.exists()}."
            )
            if self.excel_service.workbook is None:
                self.excel_service.open(str(runtime_excel_path), visible=False)
            block_failed = False
            try:
                image_paths = self._prepare_block_images(
                    block,
                    block_index,
                    total_blocks,
                    selected_filter,
                    image_dir,
                    log,
                    report_progress,
                    processed_labels,
                    total_labels,
                )
                self._build_word_document(
                    word_path,
                    block_word_path,
                    image_paths,
                    len(block),
                    block_index=block_index,
                    total_blocks=total_blocks,
                    log=log,
                )
                image_validation = self.word_service.validate_embedded_image_count(len(block))
                log(f"Bloque {block_index}/{total_blocks}: documento Word guardado en {block_word_path}.")

                baseline_mtime_ns = self._safe_mtime_ns(block_word_path)
                log(
                    f"Bloque {block_index}/{total_blocks}: documento listo para revisión e impresión manual. "
                    f"Im?genes detectadas: {image_validation.detected_count}/{image_validation.expected_count}."
                )
            except ValidationError:
                raise
            except ManualAdjustmentCancelled:
                raise
            except Exception as exc:
                block_failed = True
                recoverable_failures += 1
                self._log_recoverable_block_failure(
                    action="impresión",
                    block_index=block_index,
                    total_blocks=total_blocks,
                    selected_filter=selected_filter,
                    document_path=block_word_path,
                    exception=exc,
                    log=log,
                    status=status,
                )

            try:
                if not block_failed:
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
                self.word_service.close(save_changes=False)

            processed_labels += len(block)
            report_progress(int((processed_labels / total_labels) * 100))
        self.logger.info("Último bloque completado. Progreso total alcanzado: 100%%.")
        report_progress(100)
        log(f"Fin de generaci?n por bloques. Total bloques: {total_blocks}; etiquetas: {total_labels}.")
        return recoverable_failures

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
        except Exception as exc:
            raise RuntimeError(
                f"No se pudieron preparar los datos del bloque {block_index}/{total_blocks} para el filtro '{selected_filter}'."
            ) from exc
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
            except Exception as exc:
                self.logger.exception(
                    "Error exportando imagen del bloque %s/%s. Filtro=%s Posici\u00f3n=%s Archivo=%s",
                    block_index,
                    total_blocks,
                    selected_filter,
                    position + 1,
                    image_path,
                )
                log(
                    f"Bloque {block_index}/{total_blocks}: no se pudo exportar la etiqueta {position + 1}/{len(block)}; "
                    f"se dejar\u00e1 vac\u00eda."
                )
                image_paths.append(image_path)
            else:
                image_paths.append(Path(exported.output_path))
                log(
                    f"Bloque {block_index}/{total_blocks}: etiqueta {position + 1}/{len(block)} exportada desde "
                    f"{exported.group_name}; tama\u00f1o final {exported.target_size_px[0]}x{exported.target_size_px[1]} px; "
                    f"archivo {exported.output_path}."
                )

            if progress is not None and total_labels:
                progress(int(((progress_base + position + 1) / total_labels) * 100))

        return image_paths

    def _log_recoverable_block_failure(
        self,
        *,
        action: str,
        block_index: int,
        total_blocks: int,
        selected_filter: str,
        document_path: Path,
        exception: BaseException,
        log: LogCallback,
        status: StatusCallback,
    ) -> None:
        self.logger.exception(
            "Error recuperable durante %s del bloque %s/%s. Filtro=%s Documento=%s",
            action,
            block_index,
            total_blocks,
            selected_filter,
            document_path,
        )
        status(
            f"Bloque {block_index}/{total_blocks}: se omitirá por un error recuperable. "
            "Se continuará con el siguiente bloque."
        )
        log(
            f"Bloque {block_index}/{total_blocks}: error recuperable durante {action}. "
            f"Filtro={selected_filter}. Documento={document_path}. Detalle: {exception}. "
            "Se continuará con el siguiente bloque."
        )

    @staticmethod
    def _plan_blocks(records: list[AssetRecord]) -> list[list[AssetRecord]]:
        return [records[start : start + BLOCK_SIZE] for start in range(0, len(records), BLOCK_SIZE)]

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

