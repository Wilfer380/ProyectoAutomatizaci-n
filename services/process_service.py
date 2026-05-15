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
from utils.constants import BLOCK_SIZE, LABEL_IMAGE_HEIGHT_PX, LABEL_IMAGE_WIDTH_PX, TARGET_PRINTER_NAME, TEMP_IMAGES_DIR_NAME


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
        except Exception:
            self.logger.exception("Error durante la carga de filtros.")
            raise
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
            "Proceso iniciado. Simulación=%s Excel=%s Word=%s Filtro=%s Output=%s",
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

        temp_dir = tempfile.mkdtemp(prefix="automatizacion_sap_")
        temp_path = Path(temp_dir)
        try:
            runtime_excel_path = temp_path / excel_path_obj.name
            shutil.copy2(excel_path_obj, runtime_excel_path)
            log(f"Excel temporal de trabajo: {runtime_excel_path}")

            self.excel_service.open(str(runtime_excel_path), visible=False)
            self.validation_service.validate_required_sheets(self.excel_service.get_sheet_names())
            self.validation_service.validate_required_headers(self.excel_service.get_headers())

            identifier_audit = self.excel_service.normalize_identifier_column()
            if identifier_audit:
                sample = "; ".join(
                    f"fila {row}: '{raw}' -> '{normalized}'" for row, raw, normalized in identifier_audit[:5]
                )
                log(
                    f"Auditoría columna A: {len(identifier_audit)} valor(es) normalizados a texto. "
                    f"Ejemplos: {sample}"
                )

            status("Leyendo base de datos...")
            records = self.excel_service.get_filtered_records(selected_filter)
            self.validation_service.validate_filter_records(len(records), selected_filter)
            record_count(len(records))
            log(f"Cantidad de registros encontrados: {len(records)}")

            blocks = [records[start : start + BLOCK_SIZE] for start in range(0, len(records), BLOCK_SIZE)]
            for block in blocks:
                self.validation_service.validate_block_size(len(block), BLOCK_SIZE)

            if simulate:
                self._simulate_blocks(
                    word_path_obj,
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
                self._print_blocks(word_path_obj, blocks, real_output_dir, printer_name, log, status, progress, manual_adjust_callback)

            status("Proceso finalizado.")
            log(f"Proceso completado correctamente: {len(blocks)} bloque(s), {len(records)} registro(s).")
            self.logger.info("Proceso finalizado correctamente.")
        except ManualAdjustmentCancelled:
            self.logger.info("Proceso cancelado por el usuario durante revisión manual.")
            raise
        except ValidationError:
            self.logger.exception("Error de validación durante el proceso.")
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
        blocks: list[list[AssetRecord]],
        selected_filter: str,
        simulation_dir: Path | None,
        log: LogCallback,
        status: StatusCallback,
        progress: ProgressCallback,
        manual_adjust: ManualAdjustCallback | None,
    ) -> None:
        if simulation_dir is None:
            raise ValidationError("No se pudo crear la carpeta de simulación.")

        total_blocks = len(blocks)
        manual_adjust = manual_adjust or (lambda _block, _total, _path, _mtime: "continuar")

        for block_index, block in enumerate(blocks, start=1):
            runtime_word_path = simulation_dir / self._block_document_name(selected_filter, block_index)
            try:
                shutil.copy2(word_path, runtime_word_path)
                status(f"Simulando bloque {block_index} de {total_blocks}")
                log(f"Inicio de generación por bloques: simulación bloque {block_index}/{total_blocks}, cantidad {len(block)}. Copia: {runtime_word_path}")

                self.word_service.open(str(runtime_word_path), visible=False)
                self.word_service.prepare_placeholder_document()
                log("Plantilla Word real preparada sobre copia del bloque; se reemplazan placeholders <img1>...<img27> sin crear páginas dinámicas.")
                log(f"Layout visual Word: {self.word_service.get_label_layout_details()}.")
                image_dir = simulation_dir / TEMP_IMAGES_DIR_NAME / f"bloque_{block_index:03d}"
                self._write_and_insert_block(block, block_index, total_blocks, image_dir, log)
                self.word_service.save_document_copy(str(runtime_word_path))
                baseline_mtime_ns = self._safe_mtime_ns(runtime_word_path)
                pdf_path = simulation_dir / f"{runtime_word_path.stem}.pdf"
                self.word_service.export_pdf(str(pdf_path))
                log(f"Simulación guardada en Word: {runtime_word_path}")
                log(f"Simulación guardada en PDF: {pdf_path}")
                baseline_mtime_ns = self._safe_mtime_ns(runtime_word_path)
                self.word_service.show_to_user()
                self.word_service.release_to_user()
                log(f"Plantilla generada correctamente: {runtime_word_path}")
            except Exception:
                self.logger.exception("Error generando simulación del bloque %s/%s.", block_index, total_blocks)
                raise
            finally:
                self.word_service.close(save_changes=False)

            status(
                f"Ajustá/imprimí/guardá el bloque {block_index}/{total_blocks} en Word. "
                "Cuando hayas guardado, presioná Continuar."
            )
            log(f"Bloque {block_index}/{total_blocks}: documento abierto para revisión visual: {runtime_word_path}")
            action = manual_adjust(block_index, total_blocks, str(runtime_word_path), baseline_mtime_ns).strip().lower()
            if action == "cancelar":
                raise ManualAdjustmentCancelled("Proceso cancelado por el usuario durante la revisión del bloque.")

            saved = self._was_saved_after(runtime_word_path, baseline_mtime_ns)
            log(f"Bloque {block_index}/{total_blocks}: guardado detectado={'sí' if saved else 'no'}.")
            log(f"Bloque {block_index}/{total_blocks}: se continúa al siguiente bloque.")

            progress(int((block_index / total_blocks) * 95))

    def _print_blocks(
        self,
        word_path: Path,
        blocks: list[list[AssetRecord]],
        output_dir: Path,
        printer_name: str,
        log: LogCallback,
        status: StatusCallback,
        progress: ProgressCallback,
        manual_adjust: ManualAdjustCallback | None,
    ) -> None:
        total_blocks = len(blocks)
        total_labels = sum(len(block) for block in blocks)
        processed_labels = 0
        for block_index, block in enumerate(blocks, start=1):
            block_word_path = output_dir / self._block_document_name(word_path.stem, block_index)
            try:
                status(f"Generando bloque {block_index} de {total_blocks}")
                log(f"Inicio de generación por bloques: bloque {block_index}/{total_blocks}, cantidad {len(block)}. Copia: {block_word_path}")

                image_dir = output_dir / TEMP_IMAGES_DIR_NAME / f"bloque_{block_index:03d}"
                log(f"Bloque {block_index}/{total_blocks}: creando Word desde plantilla {word_path}.")
                self.word_service.create_from_template(str(word_path), visible=False)
                log(f"Bloque {block_index}/{total_blocks}: Word creado correctamente desde plantilla.")
                self.word_service.prepare_placeholder_document()
                log(f"Bloque {block_index}/{total_blocks}: plantilla preparada para reemplazar <img1>...<img27>.")
                log("Plantilla Word real preparada sobre copia del bloque; se reemplazan placeholders <img1>...<img27>. No hay impresión automática.")
                log(f"Layout visual Word: {self.word_service.get_label_layout_details()}.")
                self._write_and_insert_block(
                    block,
                    block_index,
                    total_blocks,
                    image_dir,
                    log,
                )
                log(f"Bloque {block_index}/{total_blocks}: imágenes insertadas y validadas.")
                self.word_service.save_document_copy(str(block_word_path))
                log(f"Bloque {block_index}/{total_blocks}: documento guardado en {block_word_path}.")
                self.word_service.print_document(printer_name)
                log(f"Bloque {block_index}/{total_blocks}: enviado a impresión en {printer_name} y guardado en {block_word_path}.")
            except Exception:
                self.logger.exception("Error generando bloque real %s/%s.", block_index, total_blocks)
                raise
            finally:
                self.word_service.close(save_changes=False)

            log(f"Bloque {block_index}/{total_blocks}: se continúa al siguiente bloque.")

            processed_labels += len(block)
            progress(int((processed_labels / total_labels) * 100))
        log(f"Fin de generación por bloques. Total bloques: {total_blocks}; etiquetas: {total_labels}.")

    def _write_and_insert_block(
        self,
        block: list[AssetRecord],
        block_index: int,
        total_blocks: int,
        image_dir: Path,
        log: LogCallback,
        manual_adjust: ManualAdjustCallback | None = None,
        printer_name: str | None = None,
        status: StatusCallback | None = None,
        progress: ProgressCallback | None = None,
        progress_base: int = 0,
        total_labels: int | None = None,
    ) -> None:
        self.excel_service.write_block_to_label_sheet(block)
        generated_assets = self.excel_service.get_generated_assets(len(block))
        source_assets = [record.asset_id for record in block]
        self.validation_service.validate_assets_match(source_assets, generated_assets)

        image_dir.mkdir(parents=True, exist_ok=True)
        log(
            f"Bloque {block_index}/{total_blocks}: se usarán placeholders <img1>...<img{len(block)}> "
            "y luego se limpiarán los placeholders restantes hasta <img27>."
        )

        image_paths: list[Path] = []
        last_inserted_page: int | None = None
        for position, record in enumerate(block):
            image_path = image_dir / f"label_{position + 1:03d}.png"
            exported = self.excel_service.export_label_shape_image(
                position,
                image_path,
                target_px=(LABEL_IMAGE_WIDTH_PX, LABEL_IMAGE_HEIGHT_PX),
            )
            image_paths.append(Path(exported.output_path))
            log(
                f"Bloque {block_index}/{total_blocks}: etiqueta {position + 1}/{len(block)} exportada desde "
                f"{exported.group_name}; tamaño final {exported.target_size_px[0]}x{exported.target_size_px[1]} px; "
                f"archivo {exported.output_path}."
            )
            placeholder_number = position + 1
            visual = self.word_service.replace_image_placeholder(placeholder_number, exported.output_path)
            last_inserted_page = visual.page_number
            log(
                f"Bloque {block_index}/{total_blocks}: etiqueta {position + 1}/{len(block)} "
                f"insertada en placeholder {visual.slot_name} (página {visual.page_number}): activo {record.asset_id}."
            )
            log(
                "Validación visual etiqueta "
                f"{position + 1} del bloque; placeholder {visual.slot_name}; "
                f"esperado {visual.expected_width_cm:.2f}x{visual.expected_height_cm:.2f} cm; "
                f"aplicado {visual.applied_width_cm:.2f}x{visual.applied_height_cm:.2f} cm; "
                f"ajuste={'sí' if visual.adjusted else 'no'} ({visual.details})."
            )
            if progress is not None and total_labels:
                progress(int(((progress_base + position + 1) / total_labels) * 100))

        cleaned_placeholders = self.word_service.clear_unused_image_placeholders(len(block))
        if cleaned_placeholders:
            log(
                f"Bloque {block_index}/{total_blocks}: placeholders no usados limpiados: "
                f"{', '.join(cleaned_placeholders)}."
            )
        else:
            log(f"Bloque {block_index}/{total_blocks}: no quedaron placeholders no usados por limpiar.")

        cleanup = self.word_service.cleanup_blank_pages(
            expected_last_image_page=last_inserted_page,
            validation_passes=3,
            cleaned_placeholders=cleaned_placeholders,
        )
        log(
            f"Bloque {block_index}/{total_blocks}: limpieza final Word: "
            f"páginas antes={cleanup.pages_before}, después={cleanup.pages_after}, "
            f"eliminadas={cleanup.removed_pages}, pasadas={cleanup.validation_passes}, "
            f"última página con imagen válida={cleanup.expected_last_image_page or 'n/d'}."
        )
        if cleanup.removed_page_numbers:
            log(
                f"Bloque {block_index}/{total_blocks}: páginas vacías/sobrantes eliminadas: "
                f"{', '.join(str(page) for page in cleanup.removed_page_numbers)}."
            )
        if cleanup.remaining_placeholders:
            raise ValidationError(
                "La limpieza final Word no pudo eliminar placeholders visibles: "
                + ", ".join(cleanup.remaining_placeholders)
            )

        validation = self.word_service.validate_embedded_image_count(len(block))
        log(
            f"Bloque {block_index}/{total_blocks}: validación conteo imágenes Word: "
            f"esperadas={validation.expected_count}, detectadas={validation.detected_count}."
        )
        if validation.extra_count:
            raise ValidationError(
                f"El Word del bloque {block_index}/{total_blocks} quedó con imágenes de más: "
                f"esperadas={validation.expected_count}, detectadas={validation.detected_count}. "
                "No se abre al usuario hasta corregir la plantilla/copia."
            )
        if validation.missing_count:
            log(
                f"Bloque {block_index}/{total_blocks}: faltan {validation.missing_count} imagen(es); "
                f"se repondrán posiciones inferidas {', '.join(str(pos) for pos in validation.missing_positions)} "
                "en página(s) nueva(s)."
            )
            repaired_positions = self._repair_missing_block_images(
                block,
                block_index,
                total_blocks,
                image_dir,
                image_paths,
                validation.missing_positions,
                log,
            )
            repaired_validation = self.word_service.validate_embedded_image_count(len(block))
            log(
                f"Bloque {block_index}/{total_blocks}: imágenes repuestas: "
                f"{', '.join(str(pos) for pos in repaired_positions)}; "
                f"conteo final detectado={repaired_validation.detected_count}."
            )
            if repaired_validation.detected_count != repaired_validation.expected_count:
                raise ValidationError(
                    f"La reparación de imágenes del bloque {block_index}/{total_blocks} no alcanzó el conteo exacto: "
                    f"esperadas={repaired_validation.expected_count}, detectadas={repaired_validation.detected_count}. "
                    "No se abre el Word porque el bloque no está completo."
                )
            validation = repaired_validation

        log(
            f"Bloque {block_index}/{total_blocks}: validación final OK; "
            f"no quedan placeholders <img...> visibles y el conteo de imágenes es exacto "
            f"({validation.detected_count}/{validation.expected_count})."
        )

    def _repair_missing_block_images(
        self,
        block: list[AssetRecord],
        block_index: int,
        total_blocks: int,
        image_dir: Path,
        image_paths: list[Path],
        missing_positions: list[int],
        log: LogCallback,
    ) -> list[int]:
        repaired_positions: list[int] = []
        for placeholder_number in missing_positions:
            if placeholder_number < 1 or placeholder_number > len(block):
                raise ValidationError(
                    f"No se puede reponer la imagen {placeholder_number} del bloque {block_index}/{total_blocks}: "
                    f"el bloque tiene {len(block)} registro(s)."
                )

            image_path = image_paths[placeholder_number - 1] if placeholder_number <= len(image_paths) else image_dir / f"label_{placeholder_number:03d}.png"
            if not image_path.exists():
                exported = self.excel_service.export_label_shape_image(
                    placeholder_number - 1,
                    image_path,
                    target_px=(LABEL_IMAGE_WIDTH_PX, LABEL_IMAGE_HEIGHT_PX),
                )
                image_path = Path(exported.output_path)
                log(
                    f"Bloque {block_index}/{total_blocks}: PNG para reposición de etiqueta {placeholder_number} "
                    f"regenerado desde {exported.group_name}: {image_path}."
                )

            record = block[placeholder_number - 1]
            visual = self.word_service.append_label_image_on_new_page(placeholder_number, image_path)
            repaired_positions.append(placeholder_number)
            log(
                f"Bloque {block_index}/{total_blocks}: etiqueta {placeholder_number}/{len(block)} "
                f"repuesta en página nueva {visual.page_number}: activo {record.asset_id}; archivo {image_path}."
            )
        return repaired_positions

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
