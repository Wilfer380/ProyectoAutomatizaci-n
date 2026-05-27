from __future__ import annotations

from pathlib import Path
from typing import Iterable

from PySide6.QtCore import QPoint, Qt, QStandardPaths, Signal
from PySide6.QtWidgets import (
    QComboBox,
    QFileDialog,
    QFrame,
    QGroupBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QStatusBar,
    QVBoxLayout,
    QWidget,
)

from models.app_settings import AppSettings
from utils.constants import APP_NAME, TARGET_PRINTER_NAME
from utils.runtime import get_user_home_dir
from view_models.main_view_model import MainViewModel


class FilterComboBox(QComboBox):
    def showPopup(self) -> None:
        super().showPopup()
        popup = self.view().window()
        popup.setMinimumWidth(self.width())
        popup.move(self.mapToGlobal(QPoint(0, self.height())))


class MainWindow(QMainWindow):
    select_excel_requested = Signal()
    select_word_requested = Signal()
    excel_path_changed = Signal(str)
    word_path_changed = Signal(str)
    refresh_filters_requested = Signal()
    start_process_requested = Signal()
    start_simulation_requested = Signal()
    cancel_process_requested = Signal()
    printer_config_requested = Signal()

    def __init__(self, settings: AppSettings, view_model: MainViewModel | None = None) -> None:
        super().__init__()
        self._settings = settings
        self.view_model = view_model
        self._build_ui()
        self._apply_settings()
        if self.view_model is not None:
            self._bind_view_model(self.view_model)

    def _build_ui(self) -> None:
        self.setWindowTitle(APP_NAME)
        self.resize(1120, 780)
        self.setMinimumSize(840, 620)

        status_bar = QStatusBar(self)
        status_bar.setSizeGripEnabled(False)
        status_bar.showMessage("Aplicación lista.")
        self.setStatusBar(status_bar)

        scroll_area = QScrollArea(self)
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.NoFrame)
        self.setCentralWidget(scroll_area)

        central_widget = QWidget()
        scroll_area.setWidget(central_widget)

        root_layout = QVBoxLayout(central_widget)
        root_layout.setContentsMargins(24, 20, 24, 20)
        root_layout.setSpacing(18)

        root_layout.addWidget(self._build_header_card())
        root_layout.addWidget(self._build_summary_bar())
        root_layout.addWidget(self._build_file_group())
        root_layout.addWidget(self._build_process_group())
        root_layout.addWidget(self._build_progress_group())
        root_layout.addStretch(1)

        self.log_panel = self._build_logs_group()
        self.log_panel.setVisible(False)
        root_layout.addWidget(self.log_panel)

        self._apply_styles()

    def _build_header_card(self) -> QFrame:
        frame = QFrame()
        frame.setObjectName("headerCard")
        layout = QVBoxLayout(frame)
        layout.setContentsMargins(24, 22, 24, 22)
        layout.setSpacing(8)

        title_label = QLabel(APP_NAME)
        title_label.setObjectName("titleLabel")

        subtitle_label = QLabel("Automatización de generación y previsualización de etiquetas desde Excel")
        subtitle_label.setObjectName("subtitleLabel")
        subtitle_label.setWordWrap(True)

        helper_label = QLabel(
            "Seleccioná la base Excel. "
            "El sistema genera una previsualización nativa de etiquetas 48x23 mm, "
            "permite confirmar o rehacer ajustes y evita depender de plantillas Word."
        )
        helper_label.setObjectName("helperLabel")
        helper_label.setWordWrap(True)

        layout.addWidget(title_label)
        layout.addWidget(subtitle_label)
        layout.addWidget(helper_label)
        return frame

    def _build_summary_bar(self) -> QFrame:
        frame = QFrame()
        frame.setObjectName("summaryBar")
        layout = QHBoxLayout(frame)
        layout.setContentsMargins(18, 14, 18, 14)
        layout.setSpacing(12)

        self.excel_state_chip = self._create_chip("Excel pendiente")
        self.printer_state_chip = self._create_chip("Impresora pendiente")

        layout.addWidget(self.excel_state_chip)
        layout.addWidget(self.printer_state_chip)
        layout.addStretch(1)
        return frame

    def _create_chip(self, text: str) -> QLabel:
        chip = QLabel(text)
        chip.setObjectName("statusChip")
        chip.setProperty("chipState", "warning")
        chip.setAlignment(Qt.AlignCenter)
        chip.setMinimumHeight(32)
        return chip

    def _build_file_group(self) -> QGroupBox:
        group = QGroupBox("1. Archivos de trabajo")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(18, 20, 18, 18)
        layout.setSpacing(16)

        layout.addWidget(
            self._build_labeled_field(
                "Archivo Excel",
                "Seleccione la base de datos Excel",
                "Se usa como origen. Se leerá Hoja1 y se trabajará con Etiqueta provisional.",
                is_excel=True,
            )
        )

        return group

    def _build_labeled_field(
        self,
        title: str,
        placeholder: str,
        hint: str,
        *,
        is_excel: bool,
    ) -> QWidget:
        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        title_label = QLabel(title)
        title_label.setObjectName("fieldTitle")

        row_layout = QHBoxLayout()
        row_layout.setContentsMargins(0, 0, 0, 0)
        row_layout.setSpacing(12)

        line_edit = QLineEdit()
        line_edit.setReadOnly(False)
        line_edit.setClearButtonEnabled(True)
        line_edit.setPlaceholderText(placeholder)
        line_edit.setMinimumHeight(44)

        button = QPushButton("Seleccionar Excel" if is_excel else "Seleccionar plantilla Word")
        button.setObjectName("secondaryButton")
        button.setMinimumWidth(180)
        button.setMinimumHeight(44)

        hint_label = QLabel(hint)
        hint_label.setObjectName("fieldHint")
        hint_label.setWordWrap(True)

        row_layout.addWidget(line_edit, stretch=1)
        row_layout.addWidget(button)

        layout.addWidget(title_label)
        layout.addLayout(row_layout)
        layout.addWidget(hint_label)

        if is_excel:
            self.excel_path_edit = line_edit
            self.select_excel_button = button
            self.excel_hint_label = hint_label
            self.select_excel_button.clicked.connect(self.select_excel_requested.emit)
            self.excel_path_edit.editingFinished.connect(
                lambda: self.excel_path_changed.emit(self.excel_path_edit.text().strip())
            )
        else:
            self.word_path_edit = line_edit
            self.select_word_button = button
            self.word_hint_label = hint_label
            self.select_word_button.clicked.connect(self.select_word_requested.emit)
            self.word_path_edit.editingFinished.connect(
                lambda: self.word_path_changed.emit(self.word_path_edit.text().strip())
            )

        return container

    def _build_process_group(self) -> QGroupBox:
        group = QGroupBox("2. Configuración del proceso")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(18, 20, 18, 18)
        layout.setSpacing(16)

        filter_container = QWidget()
        filter_layout = QVBoxLayout(filter_container)
        filter_layout.setContentsMargins(0, 0, 0, 0)
        filter_layout.setSpacing(6)

        filter_title = QLabel("Filtro")
        filter_title.setObjectName("fieldTitle")
        filter_row = QHBoxLayout()
        filter_row.setContentsMargins(0, 0, 0, 0)
        filter_row.setSpacing(12)

        self.filter_combo = FilterComboBox()
        self.filter_combo.setPlaceholderText("Seleccione un filtro")
        self.filter_combo.setMinimumHeight(44)
        self.filter_combo.setMaxVisibleItems(20)
        self.filter_combo.view().setTextElideMode(Qt.ElideNone)

        self.refresh_filters_button = QPushButton("Actualizar filtros")
        self.refresh_filters_button.setObjectName("secondaryButton")
        self.refresh_filters_button.setMinimumWidth(150)
        self.refresh_filters_button.setMinimumHeight(44)
        self.refresh_filters_button.clicked.connect(self.refresh_filters_requested.emit)

        filter_row.addWidget(self.filter_combo, stretch=1)
        filter_row.addWidget(self.refresh_filters_button)

        filter_layout.addWidget(filter_title)
        filter_layout.addLayout(filter_row)

        printer_container = QWidget()
        printer_layout = QVBoxLayout(printer_container)
        printer_layout.setContentsMargins(0, 0, 0, 0)
        printer_layout.setSpacing(6)

        printer_title = QLabel("Impresora")
        printer_title.setObjectName("fieldTitle")
        printer_row = QHBoxLayout()
        printer_row.setContentsMargins(0, 0, 0, 0)
        printer_row.setSpacing(12)

        self.printer_name_edit = QLineEdit()
        self.printer_name_edit.setReadOnly(True)
        self.printer_name_edit.setMinimumHeight(44)

        self.configure_printer_button = QPushButton("Validar impresora")
        self.configure_printer_button.setObjectName("secondaryButton")
        self.configure_printer_button.setMinimumWidth(150)
        self.configure_printer_button.setMinimumHeight(44)
        self.configure_printer_button.clicked.connect(self.printer_config_requested.emit)

        printer_row.addWidget(self.printer_name_edit, stretch=1)
        printer_row.addWidget(self.configure_printer_button)

        printer_layout.addWidget(printer_title)
        printer_layout.addLayout(printer_row)

        self.start_button = QPushButton("Generar etiquetas")
        self.start_button.setObjectName("primaryButton")
        self.start_button.setMinimumHeight(48)
        self.start_button.setEnabled(False)
        self.start_button.clicked.connect(self.start_process_requested.emit)

        self.cancel_process_button = QPushButton("Cancelar proceso")
        self.cancel_process_button.setObjectName("secondaryButton")
        self.cancel_process_button.setMinimumHeight(44)
        self.cancel_process_button.setVisible(False)
        self.cancel_process_button.setEnabled(False)
        self.cancel_process_button.clicked.connect(self.cancel_process_requested.emit)

        self.simulate_button = QPushButton("Prueba visual sin imprimir")
        self.simulate_button.setObjectName("secondaryButton")
        self.simulate_button.setMinimumHeight(44)
        self.simulate_button.setVisible(False)
        self.simulate_button.clicked.connect(self.start_simulation_requested.emit)

        layout.addWidget(filter_container)
        layout.addWidget(printer_container)
        layout.addWidget(self.simulate_button)
        layout.addWidget(self.start_button)
        layout.addWidget(self.cancel_process_button)
        return group

    def _build_progress_group(self) -> QGroupBox:
        group = QGroupBox("3. Estado del proceso")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(18, 20, 18, 18)
        layout.setSpacing(12)

        info_layout = QHBoxLayout()
        self.status_label = QLabel("Listo para iniciar.")
        self.status_label.setObjectName("statusTitle")
        self.records_label = QLabel("Registros detectados: 0")
        self.records_label.setObjectName("statusMeta")
        self.records_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

        info_layout.addWidget(self.status_label, stretch=1)
        info_layout.addWidget(self.records_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("%p% completado")
        self.progress_bar.setMinimumHeight(30)

        layout.addLayout(info_layout)
        layout.addWidget(self.progress_bar)
        return group

    def _build_logs_group(self) -> QGroupBox:
        group = QGroupBox("Logs del proceso")
        layout = QVBoxLayout(group)

        self.log_output = QPlainTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        layout.addWidget(self.log_output)
        return group

    def _apply_styles(self) -> None:
        self.setStyleSheet(
            """
            QWidget {
                background-color: #0B1120;
                color: #E2E8F0;
                font-size: 14px;
                font-family: "Segoe UI", "Arial", sans-serif;
            }
            QMainWindow {
                background-color: #0B1120;
            }
            QStatusBar {
                background: #111827;
                color: #CBD5E1;
                border-top: 1px solid #1F2937;
                min-height: 28px;
            }
            QFrame#headerCard, QFrame#summaryBar {
                background-color: #111827;
                border: 1px solid #1F2937;
                border-radius: 16px;
            }
            #titleLabel {
                color: #F8FAFC;
                font-size: 30px;
                font-weight: 700;
            }
            #subtitleLabel {
                color: #93C5FD;
                font-size: 14px;
                font-weight: 600;
            }
            #helperLabel {
                color: #CBD5E1;
                font-size: 13px;
                padding-top: 2px;
            }
            #statusChip {
                padding: 4px 12px;
                border-radius: 14px;
                font-size: 12px;
                font-weight: 700;
            }
            QLabel[chipState="warning"] {
                background: #FEF3C7;
                color: #92400E;
                border: 1px solid #FCD34D;
            }
            QLabel[chipState="ready"] {
                background: #DCFCE7;
                color: #166534;
                border: 1px solid #86EFAC;
            }
            QGroupBox {
                font-weight: 700;
                border: 1px solid #1F2937;
                border-radius: 16px;
                margin-top: 12px;
                padding: 18px 16px 16px 16px;
                background-color: #111827;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 14px;
                padding: 0 8px 0 8px;
                color: #F8FAFC;
                background-color: #111827;
            }
            QLabel {
                color: #E2E8F0;
                background: transparent;
            }
            QLabel#fieldTitle {
                color: #F8FAFC;
                font-size: 14px;
                font-weight: 700;
            }
            QLabel#fieldHint {
                color: #94A3B8;
                font-size: 12px;
            }
            QLabel#statusTitle {
                color: #F8FAFC;
                font-size: 15px;
                font-weight: 700;
            }
            QLabel#statusMeta {
                color: #93C5FD;
                font-size: 13px;
                font-weight: 600;
            }
            QLineEdit, QComboBox, QPlainTextEdit {
                border: 1px solid #334155;
                border-radius: 10px;
                padding: 10px 12px;
                background: #F8FAFC;
                color: #0F172A;
                selection-background-color: #2563EB;
                selection-color: white;
            }
            QLineEdit {
                background: #FFFFFF;
            }
            QLineEdit[readOnly="true"] {
                background: #F1F5F9;
            }
            QComboBox {
                padding-right: 34px;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 30px;
                border-left: 1px solid #CBD5E1;
                background: #EEF2F7;
                border-top-right-radius: 10px;
                border-bottom-right-radius: 10px;
            }
            QComboBox::down-arrow {
                image: none;
                width: 0px;
                height: 0px;
                border-left: 6px solid transparent;
                border-right: 6px solid transparent;
                border-top: 8px solid #334155;
                margin-right: 10px;
            }
            QComboBox QAbstractItemView {
                background: #FFFFFF;
                color: #0F172A;
                border: 1px solid #CBD5E1;
                selection-background-color: #DBEAFE;
                selection-color: #0F172A;
                outline: 0;
                min-width: 280px;
            }
            QPushButton {
                min-height: 42px;
                border-radius: 10px;
                padding: 9px 16px;
                font-weight: 600;
            }
            QPushButton#secondaryButton {
                background: #E2E8F0;
                color: #0F172A;
                border: 1px solid #CBD5E1;
            }
            QPushButton#secondaryButton:hover {
                background: #CBD5E1;
            }
            QPushButton#primaryButton {
                background: #2563EB;
                color: white;
                font-weight: 700;
                font-size: 15px;
                border: 1px solid #2563EB;
            }
            QPushButton#primaryButton:hover {
                background: #1D4ED8;
            }
            QPushButton:disabled {
                background: #1F2937;
                color: #94A3B8;
                border: 1px solid #334155;
            }
            QProgressBar {
                border: 1px solid #334155;
                border-radius: 10px;
                text-align: center;
                background-color: #0F172A;
                color: #F8FAFC;
                font-weight: 700;
            }
            QProgressBar::chunk {
                background-color: #2563EB;
                border-radius: 8px;
            }
            QPlainTextEdit {
                background: #020617;
                color: #E2E8F0;
                border: 1px solid #334155;
                padding: 10px;
            }
            QPlainTextEdit:focus, QLineEdit:focus, QComboBox:focus {
                border: 1px solid #60A5FA;
            }
            """
        )

    def _apply_settings(self) -> None:
        self.set_excel_path("")
        self.set_printer_name(self._settings.printer_name or TARGET_PRINTER_NAME)
        self.set_selected_filter(self._settings.selected_filter)

    def _bind_view_model(self, view_model: MainViewModel) -> None:
        view_model.fileSelected.connect(self.set_excel_path)
        view_model.progressChanged.connect(self.set_progress)
        view_model.processingStarted.connect(lambda: self.set_busy(True))
        view_model.processingFinished.connect(self._open_preview_subwindow)
        view_model.errorOccurred.connect(lambda message: self.show_error("Error procesando Excel", message))
        self.start_button.clicked.connect(view_model.process_file)
        self.start_button.setEnabled(True)

    def _open_preview_subwindow(self, _label_items: list[object]) -> None:
        self.set_busy(False)
        self.set_status("Etiquetas generadas. Vista previa lista para revisión.")

    def choose_excel_file(self, start_dir: str = "") -> str:
        start_dir = self._resolve_dialog_directory(start_dir, self.excel_path_edit.text())
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar archivo Excel",
            start_dir,
            "Excel (*.xlsx *.xlsm *.xls)",
        )
        return path

    def choose_word_file(self, start_dir: str = "") -> str:
        start_dir = self._resolve_dialog_directory(start_dir, self.word_path_edit.text())
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleccionar plantilla Word",
            start_dir,
            "Word (*.docx *.docm *.doc)",
        )
        return path

    def _resolve_dialog_directory(self, preferred_dir: str, current_path: str) -> str:
        if current_path:
            current = Path(current_path)
            if current.exists():
                return str(current.parent)

        if preferred_dir:
            preferred = Path(preferred_dir)
            if preferred.exists():
                return str(preferred)

        downloads = QStandardPaths.writableLocation(QStandardPaths.DownloadLocation)
        if downloads:
            return downloads

        documents = QStandardPaths.writableLocation(QStandardPaths.DocumentsLocation)
        if documents:
            return documents

        return str(get_user_home_dir())

    def set_excel_path(self, path: str) -> None:
        self.excel_path_edit.setText(path)
        self._set_chip_state(self.excel_state_chip, bool(path), "Excel cargado", "Excel pendiente")

    def set_word_path(self, path: str) -> None:
        self.word_path_edit.setText(path)
        self._set_chip_state(self.word_state_chip, bool(path), "Word cargado", "Word pendiente")

    def set_printer_name(self, printer_name: str) -> None:
        self.printer_name_edit.setText(printer_name)
        self._set_chip_state(
            self.printer_state_chip,
            bool(printer_name.strip()),
            f"Impresora: {printer_name}",
            "Impresora pendiente",
        )

    def _set_chip_state(self, chip: QLabel, ready: bool, ready_text: str, pending_text: str) -> None:
        chip.setText(ready_text if ready else pending_text)
        chip.setProperty("chipState", "ready" if ready else "warning")
        self.style().unpolish(chip)
        self.style().polish(chip)
        chip.update()

    def set_filters(self, filters: Iterable[str]) -> None:
        current = self.filter_combo.currentText()
        self.filter_combo.blockSignals(True)
        self.filter_combo.clear()
        self.filter_combo.addItems(list(filters))
        self.filter_combo.blockSignals(False)
        self.set_selected_filter(current)

    def selected_filter(self) -> str:
        return self.filter_combo.currentText().strip()

    def set_selected_filter(self, value: str) -> None:
        if not value:
            return

        index = self.filter_combo.findText(value)
        if index >= 0:
            self.filter_combo.setCurrentIndex(index)

    def set_start_enabled(self, enabled: bool) -> None:
        self.start_button.setEnabled(enabled)

    def set_status(self, message: str) -> None:
        self.status_label.setText(message)
        self.statusBar().showMessage(message)

    def set_record_count(self, count: int) -> None:
        self.records_label.setText(f"Registros detectados: {count}")

    def set_progress(self, value: int) -> None:
        self.progress_bar.setValue(value)

    def append_log(self, message: str) -> None:
        self.log_output.appendPlainText(message)
        self.statusBar().showMessage(message)

    def clear_logs(self) -> None:
        self.log_output.clear()
        self.log_panel.setVisible(False)

    def reset_after_process_finished(self, message: str) -> None:
        self.setUpdatesEnabled(False)
        try:
            self.set_busy(False)
            self.set_manual_review_mode(False)
            self.start_button.setText("Generar e imprimir")
            self.start_button.setEnabled(True)
            self.cancel_process_button.setEnabled(False)
            self.cancel_process_button.setVisible(False)
            self.simulate_button.setEnabled(False)
            self.simulate_button.setVisible(False)
            self.clear_logs()
            self.set_progress(0)
            self.set_status(message)
        finally:
            self.setUpdatesEnabled(True)

    def set_busy(self, busy: bool) -> None:
        self.select_excel_button.setEnabled(not busy)
        self.refresh_filters_button.setEnabled(not busy)
        self.configure_printer_button.setEnabled(not busy)
        self.simulate_button.setEnabled(False)
        self.simulate_button.setVisible(False)
        self.start_button.setEnabled(not busy)

    def set_manual_review_mode(
        self,
        active: bool,
        block_index: int | None = None,
        total_blocks: int | None = None,
        document_path: str = "",
    ) -> None:
        if active and block_index is not None and total_blocks is not None:
            self.start_button.setText("Continuar")
            self.start_button.setEnabled(True)
            self.simulate_button.setEnabled(False)
            self.simulate_button.setVisible(False)
            self.cancel_process_button.setVisible(True)
            self.cancel_process_button.setEnabled(True)
            self.set_status(
                f"Bloque {block_index}/{total_blocks} listo. "
                "Revisá/imprimí/guardá en Word y luego presioná Continuar."
            )
            self.append_log(
                f"Esperando revisión manual del bloque {block_index}/{total_blocks}. Archivo Word: {document_path}"
            )
            return

        self.start_button.setText("Generar e imprimir")
        self.cancel_process_button.setEnabled(False)
        self.cancel_process_button.setVisible(False)
        self.simulate_button.setEnabled(False)
        self.simulate_button.setVisible(False)

    def show_error(self, title: str, message: str) -> None:
        friendly_message = message.strip()
        if title and title != "Error en la aplicación":
            friendly_message = f"{title}: {friendly_message}"
        if "carpeta logs" not in friendly_message.lower():
            friendly_message = f"{friendly_message}\n\nEl detalle técnico fue guardado en la carpeta logs."
        QMessageBox.critical(self, "Error en la aplicación", friendly_message)

    def show_info(self, title: str, message: str) -> None:
        QMessageBox.information(self, title, message)

    def confirm_continue_without_save(self, document_path: str) -> bool:
        confirm = QMessageBox.warning(
            self,
            "Guardado no detectado",
            f"No detecté un guardado/modificación posterior en el bloque:\n{document_path}\n\n"
            "¿Querés continuar de todos modos?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        return confirm == QMessageBox.Yes

    def file_was_saved_after(self, document_path: str, baseline_mtime_ns: int) -> bool:
        try:
            return Path(document_path).stat().st_mtime_ns > baseline_mtime_ns
        except OSError:
            return False
