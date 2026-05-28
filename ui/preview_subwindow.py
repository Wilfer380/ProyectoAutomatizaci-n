from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import (
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from view_models.preview_view_model import PreviewViewModel


class PreviewSubwindow(QDialog):
    def __init__(self, view_model: PreviewViewModel):
        super().__init__()
        self.view_model = view_model

        self.setWindowTitle("Previsualización de impresión")
        self.resize(780, 820)
        self.setMinimumSize(700, 680)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)

        self.summary_label = QLabel(self)
        self.summary_label.setWordWrap(True)
        layout.addWidget(self.summary_label)

        self.scroll_area = QScrollArea(self)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.scroll_area.setObjectName("previewScrollArea")
        self.scroll_area.setStyleSheet(
            "QScrollArea#previewScrollArea { background: white; border: 1px solid #CBD5E1; }"
        )

        self.preview_container = QWidget(self.scroll_area)
        self.preview_container.setObjectName("previewContainer")
        self.preview_layout = QVBoxLayout(self.preview_container)
        self.preview_layout.setContentsMargins(18, 18, 18, 18)
        self.preview_layout.setSpacing(12)
        self.preview_layout.setAlignment(
            Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter
        )
        self.preview_container.setStyleSheet(
            "QWidget#previewContainer { background: white; }"
        )
        self.scroll_area.setWidget(self.preview_container)
        layout.addWidget(self.scroll_area, stretch=1)

        button_layout = QHBoxLayout()
        self.btn_confirmar = QPushButton("Confirmar e imprimir")
        self.btn_confirmar.setObjectName("btnConfirmar")
        self.btn_confirmar.setMinimumHeight(40)

        self.btn_rehacer = QPushButton("Rechazar")
        self.btn_rehacer.setObjectName("btnRehacer")
        self.btn_rehacer.setMinimumHeight(40)

        button_layout.addWidget(self.btn_confirmar)
        button_layout.addWidget(self.btn_rehacer)
        layout.addLayout(button_layout)

        self.btn_confirmar.clicked.connect(self.on_confirmar)
        self.btn_rehacer.clicked.connect(self.on_rehacer)
        self.view_model.previewReady.connect(self.refresh_preview)
        self.view_model.printStarted.connect(self._on_print_started)
        self.view_model.printCompleted.connect(self.accept)
        self.view_model.printFailed.connect(self.show_print_error)
        self.view_model.redoRequested.connect(self.reject)
        self.refresh_preview()

    def refresh_preview(self) -> None:
        self._clear_preview()
        images = self.view_model.preview_images()
        total = len(images)
        self.summary_label.setText(
            f"Revisá cómo quedarán impresas las {total} etiqueta(s). "
            "Usá el scroll para verlas todas antes de confirmar."
        )
        for index, image in enumerate(images, start=1):
            self.preview_layout.addWidget(self._build_label_card(index, image))
        self.preview_layout.addStretch(1)

    def _build_label_card(self, index: int, image) -> QWidget:
        card = QFrame(self.preview_container)
        card.setObjectName("labelPreviewCard")
        card.setFrameShape(QFrame.Shape.StyledPanel)
        card.setStyleSheet(
            "QFrame#labelPreviewCard { background: white; border: 1px solid #E2E8F0; }"
        )
        layout = QVBoxLayout(card)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(6)

        title = QLabel(f"Etiqueta {index}", card)
        title.setAlignment(Qt.AlignmentFlag.AlignLeft)
        layout.addWidget(title)

        image_label = QLabel(card)
        image_label.setObjectName("labelPreviewImage")
        image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        image_label.setStyleSheet("background: white; border: 1px solid #CBD5E1;")
        image_label.setPixmap(QPixmap.fromImage(image))
        layout.addWidget(image_label)
        return card

    def _clear_preview(self) -> None:
        while self.preview_layout.count():
            item = self.preview_layout.takeAt(0)
            if item is not None:
                widget = item.widget()
                if widget is not None:
                    widget.deleteLater()

    def on_confirmar(self):
        self.view_model.confirm()

    def on_rehacer(self):
        self.view_model.redo()

    def _on_print_started(self) -> None:
        self.btn_confirmar.setEnabled(False)
        self.btn_rehacer.setEnabled(False)
        self.summary_label.setText("Enviando etiquetas a impresión…")

    def show_print_error(self, message: str) -> None:
        self.btn_confirmar.setEnabled(True)
        self.btn_rehacer.setEnabled(True)
        self.summary_label.setText(
            "No se pudo imprimir. Revisá el error y volvé a intentar."
        )
        QMessageBox.critical(self, "No se pudo imprimir", message)
