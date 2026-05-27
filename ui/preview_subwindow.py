from PySide6.QtWidgets import (
    QDialog,
    QGraphicsScene,
    QGraphicsView,
    QHBoxLayout,
    QPushButton,
    QVBoxLayout,
)

from view_models.preview_view_model import PreviewViewModel


class PreviewSubwindow(QDialog):
    def __init__(self, view_model: PreviewViewModel):
        super().__init__()
        self.view_model = view_model

        self.setWindowTitle("Previsualización de Etiqueta")

        layout = QVBoxLayout(self)

        self.scene = QGraphicsScene(self)
        self.scene.setSceneRect(0, 0, 480, 230)  # 48x23mm at 10 px/mm logical preview scale.
        self.view = QGraphicsView(self.scene)
        self.view.setMinimumSize(520, 270)
        self.scene.addRect(self.scene.sceneRect())
        layout.addWidget(self.view)

        button_layout = QHBoxLayout()
        self.btn_confirmar = QPushButton("Confirmar")
        self.btn_confirmar.setObjectName("btnConfirmar")

        self.btn_rehacer = QPushButton("Rehacer")
        self.btn_rehacer.setObjectName("btnRehacer")

        button_layout.addWidget(self.btn_confirmar)
        button_layout.addWidget(self.btn_rehacer)

        layout.addLayout(button_layout)

        self.btn_confirmar.clicked.connect(self.on_confirmar)
        self.btn_rehacer.clicked.connect(self.on_rehacer)

    def on_confirmar(self):
        self.view_model.confirm()

    def on_rehacer(self):
        self.view_model.redo()
