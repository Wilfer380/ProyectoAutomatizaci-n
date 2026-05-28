from PySide6.QtCore import QObject, Signal

class LabelItemViewModel(QObject):
    layoutChanged = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.asset_id = ""
        self.asset_name = ""
        self.section = ""
        self.image_data = None
        self.image_offset_x = 0
        self.image_offset_y = 0
        self.image_scale = 1.0

    def set_image_offset_x(self, value):
        if self.image_offset_x != value:
            self.image_offset_x = value
            self.layoutChanged.emit()

    def set_image_offset_y(self, value):
        if self.image_offset_y != value:
            self.image_offset_y = value
            self.layoutChanged.emit()

    def set_image_scale(self, value):
        if self.image_scale != value:
            self.image_scale = value
            self.layoutChanged.emit()
