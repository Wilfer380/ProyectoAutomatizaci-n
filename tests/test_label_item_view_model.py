import unittest
from PySide6.QtCore import QObject
from view_models.label_item_view_model import LabelItemViewModel

class TestLabelItemViewModel(unittest.TestCase):
    def test_initialization(self):
        vm = LabelItemViewModel()
        self.assertIsInstance(vm, QObject)
        self.assertEqual(vm.asset_id, "")
        self.assertEqual(vm.asset_name, "")
        self.assertEqual(vm.section, "")
        self.assertIsNone(vm.image_data)
        self.assertEqual(vm.image_offset_x, 0)
        self.assertEqual(vm.image_offset_y, 0)
        self.assertEqual(vm.image_scale, 1.0)

    def test_layout_changed_signal(self):
        vm = LabelItemViewModel()
        
        signal_emitted = False
        def on_layout_changed():
            nonlocal signal_emitted
            signal_emitted = True
            
        vm.layoutChanged.connect(on_layout_changed)
        
        # When layout properties change, the signal should be emitted
        vm.set_image_offset_x(10)
        self.assertTrue(signal_emitted)
        
        signal_emitted = False
        vm.set_image_offset_y(20)
        self.assertTrue(signal_emitted)
        
        signal_emitted = False
        vm.set_image_scale(1.5)
        self.assertTrue(signal_emitted)

if __name__ == '__main__':
    unittest.main()
