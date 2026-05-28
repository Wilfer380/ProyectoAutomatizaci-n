import unittest
from PySide6.QtCore import QObject
from view_models.preview_view_model import PreviewViewModel
from view_models.label_item_view_model import LabelItemViewModel


class TestPreviewViewModel(unittest.TestCase):
    def test_initialization(self):
        vm = PreviewViewModel()
        self.assertIsInstance(vm, QObject)
        self.assertEqual(vm.label_items, [])
        self.assertEqual(vm.current_page_index, 0)

    def test_set_items_emits_preview_ready(self):
        vm = PreviewViewModel()
        items = [LabelItemViewModel(), LabelItemViewModel()]

        signal_emitted = False

        def on_preview_ready():
            nonlocal signal_emitted
            signal_emitted = True

        vm.previewReady.connect(on_preview_ready)

        vm.set_items(items)
        self.assertTrue(signal_emitted)
        self.assertEqual(len(vm.label_items), 2)
        self.assertEqual(vm.current_page_index, 0)

    def test_navigation(self):
        vm = PreviewViewModel()
        items = [LabelItemViewModel(), LabelItemViewModel(), LabelItemViewModel()]
        vm.set_items(items)

        self.assertEqual(vm.current_page_index, 0)
        vm.next_page()
        self.assertEqual(vm.current_page_index, 1)
        vm.next_page()
        self.assertEqual(vm.current_page_index, 2)
        vm.next_page()  # Should not go out of bounds
        self.assertEqual(vm.current_page_index, 2)

        vm.previous_page()
        self.assertEqual(vm.current_page_index, 1)
        vm.previous_page()
        self.assertEqual(vm.current_page_index, 0)
        vm.previous_page()  # Should not go out of bounds
        self.assertEqual(vm.current_page_index, 0)

    def test_confirm_emits_print_signals_and_prints_current_items(self):
        printed_items = []
        vm = PreviewViewModel(print_callback=lambda items: printed_items.extend(items))
        items = [LabelItemViewModel(), LabelItemViewModel()]
        vm.set_items(items)

        started_emitted = False

        def on_started():
            nonlocal started_emitted
            started_emitted = True

        vm.printStarted.connect(on_started)

        completed_emitted = False

        def on_completed():
            nonlocal completed_emitted
            completed_emitted = True

        vm.printCompleted.connect(on_completed)

        vm.confirm()

        self.assertTrue(started_emitted)
        self.assertTrue(completed_emitted)
        self.assertEqual(printed_items, items)

    def test_confirm_emits_print_failed_when_callback_fails(self):
        vm = PreviewViewModel(
            print_callback=lambda _items: (_ for _ in ()).throw(
                RuntimeError("driver missing")
            )
        )
        failures = []
        completed = []
        vm.printFailed.connect(lambda message: failures.append(message))
        vm.printCompleted.connect(lambda: completed.append(True))

        vm.confirm()

        self.assertEqual(failures, ["driver missing"])
        self.assertEqual(completed, [])

    def test_redo_emits_signal(self):
        vm = PreviewViewModel()

        redo_emitted = False

        def on_redo():
            nonlocal redo_emitted
            redo_emitted = True

        vm.redoRequested.connect(on_redo)

        vm.redo()
        self.assertTrue(redo_emitted)


if __name__ == "__main__":
    unittest.main()
