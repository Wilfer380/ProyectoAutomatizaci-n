from __future__ import annotations

import inspect
import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from services.process_service import ProcessService


class ProcessServiceWordComTests(unittest.TestCase):
    def test_process_service_uses_word_com_generation_flow(self) -> None:
        source = inspect.getsource(ProcessService)
        for token in (
            "shutil.copy2(word_path, block_word_path)",
            "self.word_service.open(str(block_word_path), visible=False)",
            "replace_image_placeholder(",
            "clear_unused_image_placeholders(",
            "cleanup_blank_pages(",
            "save_document_copy(",
            "print_document(",
            "show_to_user(",
            "release_to_user(",
        ):
            self.assertIn(token, source)


if __name__ == "__main__":
    unittest.main()
