from __future__ import annotations

import inspect
import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from services.process_service import ProcessService


class ProcessServiceDocumentPipelineTests(unittest.TestCase):
    def test_process_service_uses_headless_word_generation_flow(self) -> None:
        source = inspect.getsource(ProcessService)

        for token in (
            "self._plan_blocks(records)",
            "self._build_word_document(",
            "self.word_service.build_document_from_template(",
            "self.word_service.build_document_without_com(",
            "validate_embedded_image_count(",
            "manual_adjust(",
        ):
            self.assertIn(token, source)

        for token in (
            "print_document(",
            "set_default_printer(",
        ):
            self.assertNotIn(token, source)


if __name__ == "__main__":
    unittest.main()
