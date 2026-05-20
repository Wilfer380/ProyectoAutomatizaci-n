from __future__ import annotations

import inspect
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

from PIL import Image
from docx import Document as DocxDocument


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import services.word_service as word_service_module
from services.word_service import WordService


class WordServiceNoComTests(unittest.TestCase):
    def test_build_document_without_com_replaces_placeholders(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            template_path = temp_path / "template.docx"
            output_path = temp_path / "output.docx"
            image_one = temp_path / "image-1.png"
            image_two = temp_path / "image-2.png"

            self._write_image(image_one, (220, 40, 40))
            self._write_image(image_two, (40, 110, 220))

            template = DocxDocument()
            template.add_paragraph("Body <img1>")
            table = template.add_table(rows=1, cols=1)
            table.cell(0, 0).text = "Table <img2>"
            template.sections[0].header.paragraphs[0].text = "Header <img1>"
            template.sections[0].footer.paragraphs[0].text = "Footer <img2>"
            template.save(template_path)

            WordService().build_document_without_com(
                str(template_path),
                str(output_path),
                [image_one, image_two],
                2,
            )

            self.assertTrue(output_path.exists())

            generated = DocxDocument(output_path)
            collected_text = self._collect_text(generated)
            self.assertIn("Body", collected_text)
            self.assertIn("Table", collected_text)
            self.assertIn("Header", collected_text)
            self.assertIn("Footer", collected_text)
            self.assertNotIn("<img1>", collected_text)
            self.assertNotIn("<img2>", collected_text)

            with zipfile.ZipFile(output_path) as archive:
                self.assertTrue(any(name.startswith("word/media/") for name in archive.namelist()))

    def test_word_service_exposes_only_docx_builder_and_no_com_keywords(self) -> None:
        public_methods = {
            name
            for name, member in inspect.getmembers(WordService, predicate=inspect.isfunction)
            if not name.startswith("_")
        }
        expected_methods = {
            "open",
            "create_from_template",
            "prepare_placeholder_document",
            "replace_image_placeholder",
            "clear_unused_image_placeholders",
            "cleanup_blank_pages",
            "print_document",
            "save_document_copy",
            "close",
            "build_document_without_com",
        }
        self.assertTrue(expected_methods.issubset(public_methods))

        source = inspect.getsource(word_service_module)
        for token in (
            "dynamic.Dispatch(\"Word.Application\")",
            "CoInitialize",
            "CoUninitialize",
            "Word.Application",
            "Documents.Open",
            "InlineShapes.AddPicture",
            "PrintOut(",
            "SaveAs2(",
        ):
            self.assertIn(token, source)

    @staticmethod
    def _collect_text(doc: DocxDocument) -> str:
        parts: list[str] = []
        parts.extend(paragraph.text for paragraph in doc.paragraphs)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    parts.extend(paragraph.text for paragraph in cell.paragraphs)
        for section in doc.sections:
            parts.extend(paragraph.text for paragraph in section.header.paragraphs)
            parts.extend(paragraph.text for paragraph in section.footer.paragraphs)
        return "\n".join(parts)

    @staticmethod
    def _write_image(path: Path, rgb: tuple[int, int, int]) -> None:
        image = Image.new("RGB", (64, 32), rgb)
        image.save(path)


if __name__ == "__main__":
    unittest.main()
