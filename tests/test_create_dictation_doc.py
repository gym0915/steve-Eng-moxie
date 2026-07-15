import importlib.util
import tempfile
import unittest
import zipfile
from pathlib import Path

from docx import Document
from PIL import Image, ImageChops


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "scripts" / "create_dictation_doc.py"


def load_module():
    spec = importlib.util.spec_from_file_location("create_dictation_doc", SCRIPT)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class EnglishDictationGeneratorTests(unittest.TestCase):
    def test_reads_five_column_a4_contract_from_bundled_html(self):
        module = load_module()
        contract = module.load_template_contract()
        self.assertEqual(contract["columns"], 5)
        self.assertEqual(contract["accent_color"], "C0392B")
        self.assertEqual(contract["ink_color"], "1A1A1A")
        self.assertEqual(contract["muted_color"], "888")
        self.assertEqual(contract["title"], "看中文写英文")
        self.assertEqual(contract["meta_fields"], [
            "姓名：____________",
            "班级：__________",
            "日期：______月______日",
            "得分：__________",
        ])
        self.assertEqual(contract["page_margin_mm"], 12)
        self.assertEqual(contract["prompt_px"], 16)
        self.assertEqual(contract["answer_px"], 17)
        self.assertEqual(contract["meta_border_px"], 2)
        self.assertEqual(contract["meta_padding_px"], 8)
        self.assertEqual(contract["meta_padding_x_px"], 4)
        self.assertEqual(contract["meta_margin_bottom_px"], 22)
        self.assertEqual(contract["section_border_px"], 5)
        self.assertEqual(contract["section_padding_left_px"], 10)
        self.assertEqual(contract["section_margin_bottom_px"], 18)
        self.assertEqual(contract["count_margin_left_px"], 8)
        self.assertEqual(contract["prompt_line_height_px"], 24)
        self.assertEqual(contract["answer_line_height_px"], 24)

    def test_long_prompt_wraps_without_touching_image_edges(self):
        module = load_module()
        with tempfile.TemporaryDirectory() as tmp:
            path = module._build_prompt_image(
                Path(tmp),
                "这是一个用于验证长中文释义不会被裁切的测试句子",
                0,
            )
            image = Image.open(path).convert("RGB")
            difference = ImageChops.difference(image, Image.new("RGB", image.size, "white"))
            bounds = difference.getbbox()
        self.assertGreater(image.height, 86)
        self.assertIsNotNone(bounds)
        self.assertGreater(bounds[0], 0)
        self.assertGreater(bounds[1], 0)
        self.assertLess(bounds[2], image.width)
        self.assertLess(bounds[3], image.height)

    def test_rejects_empty_items(self):
        module = load_module()
        with self.assertRaisesRegex(ValueError, "至少提供一个英语默写词条"):
            module.validate_items([])

    def test_rejects_prompt_with_english_metadata(self):
        module = load_module()
        with self.assertRaisesRegex(ValueError, "中文提示不得包含英文字母"):
            module.validate_items([{"prompt": "n. 星期一", "answer": "Monday"}])

    def test_preserves_source_order(self):
        module = load_module()
        items = [
            {"prompt": "星期一", "answer": "Monday"},
            {"prompt": "星期二", "answer": "Tuesday"},
            {"prompt": "星期三", "answer": "Wednesday"},
        ]
        self.assertEqual(module.validate_items(items), items)

    def test_generates_a4_five_column_student_doc_without_answers(self):
        module = load_module()
        items = [
            {"prompt": "星期一", "answer": "Monday"},
            {"prompt": "星期二", "answer": "Tuesday"},
            {"prompt": "通常", "answer": "usually"},
            {"prompt": "有时", "answer": "sometimes"},
            {"prompt": "总是", "answer": "always"},
            {"prompt": "经常", "answer": "often"},
        ]

        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "dictation.docx"
            module.create_document(items, output)
            doc = Document(output)
            section = doc.sections[0]
            self.assertAlmostEqual(section.page_width.cm, 21.0, places=1)
            self.assertAlmostEqual(section.page_height.cm, 29.7, places=1)
            self.assertEqual(len(doc.tables[-1].columns), 5)
            self.assertEqual(len(doc.tables[-1].rows), 2)
            self.assertEqual(len(doc.inline_shapes), len(items) + 2)
            with zipfile.ZipFile(output) as package:
                xml = package.read("word/document.xml").decode("utf-8")

        for prompt in ("星期一", "星期二", "通常", "有时", "总是", "经常"):
            self.assertIn(prompt, xml)
        for answer in ("Monday", "Tuesday", "usually", "sometimes", "always", "often"):
            self.assertNotIn(answer, xml)
        self.assertEqual(xml.count("<w:t>________</w:t>"), 6)
        self.assertIn("看中文写英文", xml)
        self.assertIn("共6题", xml)

    def test_repeats_template_header_and_preserves_order_after_sixty_five_items(self):
        module = load_module()
        items = [
            {"prompt": f"第{i}题", "answer": f"answer{i}"}
            for i in range(1, 67)
        ]
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "multipage.docx"
            module.create_document(items, output)
            doc = Document(output)
            with zipfile.ZipFile(output) as package:
                xml = package.read("word/document.xml").decode("utf-8")
        self.assertEqual(len(doc.sections), 2)
        self.assertEqual(len(doc.tables), 2)
        self.assertEqual(len(doc.inline_shapes), 66 + 4)
        self.assertEqual(xml.count("姓名、班级、日期、得分填写栏"), 2)
        self.assertLess(xml.index("中文提示：第65题"), xml.index("中文提示：第66题"))

    def test_long_prompt_rows_reduce_items_per_page_to_keep_headers_repeating(self):
        module = load_module()
        long_prompt = "这是一个需要自动换行并参与分页高度预算的中文释义测试句子"
        items = [
            {"prompt": long_prompt, "answer": f"answer{i}"}
            for i in range(25)
        ]
        pages = module._paginate_items(items)
        self.assertGreater(len(pages), 1)
        self.assertEqual(sum(len(page) for page in pages), 25)
        for page in pages:
            rows = [page[index:index + 5] for index in range(0, len(page), 5)]
            line_units = sum(max(len(module._wrap_prompt_lines(item["prompt"])) for item in row) for row in rows)
            self.assertLessEqual(line_units, module.MAX_CONTENT_LINE_UNITS_PER_PAGE)


if __name__ == "__main__":
    unittest.main()
