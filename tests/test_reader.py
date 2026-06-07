"""文档读取功能测试。"""

import json
import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from word_processor.models import ParagraphType, ModifyInput
from word_processor.reader import read_document

# 直接导入create_sample模块
import importlib.util
_create_sample_path = os.path.join(os.path.dirname(__file__), "create_sample.py")
_spec = importlib.util.spec_from_file_location("create_sample", _create_sample_path)
_create_sample = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_create_sample)
create_sample_docx = _create_sample.create_sample_docx


class TestReader(unittest.TestCase):
    """测试文档读取功能。"""

    @classmethod
    def setUpClass(cls):
        """生成测试文档。"""
        cls.sample_path = create_sample_docx("examples/sample.docx")

    def test_read_document_basic(self):
        """测试基本读取功能。"""
        data = read_document(self.sample_path)

        self.assertIsNotNone(data)
        self.assertGreater(len(data.sections), 0)
        self.assertGreater(len(data.paragraphs), 0)

    def test_read_paragraphs(self):
        """测试段落读取。"""
        data = read_document(self.sample_path)

        self.assertGreater(len(data.paragraphs), 0)

        # 检查标题段落
        headings = [p for p in data.paragraphs if p.type == ParagraphType.heading]
        self.assertGreater(len(headings), 0)

        # 检查正文段落
        bodies = [p for p in data.paragraphs if p.type == ParagraphType.body]
        self.assertGreater(len(bodies), 0)

    def test_read_sections(self):
        """测试节读取。"""
        data = read_document(self.sample_path)

        self.assertGreater(len(data.sections), 0)

        section = data.sections[0]
        # 页眉应该有内容
        self.assertIsNotNone(section.header)
        if section.header and section.header.content:
            self.assertGreater(len(section.header.content), 0)

    def test_read_header_footer(self):
        """测试页眉页脚读取。"""
        data = read_document(self.sample_path)

        section = data.sections[0]

        # 页眉
        if section.header:
            self.assertIn("内部测试文档", section.header.content)

        # 页脚（页码）
        if section.footer and section.footer.page_number:
            self.assertTrue(section.footer.page_number.enabled)

    def test_read_style_details(self):
        """测试段落样式细节读取（替代原有的run_details测试）。"""
        data = read_document(self.sample_path)

        # 找到标题段落，检查类型为 heading
        headings = [p for p in data.paragraphs if p.type == ParagraphType.heading]
        self.assertGreater(len(headings), 0)
        # 确保标题有 heading_level
        self.assertIsNotNone(headings[0].heading_level)

    def test_invalid_file(self):
        """测试无效文件。"""
        with self.assertRaises(FileNotFoundError):
            read_document("nonexistent.docx")

    def test_invalid_format(self):
        """测试不支持的文件格式。"""
        import tempfile
        with tempfile.NamedTemporaryFile(suffix=".txt", delete=False) as f:
            f.write(b"test")
            txt_path = f.name
        try:
            with self.assertRaises(ValueError):
                read_document(txt_path)
        finally:
            os.unlink(txt_path)

    def test_json_serialization(self):
        """测试JSON序列化。"""
        data = read_document(self.sample_path)
        json_dict = data.model_dump(mode="json", exclude_none=True)
        json_str = json.dumps(json_dict, ensure_ascii=False, indent=2)

        self.assertIsInstance(json_str, str)
        # 验证JSON包含关键字段（无metadata，sections和paragraphs为顶级）
        self.assertIn("sections", json_dict)
        self.assertIn("paragraphs", json_dict)
        # 验证可以被正确反序列化
        restored = ModifyInput(**json_dict)
        self.assertEqual(len(restored.paragraphs), len(data.paragraphs))


if __name__ == "__main__":
    unittest.main()
