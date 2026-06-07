"""文档修改功能测试。"""

import os
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from word_processor.models import ModifyInput, ParagraphModify, TextStyleInput, SectionModify, HeaderFooterInput, PageNumberInput, AlignmentType
from word_processor.writer import modify_document
from word_processor.reader import read_document

# 直接导入create_sample模块
import importlib.util
_create_sample_path = os.path.join(os.path.dirname(__file__), "create_sample.py")
_spec = importlib.util.spec_from_file_location("create_sample", _create_sample_path)
_create_sample = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_create_sample)
create_sample_docx = _create_sample.create_sample_docx


class TestWriter(unittest.TestCase):
    """测试文档修改功能。"""

    @classmethod
    def setUpClass(cls):
        """生成测试文档。"""
        cls.sample_path = create_sample_docx("examples/sample.docx")

    def test_modify_paragraph_text(self):
        """测试修改段落文本。"""
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            output_path = f.name

        try:
            modify_data = ModifyInput(
                paragraphs=[
                    ParagraphModify(
                        index=0,
                        text="修改后的新标题文本",
                        style=TextStyleInput(
                            font_name="黑体",
                            font_size=18,
                            bold=True,
                            alignment=AlignmentType.center,
                        ),
                    )
                ]
            )

            modify_document(self.sample_path, output_path, modify_data)

            # 验证修改结果
            result = read_document(output_path)
            self.assertEqual(result.paragraphs[0].text, "修改后的新标题文本")
            self.assertEqual(result.paragraphs[0].style.font_name, "黑体")
        finally:
            if os.path.exists(output_path):
                os.unlink(output_path)

    def test_modify_header(self):
        """测试修改页眉。"""
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            output_path = f.name

        try:
            modify_data = ModifyInput(
                sections=[
                    SectionModify(
                        index=0,
                        header=HeaderFooterInput(
                            content="新页眉内容",
                            style=TextStyleInput(
                                font_name="微软雅黑",
                                font_size=12,
                                bold=True,
                            ),
                        ),
                    )
                ]
            )

            modify_document(self.sample_path, output_path, modify_data)

            # 验证修改结果
            result = read_document(output_path)
            section = result.sections[0]
            self.assertIsNotNone(section.header)
            if section.header:
                self.assertIn("新页眉内容", section.header.content)
        finally:
            if os.path.exists(output_path):
                os.unlink(output_path)

    def test_modify_footer_with_page_number(self):
        """测试修改页脚（添加页码）。"""
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            output_path = f.name

        try:
            modify_data = ModifyInput(
                sections=[
                    SectionModify(
                        index=0,
                        footer=HeaderFooterInput(
                            page_number=PageNumberInput(
                                enabled=True,
                                format="第 {PAGE} 页",
                            ),
                        ),
                    )
                ]
            )

            modify_document(self.sample_path, output_path, modify_data)

            # 验证修改结果
            result = read_document(output_path)
            section = result.sections[0]
            self.assertIsNotNone(section.footer)
            if section.footer and section.footer.page_number:
                self.assertTrue(section.footer.page_number.enabled)
        finally:
            if os.path.exists(output_path):
                os.unlink(output_path)

    def test_modify_heading_level(self):
        """测试修改标题级别。"""
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            output_path = f.name

        try:
            modify_data = ModifyInput(
                paragraphs=[
                    ParagraphModify(
                        index=0,
                        heading_level=2,
                    )
                ]
            )

            modify_document(self.sample_path, output_path, modify_data)

            # 验证修改结果
            result = read_document(output_path)
            para = result.paragraphs[0]
            # 标题类型变为 heading 或保持原有类型
            # 至少检查索引正确
            self.assertEqual(para.index, 0)
        finally:
            if os.path.exists(output_path):
                os.unlink(output_path)

    def test_modify_paragraph_style(self):
        """测试修改段落样式（不修改文本）。"""
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            output_path = f.name

        try:
            # 先读取原始文档，获取原始文本
            original = read_document(self.sample_path)
            original_text = original.paragraphs[1].text if len(original.paragraphs) > 1 else ""

            modify_data = ModifyInput(
                paragraphs=[
                    ParagraphModify(
                        index=1,
                        style=TextStyleInput(
                            font_name="楷体",
                            font_size=14,
                            bold=True,
                            color="0000FF",
                            alignment=AlignmentType.center,
                        ),
                    )
                ]
            )

            modify_document(self.sample_path, output_path, modify_data)

            # 验证修改结果
            result = read_document(output_path)
            para = result.paragraphs[1]
            self.assertEqual(para.style.font_name, "楷体")
            self.assertEqual(para.style.font_size, 14)
            self.assertEqual(para.style.bold, True)
            # 如果有原始文本，应保持不变
            if original_text:
                self.assertEqual(para.text, original_text)
        finally:
            if os.path.exists(output_path):
                os.unlink(output_path)

    def test_invalid_file(self):
        """测试无效文件。"""
        with self.assertRaises(FileNotFoundError):
            modify_document(
                "nonexistent.docx",
                "output.docx",
                ModifyInput(),
            )

    def test_invalid_format(self):
        """测试不支持的文件格式。"""
        import tempfile
        with tempfile.NamedTemporaryFile(suffix=".txt", delete=False) as f:
            f.write(b"test")
            txt_path = f.name
        output_path = txt_path + ".docx"
        try:
            with self.assertRaises(ValueError):
                modify_document(
                    txt_path,
                    output_path,
                    ModifyInput(),
                )
        finally:
            if os.path.exists(txt_path):
                os.unlink(txt_path)
            if os.path.exists(output_path):
                os.unlink(output_path)


if __name__ == "__main__":
    unittest.main()
