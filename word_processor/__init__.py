"""Word文档识别与修改工具包。

提供对Word文档(.docx)的读取识别和修改写入功能。
"""

from word_processor.models import ModifyInput, ParagraphType, AlignmentType
from word_processor.reader import read_document
from word_processor.writer import modify_document

__all__ = [
    "ModifyInput",
    "ParagraphType",
    "AlignmentType",
    "read_document",
    "modify_document",
]
