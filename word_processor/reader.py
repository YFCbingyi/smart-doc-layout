"""Word文档读取识别模块。

读取 .docx 文件，提取段落类型、样式、页眉、页码等信息，
并支持智能识别：主标题/副标题、四级标题（一、/（一）/1./(1)）。

输出格式与修改输入格式完全一致（使用 ModifyInput 模型）。
"""

from __future__ import annotations

import os
import re
from typing import Optional

from docx import Document as DocxDocument
from docx.document import Document as DocumentType
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.section import Section, _Header, _Footer
from docx.shared import Pt, RGBColor
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from word_processor.models import (
    AlignmentType,
    HeaderFooterInput,
    ModifyInput,
    PageNumberInput,
    ParagraphModify,
    ParagraphType,
    SectionModify,
    TextStyleInput,
)
from word_processor.page_number import has_page_number, get_page_number_format

# ═══════════════════════════════════════════════════
# 四级标题正则模式
# ═══════════════════════════════════════════════════

_PATTERN_LEVEL1 = re.compile(r"^[一二三四五六七八九十百千]+[、]")
_PATTERN_LEVEL2 = re.compile(r"^[（(][一二三四五六七八九十百千]+[）)]")
_PATTERN_LEVEL3 = re.compile(r"^\d+\.[\s]")
_PATTERN_LEVEL4 = re.compile(r"^[（(]\d+[）)]")

_HEADING_PATTERNS: list[tuple[re.Pattern, int]] = [
    (_PATTERN_LEVEL1, 1),
    (_PATTERN_LEVEL2, 2),
    (_PATTERN_LEVEL3, 3),
    (_PATTERN_LEVEL4, 4),
]


def _alignment_to_model(alignment: Optional[WD_ALIGN_PARAGRAPH]) -> AlignmentType:
    """将 python-docx 对齐方式转换为模型枚举。"""
    mapping = {
        WD_ALIGN_PARAGRAPH.LEFT: AlignmentType.left,
        WD_ALIGN_PARAGRAPH.CENTER: AlignmentType.center,
        WD_ALIGN_PARAGRAPH.RIGHT: AlignmentType.right,
        WD_ALIGN_PARAGRAPH.JUSTIFY: AlignmentType.justify,
    }
    return mapping.get(alignment, AlignmentType.unknown)


def _get_font_color(run: Run) -> Optional[str]:
    """获取Run的字体颜色，返回十六进制字符串。"""
    try:
        color = run.font.color
        if color and color.rgb:
            return str(color.rgb)
    except Exception:
        pass
    return None


def _get_paragraph_type(paragraph: Paragraph) -> tuple[ParagraphType, Optional[int]]:
    """根据Word样式名判断段落类型和标题级别（仅基于样式）。"""
    style = paragraph.style
    if style is None or style.name is None:
        return ParagraphType.body, None

    style_name = style.name

    if style_name.startswith("Heading"):
        try:
            level = int(style_name.split()[-1])
            return ParagraphType.heading, level
        except (ValueError, IndexError):
            return ParagraphType.heading, 1

    if "List" in style_name or style_name.startswith("List"):
        return ParagraphType.list, None

    return ParagraphType.body, None


def _detect_heading_by_content(para: ParagraphModify) -> tuple[ParagraphType, Optional[int]]:
    """根据段落文本内容智能识别标题类型。"""
    text = (para.text or "").strip()
    if not text:
        return para.type or ParagraphType.body, para.heading_level

    for pattern, level in _HEADING_PATTERNS:
        if pattern.match(text):
            return ParagraphType.heading, level

    return para.type or ParagraphType.body, para.heading_level


def _detect_title_subtitle(paragraphs: list[ParagraphModify]) -> None:
    """智能识别文档开头的主标题和副标题。"""
    body_sizes = [
        p.style.font_size
        for p in paragraphs
        if p.style is not None and p.style.font_size is not None and p.type == ParagraphType.body
    ]
    avg_body_size = sum(body_sizes) / len(body_sizes) if body_sizes else 12.0

    title_candidate: Optional[ParagraphModify] = None
    for p in paragraphs:
        text = (p.text or "").strip()
        if not text:
            continue
        if p.style is None or p.style.alignment != AlignmentType.center:
            break

        if title_candidate is None:
            p.type = ParagraphType.title
            title_candidate = p
        else:
            same_font = (
                p.style is not None
                and title_candidate.style is not None
                and p.style.font_name == title_candidate.style.font_name
                and p.style.font_size == title_candidate.style.font_size
            )
            p.type = ParagraphType.title if same_font else ParagraphType.subtitle


def _extract_paragraph_style(paragraph: Paragraph) -> TextStyleInput:
    """提取段落样式。"""
    pf = paragraph.paragraph_format
    alignment = paragraph.alignment

    font_name = None
    font_size = None
    bold = None
    color = None
    if paragraph.runs:
        first_run = paragraph.runs[0]
        font = first_run.font
        font_name = font.name if font.name else None
        font_size = font.size.pt if font.size else None
        bold = font.bold if font.bold is not None else None
        color = _get_font_color(first_run)

    # 规范化行距：将 OXML EMU 值转换为磅值
    raw_ls = pf.line_spacing
    line_spacing = None
    if raw_ls is not None:
        if isinstance(raw_ls, (int, float)) and raw_ls > 1000:
            line_spacing = round(raw_ls / 12700, 1)  # EMU → 磅
        else:
            line_spacing = raw_ls  # 已经是磅值或倍数

    return TextStyleInput(
        font_name=font_name,
        font_size=font_size,
        bold=bold,
        color=color,
        alignment=_alignment_to_model(alignment),
        line_spacing=line_spacing,
        space_before=pf.space_before.pt if pf.space_before else None,
        space_after=pf.space_after.pt if pf.space_after else None,
        first_line_indent=pf.first_line_indent.pt if pf.first_line_indent else None,
    )


def _extract_paragraph_data(index: int, paragraph: Paragraph) -> ParagraphModify:
    """提取单个段落的数据（基于样式）。"""
    para_type, heading_level = _get_paragraph_type(paragraph)

    return ParagraphModify(
        index=index,
        type=para_type,
        heading_level=heading_level,
        text=paragraph.text,
        style=_extract_paragraph_style(paragraph),
    )


def _extract_header_footer(
    hf_obj: Optional[_Header | _Footer],
) -> Optional[HeaderFooterInput]:
    """提取页眉或页脚的数据，返回 HeaderFooterInput。"""
    if hf_obj is None:
        return None

    try:
        full_text = ""
        content_style: Optional[TextStyleInput] = None
        pn_config: Optional[PageNumberInput] = None

        for i, para in enumerate(hf_obj.paragraphs):
            if para.text:
                if full_text:
                    full_text += "\n"
                full_text += para.text

            # 提取样式（取第一个段落）
            if i == 0:
                content_style = _extract_paragraph_style(para)

            # 检测页码
            if pn_config is None and has_page_number(para):
                fmt = get_page_number_format(para)
                if fmt == "PAGE":
                    fmt = "{PAGE}"
                pn_config = PageNumberInput(enabled=True, format=fmt or "{PAGE}")

        return HeaderFooterInput(
            content=full_text if full_text else None,
            style=content_style,
            page_number=pn_config,
        )
    except Exception:
        return None


def _extract_section_data(index: int, section: Section) -> SectionModify:
    """提取节(Section)的数据。"""
    header = _extract_header_footer(section.header)
    footer = _extract_header_footer(section.footer)

    # 提取偶数页页眉/页脚
    even_header = None
    even_footer = None
    odd_even_enabled = False

    try:
        odd_even_enabled = section.different_first_page_header_footer or section.even_page_header.is_linked_to_previous is not None
        # 实际检查偶数页页眉和页脚是否有内容
        even_header = _extract_header_footer(section.even_page_header)
        even_footer = _extract_header_footer(section.even_page_footer)
        if even_header is not None or even_footer is not None:
            odd_even_enabled = True
    except Exception:
        pass

    return SectionModify(
        index=index,
        header=header,
        footer=footer,
        even_page_header=even_header,
        even_page_footer=even_footer,
        enable_odd_even=odd_even_enabled,
    )


def read_document(file_path: str, smart: bool = True) -> ModifyInput:
    """读取Word文档，提取结构化数据。

    处理流程：
      1. 按Word样式提取基本段落数据
      2. 智能识别四级标题（一、/（一）/1./(1)）(smart=True时)
      3. 智能识别主标题/副标题 (smart=True时)

    Args:
        file_path: .docx文件路径。
        smart: 是否启用智能识别（识别主副标题、四级标题内容）。
               默认True；设为False时仅基于Word实际样式识别。

    Returns:
        ModifyInput: 统一数据模型，可直接编辑作为修改输入。

    Raises:
        FileNotFoundError: 文件不存在。
        ValueError: 文件格式不支持。
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"文件不存在: {file_path}")

    if not file_path.lower().endswith(".docx"):
        raise ValueError("仅支持 .docx 格式的Word文档")

    doc: DocumentType = DocxDocument(file_path)

    # 提取节数据
    sections = [_extract_section_data(i, section) for i, section in enumerate(doc.sections)]

    # ── 第1步：按样式提取段落数据 ──
    paragraphs = [_extract_paragraph_data(i, para) for i, para in enumerate(doc.paragraphs)]

    if smart:
        # ── 第2步：智能识别四级标题（按内容） ──
        for p in paragraphs:
            if p.type == ParagraphType.body:
                new_type, new_level = _detect_heading_by_content(p)
                if new_type == ParagraphType.heading:
                    p.type = new_type
                    p.heading_level = new_level

        # ── 第3步：智能识别主标题/副标题（文档开头） ──
        _detect_title_subtitle(paragraphs)

    return ModifyInput(
        sections=sections,
        paragraphs=paragraphs,
    )
