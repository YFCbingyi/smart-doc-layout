"""页码处理工具函数。

python-docx不直接支持页码字段插入，需要通过操作底层OXML实现。
"""

from __future__ import annotations

from typing import Optional

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph


def has_page_number(paragraph: Paragraph) -> bool:
    """检测段落中是否包含页码字段。

    Args:
        paragraph: python-docx段落对象。

    Returns:
        如果段落中包含 PAGE 或 NUMPAGES 字段，返回 True。
    """
    for run in paragraph.runs:
        fld_chars = run._element.findall(qn("w:fldChar"))
        instr_texts = run._element.findall(qn("w:instrText"))
        for elem in fld_chars:
            if elem.get(qn("w:fldCharType")) in ("begin", "separate", "end"):
                # 检查附近是否有 instrText 包含 PAGE 或 NUMPAGES
                parent = run._element.getparent()
                if parent is not None:
                    for instr in parent.iter(qn("w:instrText")):
                        if instr.text and ("PAGE" in instr.text.upper() or "NUMPAGES" in instr.text.upper()):
                            return True
        for elem in instr_texts:
            if elem.text and ("PAGE" in elem.text.upper() or "NUMPAGES" in elem.text.upper()):
                return True

    # 额外检查：遍历整个段落XML查找fldChar
    for fld_char in paragraph._element.iter(qn("w:fldChar")):
        fld_type = fld_char.get(qn("w:fldCharType"))
        if fld_type == "begin":
            # 查找后续兄弟元素中的 instrText
            parent = fld_char.getparent()
            if parent is not None:
                for instr in parent.iter(qn("w:instrText")):
                    if instr.text and ("PAGE" in instr.text.upper() or "NUMPAGES" in instr.text.upper()):
                        return True

    return False


def get_page_number_format(paragraph: Paragraph) -> Optional[str]:
    """获取段落中的页码格式字符串。

    Args:
        paragraph: python-docx段落对象。

    Returns:
        页码格式字符串，如 '{PAGE}' / '{NUMPAGES}'，未找到时返回 None。
    """
    for fld_char in paragraph._element.iter(qn("w:fldChar")):
        fld_type = fld_char.get(qn("w:fldCharType"))
        if fld_type == "begin":
            parent = fld_char.getparent()
            if parent is not None:
                for instr in parent.iter(qn("w:instrText")):
                    if instr.text:
                        text = instr.text.strip()
                        if "PAGE" in text.upper() or "NUMPAGES" in text.upper():
                            return text
    return None


def _make_field_element(field_code: str) -> object:
    """创建一个OXML字段元素（如 PAGE / NUMPAGES）。

    参考 https://stackoverflow.com/questions/26465889 实现。

    Args:
        field_code: 字段代码，如 'PAGE' 或 'NUMPAGES'。

    Returns:
        包含字段的OXML元素。
    """
    run = OxmlElement("w:r")
    run.set(qn("xml:space"), "preserve")

    # w:fldChar w:fldCharType="begin"
    fld_char_begin = OxmlElement("w:fldChar")
    fld_char_begin.set(qn("w:fldCharType"), "begin")
    run.append(fld_char_begin)

    # w:instrText
    instr_text = OxmlElement("w:instrText")
    instr_text.set(qn("xml:space"), "preserve")
    instr_text.text = field_code
    run.append(instr_text)

    # w:fldChar w:fldCharType="end"
    fld_char_end = OxmlElement("w:fldChar")
    fld_char_end.set(qn("w:fldCharType"), "end")
    run.append(fld_char_end)

    return run


def add_page_number(
    paragraph: Paragraph,
    format_str: str = "第 {PAGE} 页",
    add_space_prefix: bool = True,
) -> None:
    """向段落中添加页码字段。

    Args:
        paragraph: python-docx段落对象。
        format_str: 页码格式字符串，支持 {PAGE} 和 {NUMPAGES} 占位符。
                    默认: '第 {PAGE} 页'
        add_space_prefix: 是否在页码前添加空格。
    """
    # 如果已有页码，先移除
    if has_page_number(paragraph):
        remove_page_number(paragraph)

    # 清除段落原有内容
    paragraph.clear()

    # 解析格式字符串，将 {PAGE} 和 {NUMPAGES} 替换为字段
    parts = []
    i = 0
    while i < len(format_str):
        if format_str[i : i + 6] == "{PAGE}":
            parts.append(("field", "PAGE"))
            i += 6
        elif format_str[i : i + 10] == "{NUMPAGES}":
            parts.append(("field", "NUMPAGES"))
            i += 10
        else:
            # 收集普通文本
            j = i
            while j < len(format_str):
                remaining = format_str[j:]
                if remaining.startswith("{PAGE}") or remaining.startswith("{NUMPAGES}"):
                    break
                j += 1
            if j > i:
                parts.append(("text", format_str[i:j]))
            i = j

    # 构建段落内容
    for part_type, part_content in parts:
        if part_type == "text":
            run = paragraph.add_run(part_content)
        elif part_type == "field":
            field_elem = _make_field_element(part_content)
            paragraph._element.append(field_elem)

    if add_space_prefix:
        # 如果段落原本有其他内容，在前面加空格（实际已clear，不需额外处理）
        pass


def remove_page_number(paragraph: Paragraph) -> None:
    """移除段落中的页码字段。

    Args:
        paragraph: python-docx段落对象。
    """
    # 收集所有需要移除的 run 元素
    runs_to_remove = []
    for run_elem in paragraph._element.findall(qn("w:r")):
        fld_chars = run_elem.findall(qn("w:fldChar"))
        instr_texts = run_elem.findall(qn("w:instrText"))
        has_page_ref = any(
            instr.text and ("PAGE" in instr.text.upper() or "NUMPAGES" in instr.text.upper())
            for instr in instr_texts
        )
        has_fld_char = any(
            fld.get(qn("w:fldCharType")) in ("begin", "end")
            for fld in fld_chars
        )
        if has_page_ref or has_fld_char:
            runs_to_remove.append(run_elem)

    for run_elem in runs_to_remove:
        paragraph._element.remove(run_elem)
