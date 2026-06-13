"""Word文档识别与修改的数据模型定义。

使用Pydantic定义JSON schema，确保数据结构的清晰与类型安全。

⚠️ 识别输出和修改输入使用同一套模型（以修改输入为基准），
   保证 JSON 格式完全一致，方便用户直接编辑识别结果作为修改输入。
"""

from __future__ import annotations

from enum import Enum
from typing import Optional

from pydantic import BaseModel, Field


class ParagraphType(str, Enum):
    """段落类型枚举。"""

    title = "title"
    subtitle = "subtitle"
    heading = "heading"
    body = "body"
    list = "list"
    table = "table"
    other = "other"


class AlignmentType(str, Enum):
    """段落对齐方式枚举。"""

    left = "left"
    center = "center"
    right = "right"
    justify = "justify"
    unknown = "unknown"


class TextStyleInput(BaseModel):
    """文本样式（识别输出与修改输入共用）。

    所有字段可选，识别时尽可能填充，修改时按需填写。
    """

    font_name: Optional[str] = Field(None, description="字体名称")
    font_size: Optional[float] = Field(None, description="字号（磅值）")
    bold: Optional[bool] = Field(None, description="是否加粗")
    italic: Optional[bool] = Field(None, description="是否斜体")
    underline: Optional[bool] = Field(None, description="是否有下划线")
    color: Optional[str] = Field(None, description="字体颜色，十六进制，如 'FF0000'")
    alignment: Optional[AlignmentType] = Field(
        None, description="对齐方式：left/center/right/justify/unknown"
    )
    line_spacing: Optional[float] = Field(None, description="行距（倍数或磅值，由line_spacing_is_pt决定）")
    line_spacing_is_pt: Optional[bool] = Field(None, description="行距是否为磅值（True=磅值，False=倍数）")
    space_before: Optional[float] = Field(None, description="段前间距（磅值）")
    space_after: Optional[float] = Field(None, description="段后间距（磅值）")
    first_line_indent: Optional[float] = Field(None, description="首行缩进（磅值）")


class PageNumberInput(BaseModel):
    """页码修改/识别输入模型。"""

    enabled: bool = Field(default=True, description="是否启用页码")
    format: str = Field(default="第 {PAGE} 页", description="页码格式，使用 {PAGE} {NUMPAGES} 占位")


class HeaderFooterInput(BaseModel):
    """页眉/页脚数据（识别输出与修改输入共用）。"""

    content: Optional[str] = Field(None, description="页眉/页脚文本内容")
    style: Optional[TextStyleInput] = Field(None, description="页眉/页脚样式")
    page_number: Optional[PageNumberInput] = Field(None, description="页码设置")


class SectionModify(BaseModel):
    """节（Section）数据（识别输出与修改输入共用）。"""

    index: int = Field(..., description="节序号（从0开始）")
    header: Optional[HeaderFooterInput] = Field(None, description="奇数页页眉数据")
    footer: Optional[HeaderFooterInput] = Field(None, description="奇数页页脚数据")
    even_page_header: Optional[HeaderFooterInput] = Field(
        None, description="偶数页页眉数据（启用奇偶分页时有效）"
    )
    even_page_footer: Optional[HeaderFooterInput] = Field(
        None, description="偶数页页脚数据（启用奇偶分页时有效）"
    )
    enable_odd_even: bool = Field(
        default=False, description="是否启用了奇偶页不同的页眉页脚"
    )
    footer_distance_cm: Optional[float] = Field(
        None, ge=0, le=10, description="页脚距底端距离(cm)"
    )


class ParagraphModify(BaseModel):
    """段落数据（识别输出与修改输入共用）。

    识别时填充所有字段；修改时只需填写需要变更的字段。
    """

    index: int = Field(..., description="段落序号（从0开始）")
    type: Optional[ParagraphType] = Field(None, description="段落类型")
    heading_level: Optional[int] = Field(None, description="标题级别（1-9）")
    text: Optional[str] = Field(None, description="段落文本内容")
    style: Optional[TextStyleInput] = Field(None, description="段落样式")


class ModifyInput(BaseModel):
    """文档识别/修改统一数据模型。

    识别输出 → 填充所有字段供查看；
    修改输入 → 按需填写要变更的字段（未填字段保持原样）。
    """

    sections: list[SectionModify] = Field(default_factory=list, description="节列表")
    paragraphs: list[ParagraphModify] = Field(default_factory=list, description="段落列表")
