"""Word文档处理 Gradio Web 图形界面。

提供基于Gradio的Web UI，用于识别和修改Word文档的样式、内容、页眉/页脚等。
"""

from __future__ import annotations

import json
import os
import tempfile
from typing import Any, Optional

import gradio as gr

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
from word_processor.reader import read_document
from word_processor.writer import modify_document

# ─────────────────────────────────────────────
# 常量
# ─────────────────────────────────────────────

_FONT_CHOICES = ["宋体", "黑体", "仿宋", "楷体"]
_ALIGN_CHOICES = ["left", "center", "right", "unknown", "justify"]

# 标准 Word 中文字号映射表（中文名称 → pt值）
_WORD_FONT_SIZES: list[tuple[str, float]] = [
    ("初号", 42),
    ("小初", 36),
    ("一号", 26),
    ("小一", 24),
    ("二号", 22),
    ("小二", 18),
    ("三号", 16),
    ("小三", 15),
    ("四号", 14),
    ("小四", 12),
    ("五号", 10.5),
    ("小五", 9),
]
_FONT_SIZE_CHOICES = [f"{name}({int(pt) if pt == int(pt) else pt}pt)" for name, pt in _WORD_FONT_SIZES]

_DF_HEADERS = ["选择", "序号", "段落类型", "文本内容"]
_DF_DATATYPE = ["bool", "str", "str", "str"]
_PARAGRAPH_TYPE_CHOICES = ["主标题", "副标题", "标题一", "标题二", "标题三", "标题四", "正文"]

# ─────────────────────────────────────────────
# 辅助函数
# ─────────────────────────────────────────────


_PARAGRAPH_TYPE_CN: dict[str, str] = {
    "title": "主标题",
    "subtitle": "副标题",
    "body": "正文",
}


def _type_to_cn(p: ParagraphModify) -> str:
    """将段落类型和标题级别转换为中文显示。"""
    if p.type == ParagraphType.title:
        return "主标题"
    elif p.type == ParagraphType.subtitle:
        return "副标题"
    elif p.type == ParagraphType.heading:
        level_map = {1: "标题一", 2: "标题二", 3: "标题三", 4: "标题四"}
        return level_map.get(p.heading_level or 1, "标题一")
    else:
        return "正文"


def _make_style_preview(style: Optional[TextStyleInput]) -> str:
    """生成样式的简短文字预览。"""
    if style is None:
        return ""
    parts: list[str] = []
    if style.font_name:
        parts.append(style.font_name)
    if style.font_size:
        parts.append(f"{style.font_size}pt")
    if style.bold:
        parts.append("加粗")
    if style.color:
        parts.append(f"#{style.color}")
    return " ".join(parts) if parts else ""


def _find_first_style(
    paragraphs: list[ParagraphModify], *types: ParagraphType, heading_level: Optional[int] = None
) -> TextStyleInput:
    """从段落列表中提取首个匹配类型的样式。"""
    for p in paragraphs:
        if p.style is not None and p.type in types:
            if heading_level is not None and p.heading_level != heading_level:
                continue
            return p.style
    return TextStyleInput()


def _safe_choice(value: Optional[str], choices: list[str]) -> Optional[str]:
    """确保值在可选列表中，否则返回 None。"""
    if value in choices:
        return value
    return None


def _closest_font_size(pt_value: Optional[float]) -> Optional[str]:
    """将pt值匹配到最接近的标准字号选项。

    返回选项字符串如 "三号(16pt)"，若无pt值则返回 None。
    """
    if pt_value is None:
        return None
    closest = min(_FONT_SIZE_CHOICES, key=lambda c: abs(
        float(c.split("(")[1].rstrip("pt)")) - pt_value
    ))
    return closest


def _pt_to_chars(pt_val: Optional[float], font_size_pt: Optional[float] = None) -> float:
    """将磅值转换为字符数（1字符 ≈ font_size_pt 磅）。"""
    if pt_val is None or not pt_val:
        return 0.0
    fs = font_size_pt or 16.0
    return round(pt_val / fs, 1)


def _parse_font_size(size_str: Optional[str]) -> Optional[float]:
    """从 \"中文名(pt值)\" 字符串中解析pt值。"""
    if not size_str:
        return None
    try:
        return float(size_str.split("(")[1].rstrip("pt)"))
    except (IndexError, ValueError):
        return None


# ─────────────────────────────────────────────
# 界面构建
# ─────────────────────────────────────────────

_CSS = """
footer {display: none !important;}
"""


def build_ui() -> gr.Blocks:
    """构建 Gradio Web UI。"""
    with gr.Blocks(title="Word文档处理器") as demo:
        gr.Markdown("# 📄 Word文档处理器")
        gr.Markdown("上传Word文档(.docx)，识别并修改文档样式、内容、页眉页脚。")

        # ── 状态 ──
        state_data = gr.State()  # 存储 ModifyInput dict
        state_path = gr.State()  # 存储输入文件路径


        # ── 1. 文件输入/输出 ──
        with gr.Row():
            file_input = gr.File(
                label="上传Word文档",
                file_types=[".docx"],
                file_count="single",
            )
            file_output = gr.File(
                label="下载处理后的文档",
                interactive=False,
            )

        # ── 2. 识别按钮 ──
        with gr.Row():
            btn_recognize_raw = gr.Button("原文格式识别", variant="secondary", size="lg")
            btn_recognize_smart = gr.Button("智能辅助识别", variant="secondary", size="lg")

        # ── 3. 基本配置 ──
        with gr.Accordion("基本配置", open=False):
            with gr.Row():
                txt_header_content = gr.Textbox(label="页眉内容", scale=4)
                dd_header_font = gr.Dropdown(
                    choices=_FONT_CHOICES, label="页眉字体", scale=2
                )
                dd_header_size = gr.Dropdown(
                    choices=_FONT_SIZE_CHOICES, label="页眉字号", value="五号(10.5pt)"
                )
                dd_header_align = gr.Dropdown(
                    choices=_ALIGN_CHOICES, label="页眉对齐", value="center"
                )
            with gr.Row():
                chk_page_number = gr.Checkbox(label="页码启用", value=False)
                txt_page_format = gr.Textbox(
                    label="页码格式", value="{PAGE}", scale=3
                )
                dd_page_align = gr.Dropdown(
                    choices=_ALIGN_CHOICES, label="页码对齐", value="center", scale=2
                )
                num_footer_distance = gr.Number(
                    label="页脚距(cm)", value=1.5, minimum=0, maximum=10, step=0.1
                )
            with gr.Row():
                chk_odd_even = gr.Checkbox(label="奇偶分页启用", value=False)
            with gr.Row(visible=False) as even_config_row:
                txt_even_header = gr.Textbox(label="偶数页页眉", scale=3)
                dd_even_header_font = gr.Dropdown(
                    choices=_FONT_CHOICES, label="偶数页字体", scale=2
                )
                dd_even_header_size = gr.Dropdown(
                    choices=_FONT_SIZE_CHOICES, label="偶数页字号", value="五号(10.5pt)"
                )
                dd_even_header_align = gr.Dropdown(
                    choices=_ALIGN_CHOICES, label="偶数页对齐", value="center"
                )
            with gr.Row(visible=False) as even_footer_row:
                dd_even_page_align = gr.Dropdown(
                    choices=_ALIGN_CHOICES,
                    label="偶数页页码对齐",
                    value="center",
                )
            chk_odd_even.change(
                fn=lambda v: (gr.update(visible=v), gr.update(visible=v)),
                inputs=[chk_odd_even],
                outputs=[even_config_row, even_footer_row],
            )
            with gr.Row():
                num_margin_top = gr.Number(
                    label="上边距(cm)", value=2.54, minimum=0, maximum=10, step=0.1
                )
                num_margin_bottom = gr.Number(
                    label="下边距(cm)", value=2.54, minimum=0, maximum=10, step=0.1
                )
                num_margin_left = gr.Number(
                    label="左边距(cm)", value=2.54, minimum=0, maximum=10, step=0.1
                )
                num_margin_right = gr.Number(
                    label="右边距(cm)", value=2.54, minimum=0, maximum=10, step=0.1
                )

        # ── 4. 样式配置 ──
        with gr.Accordion("样式配置", open=True):
            with gr.Row():
                dd_title_font = gr.Dropdown(choices=_FONT_CHOICES, label="主标题字体", scale=2)
                dd_title_size = gr.Dropdown(choices=_FONT_SIZE_CHOICES, label="主标题字号", value="三号(16pt)")
                chk_title_bold = gr.Checkbox(label="主标题加粗", value=True)
                dd_title_align = gr.Dropdown(choices=_ALIGN_CHOICES, label="主标题对齐", value="center", scale=2)
                num_title_indent = gr.Number(label="主标题缩进(字符)", value=0, minimum=0, maximum=100, step=0.5)
            with gr.Row():
                dd_subtitle_font = gr.Dropdown(choices=_FONT_CHOICES, label="副标题字体", scale=2)
                dd_subtitle_size = gr.Dropdown(choices=_FONT_SIZE_CHOICES, label="副标题字号", value="三号(16pt)")
                chk_subtitle_bold = gr.Checkbox(label="副标题加粗", value=False)
                dd_subtitle_align = gr.Dropdown(choices=_ALIGN_CHOICES, label="副标题对齐", value="center", scale=2)
                num_subtitle_indent = gr.Number(label="副标题缩进(字符)", value=0, minimum=0, maximum=100, step=0.5)
            with gr.Row():
                dd_heading1_font = gr.Dropdown(choices=_FONT_CHOICES, label="标题一字体", scale=2)
                dd_heading1_size = gr.Dropdown(choices=_FONT_SIZE_CHOICES, label="标题一字号", value="三号(16pt)")
                chk_heading1_bold = gr.Checkbox(label="标题一加粗", value=True)
                dd_heading1_align = gr.Dropdown(choices=_ALIGN_CHOICES, label="标题一对齐", value="center", scale=2)
                num_heading1_indent = gr.Number(label="标题一缩进(字符)", value=0, minimum=0, maximum=100, step=0.5)
            with gr.Row():
                dd_heading2_font = gr.Dropdown(choices=_FONT_CHOICES, label="标题二字体", scale=2)
                dd_heading2_size = gr.Dropdown(choices=_FONT_SIZE_CHOICES, label="标题二字号", value="三号(16pt)")
                chk_heading2_bold = gr.Checkbox(label="标题二加粗", value=True)
                dd_heading2_align = gr.Dropdown(choices=_ALIGN_CHOICES, label="标题二对齐", value="center", scale=2)
                num_heading2_indent = gr.Number(label="标题二缩进(字符)", value=0, minimum=0, maximum=100, step=0.5)
            with gr.Row():
                dd_heading3_font = gr.Dropdown(choices=_FONT_CHOICES, label="标题三字体", scale=2)
                dd_heading3_size = gr.Dropdown(choices=_FONT_SIZE_CHOICES, label="标题三字号", value="三号(16pt)")
                chk_heading3_bold = gr.Checkbox(label="标题三加粗", value=True)
                dd_heading3_align = gr.Dropdown(choices=_ALIGN_CHOICES, label="标题三对齐", value="center", scale=2)
                num_heading3_indent = gr.Number(label="标题三缩进(字符)", value=0, minimum=0, maximum=100, step=0.5)
            with gr.Row():
                dd_heading4_font = gr.Dropdown(choices=_FONT_CHOICES, label="标题四字体", scale=2)
                dd_heading4_size = gr.Dropdown(choices=_FONT_SIZE_CHOICES, label="标题四字号", value="三号(16pt)")
                chk_heading4_bold = gr.Checkbox(label="标题四加粗", value=True)
                dd_heading4_align = gr.Dropdown(choices=_ALIGN_CHOICES, label="标题四对齐", value="center", scale=2)
                num_heading4_indent = gr.Number(label="标题四缩进(字符)", value=0, minimum=0, maximum=100, step=0.5)
            with gr.Row():
                dd_body_font = gr.Dropdown(choices=_FONT_CHOICES, label="正文字体", scale=2)
                dd_body_size = gr.Dropdown(choices=_FONT_SIZE_CHOICES, label="正文字号", value="四号(14pt)")
                dd_body_align = gr.Dropdown(choices=_ALIGN_CHOICES, label="正文对齐", value="unknown", scale=2)
                num_body_indent = gr.Number(label="正文缩进(字符)", value=2, minimum=0, maximum=100, step=0.5)
                num_body_line_spacing = gr.Number(label="行距", value=1.5, minimum=0.5, maximum=72, step=0.5)

        # ── 5. 段落列表 ──
        gr.Markdown("### 段落列表")

        # ── 操作区 ──
        with gr.Row():
            chk_select_all = gr.Checkbox(label="☑ 全选", value=False)
            dd_batch_type = gr.Dropdown(
                choices=_PARAGRAPH_TYPE_CHOICES,
                label="批量修改类型",
                scale=2,
            )
            btn_apply = gr.Button("应用", variant="primary", scale=1)

        df_paragraphs = gr.Dataframe(
            headers=_DF_HEADERS,
            datatype=_DF_DATATYPE,
            column_count=4,
            interactive=True,
            static_columns=[1, 2, 3],
            label="勾选需要修改的段落，通过上方操作区批量修改",
        )
        gr.Markdown("💡 使用说明：勾选段落左侧的复选框，选中「全选」可选中全部；选择类型后点击「应用」批量修改选中段落的类型")

        # ── 调试模式 ──
        with gr.Row():
            chk_debug_mode = gr.Checkbox(
                label="调试模式（保存识别/修改的JSON文件）", value=True
            )

        # ── 6. 处理按钮 ──
        btn_process = gr.Button("开始处理 → 生成文档", variant="primary", size="lg")

        # ═════════════════════════════════════
        # 回调：识别
        # ═════════════════════════════════════

        def _df_value_to_rows(df_value: Any) -> list[list]:
            """将 Dataframe 的值统一转换为可变的 list[list]。"""
            if df_value is None:
                return []
            import pandas as pd
            if isinstance(df_value, pd.DataFrame):
                result: list[list] = []
                for _, row in df_value.iterrows():
                    row_list: list = []
                    for v in row:
                        if pd.isna(v):
                            row_list.append("")
                        elif isinstance(v, bool):
                            row_list.append(v)
                        else:
                            row_list.append(str(v))
                    result.append(row_list)
                return result
            elif isinstance(df_value, list):
                return [list(row) for row in df_value]
            return []


        def _recognize(
            file: Any,
            smart_mode: bool,
            debug_mode: bool,
            _state_data: Any,
            _state_path: Any,
        ) -> tuple:
            """执行文档识别并更新界面组件。"""
            if file is None:
                raise gr.Error("请先上传Word文档")

            # Gradio 6.x 返回 NamedString(str子类)，旧版返回 dict
            if isinstance(file, dict):
                file_path: str = file.get("name") or file.get("path") or ""
            else:
                file_path = str(file)
            if not file_path or file_path == "None":
                raise gr.Error("无法获取文件路径")

            try:
                modify_data: ModifyInput = read_document(file_path, smart=smart_mode)
            except Exception as e:
                raise gr.Error(f"文档识别失败: {e}") from e

            # ── 调试：保存识别后的 JSON ──
            if debug_mode and modify_data is not None:
                stem = os.path.splitext(os.path.basename(file_path))[0]
                debug_path = os.path.join(os.path.dirname(file_path), f"{stem}_recognized.json")
                try:
                    with open(debug_path, "w", encoding="utf-8") as f:
                        json.dump(
                            modify_data.model_dump(mode="json"),
                            f,
                            ensure_ascii=False,
                            indent=2,
                        )
                except OSError:
                    pass  # 保存失败不阻断流程

            sections = modify_data.sections
            paragraphs = modify_data.paragraphs

            # ── 从节中提取页眉/页脚/页码信息 ──
            header_content = ""
            header_style = TextStyleInput()
            page_number_enabled = False
            page_format = "{PAGE}"
            page_align = "center"
            even_header_content = ""
            even_header_style = TextStyleInput()
            even_page_align = "center"
            odd_even_enabled = False

            if sections:
                sec = sections[0]
                if sec.header is not None:
                    header_content = sec.header.content or ""
                    header_style = sec.header.style or TextStyleInput()
                    if sec.header.page_number is not None:
                        page_number_enabled = sec.header.page_number.enabled
                        page_format = sec.header.page_number.format
                if sec.footer is not None and sec.footer.page_number is not None:
                    page_number_enabled = sec.footer.page_number.enabled
                    page_format = sec.footer.page_number.format
                if sec.even_page_header is not None:
                    even_header_content = sec.even_page_header.content or ""
                    even_header_style = sec.even_page_header.style or TextStyleInput()
                if sec.even_page_footer is not None and sec.even_page_footer.page_number is not None:
                    even_page_align = sec.even_page_footer.page_number.format or "center"
                odd_even_enabled = sec.enable_odd_even
                footer_distance_cm = sec.footer_distance_cm if sec.footer_distance_cm is not None else 1.5

            # ── 段落 Dataframe ──
            df_data: list[list] = []
            for p in paragraphs:
                df_data.append([
                    False,
                    str(p.index + 1),
                    _type_to_cn(p),
                    (p.text or "")[:80],
                ])

            # ── 提取样式 ──
            title_style = _find_first_style(paragraphs, ParagraphType.title) or TextStyleInput()
            subtitle_style = _find_first_style(paragraphs, ParagraphType.subtitle) or TextStyleInput()
            heading1_style = _find_first_style(paragraphs, ParagraphType.heading, heading_level=1) or TextStyleInput()
            heading2_style = _find_first_style(paragraphs, ParagraphType.heading, heading_level=2) or TextStyleInput()
            heading3_style = _find_first_style(paragraphs, ParagraphType.heading, heading_level=3) or TextStyleInput()
            heading4_style = _find_first_style(paragraphs, ParagraphType.heading, heading_level=4) or TextStyleInput()
            body_style = _find_first_style(paragraphs, ParagraphType.body) or TextStyleInput()

            return (
                gr.update(value=header_content),
                gr.update(value=_safe_choice(header_style.font_name, _FONT_CHOICES)),
                gr.update(value=_closest_font_size(header_style.font_size)),
                gr.update(value=_safe_choice(
                    header_style.alignment.value if header_style.alignment else None,
                    _ALIGN_CHOICES,
                ) or "center"),
                gr.update(value=page_number_enabled),
                gr.update(value=page_format),
                gr.update(value=page_align),
                gr.update(value=odd_even_enabled),
                gr.update(value=even_header_content, visible=odd_even_enabled),
                gr.update(value=_safe_choice(even_header_style.font_name, _FONT_CHOICES), visible=odd_even_enabled),
                gr.update(value=_closest_font_size(even_header_style.font_size), visible=odd_even_enabled),
                gr.update(value=_safe_choice(
                    even_header_style.alignment.value if even_header_style.alignment else None,
                    _ALIGN_CHOICES,
                ) or "center", visible=odd_even_enabled),
                gr.update(value=even_page_align, visible=odd_even_enabled),
                gr.update(value=footer_distance_cm),
                gr.update(value=2.54),
                gr.update(value=2.54),
                gr.update(value=2.54),
                gr.update(value=2.54),
                # 主标题 (5个)
                gr.update(value=_safe_choice(title_style.font_name, _FONT_CHOICES)),
                gr.update(value=_closest_font_size(title_style.font_size)),
                gr.update(value=title_style.bold if title_style.bold is not None else True),
                gr.update(value=_safe_choice(title_style.alignment.value if title_style.alignment else None, _ALIGN_CHOICES) or "center"),
                gr.update(value=_pt_to_chars(title_style.first_line_indent, title_style.font_size)),
                # 副标题 (5个)
                gr.update(value=_safe_choice(subtitle_style.font_name, _FONT_CHOICES)),
                gr.update(value=_closest_font_size(subtitle_style.font_size)),
                gr.update(value=subtitle_style.bold if subtitle_style.bold is not None else False),
                gr.update(value=_safe_choice(subtitle_style.alignment.value if subtitle_style.alignment else None, _ALIGN_CHOICES) or "center"),
                gr.update(value=_pt_to_chars(subtitle_style.first_line_indent, subtitle_style.font_size)),
                # 标题一 (5个)
                gr.update(value=_safe_choice(heading1_style.font_name, _FONT_CHOICES)),
                gr.update(value=_closest_font_size(heading1_style.font_size)),
                gr.update(value=heading1_style.bold if heading1_style.bold is not None else True),
                gr.update(value=_safe_choice(heading1_style.alignment.value if heading1_style.alignment else None, _ALIGN_CHOICES) or "center"),
                gr.update(value=_pt_to_chars(heading1_style.first_line_indent, heading1_style.font_size)),
                # 标题二 (5个)
                gr.update(value=_safe_choice(heading2_style.font_name, _FONT_CHOICES)),
                gr.update(value=_closest_font_size(heading2_style.font_size)),
                gr.update(value=heading2_style.bold if heading2_style.bold is not None else True),
                gr.update(value=_safe_choice(heading2_style.alignment.value if heading2_style.alignment else None, _ALIGN_CHOICES) or "center"),
                gr.update(value=_pt_to_chars(heading2_style.first_line_indent, heading2_style.font_size)),
                # 标题三 (5个)
                gr.update(value=_safe_choice(heading3_style.font_name, _FONT_CHOICES)),
                gr.update(value=_closest_font_size(heading3_style.font_size)),
                gr.update(value=heading3_style.bold if heading3_style.bold is not None else True),
                gr.update(value=_safe_choice(heading3_style.alignment.value if heading3_style.alignment else None, _ALIGN_CHOICES) or "center"),
                gr.update(value=_pt_to_chars(heading3_style.first_line_indent, heading3_style.font_size)),
                # 标题四 (5个)
                gr.update(value=_safe_choice(heading4_style.font_name, _FONT_CHOICES)),
                gr.update(value=_closest_font_size(heading4_style.font_size)),
                gr.update(value=heading4_style.bold if heading4_style.bold is not None else True),
                gr.update(value=_safe_choice(heading4_style.alignment.value if heading4_style.alignment else None, _ALIGN_CHOICES) or "center"),
                gr.update(value=_pt_to_chars(heading4_style.first_line_indent, heading4_style.font_size)),
                # 正文 (6个)
                gr.update(value=_safe_choice(body_style.font_name, _FONT_CHOICES)),
                gr.update(value=_closest_font_size(body_style.font_size)),
                gr.update(value=_safe_choice(body_style.alignment.value if body_style.alignment else None, _ALIGN_CHOICES) or "unknown"),
                gr.update(value=_pt_to_chars(body_style.first_line_indent, body_style.font_size) or 2),
                gr.update(value=body_style.line_spacing if body_style.line_spacing else 1.5),
                gr.update(value=df_data),
                modify_data.model_dump(mode="json"),
                file_path,
            )

        # ═════════════════════════════════════
        # 回调：处理生成
        # ═════════════════════════════════════

        def _process(
            file: Any,
            _state_data: Any,
            _state_path: Any,
            debug_mode: bool,
            # 基本配置
            header_content: str,
            header_font: str,
            header_size: str,
            header_align: str,
            page_enabled: bool,
            page_format_val: str,
            page_align_val: str,
            odd_even: bool,
            even_header: str,
            even_header_font: str,
            even_header_size: str,
            even_header_align: str,
            even_page_align_val: str,
            footer_distance_cm: float,
            margin_top: float,
            margin_bottom: float,
            margin_left: float,
            margin_right: float,
            # 样式配置
            title_font: str,
            title_size: str,
            title_bold: bool,
            title_align: str,
            title_indent_val: float,
            subtitle_font: str,
            subtitle_size: str,
            subtitle_bold: bool,
            subtitle_align: str,
            subtitle_indent_val: float,
            heading1_font: str,
            heading1_size: str,
            heading1_bold: bool,
            heading1_align: str,
            heading1_indent_val: float,
            heading2_font: str,
            heading2_size: str,
            heading2_bold: bool,
            heading2_align: str,
            heading2_indent_val: float,
            heading3_font: str,
            heading3_size: str,
            heading3_bold: bool,
            heading3_align: str,
            heading3_indent_val: float,
            heading4_font: str,
            heading4_size: str,
            heading4_bold: bool,
            heading4_align: str,
            heading4_indent_val: float,
            body_font: str,
            body_size: str,
            body_align: str,
            body_indent: float,
            body_line_spacing: float,
            # 段落
            df_value: Any,
        ) -> tuple:
            """收集界面参数，生成修改后的文档。"""
            if _state_data is None:
                raise gr.Error("请先识别文档")

            file_path = _state_path
            if not file_path or not os.path.exists(file_path):
                raise gr.Error("原始文件不存在，请重新上传")

            try:
                # ── 从下拉框解析字号 ──
                header_size_pt = _parse_font_size(header_size)
                title_size_pt = _parse_font_size(title_size)
                subtitle_size_pt = _parse_font_size(subtitle_size)
                heading1_size_pt = _parse_font_size(heading1_size)
                heading2_size_pt = _parse_font_size(heading2_size)
                heading3_size_pt = _parse_font_size(heading3_size)
                heading4_size_pt = _parse_font_size(heading4_size)
                body_size_pt = _parse_font_size(body_size)

                # ── 构建节数据 ──
                # 页眉（不带页码）
                header_style_obj = TextStyleInput(
                    font_name=header_font or None,
                    font_size=header_size_pt,
                    alignment=AlignmentType(header_align) if _safe_choice(header_align, _ALIGN_CHOICES) else None,
                )

                header_input = HeaderFooterInput(
                    content=header_content or None,
                    style=header_style_obj,
                )

                # 页脚（页码）
                page_number = PageNumberInput(
                    enabled=page_enabled,
                    format=page_format_val or "{PAGE}",
                )
                footer_style_obj = TextStyleInput(
                    alignment=AlignmentType(page_align_val) if _safe_choice(page_align_val, _ALIGN_CHOICES) else None,
                )
                footer_input = HeaderFooterInput(
                    content=None,
                    style=footer_style_obj,
                    page_number=page_number,
                )

                sections: list[SectionModify] = [
                    SectionModify(
                        index=0,
                        header=header_input,
                        footer=footer_input,
                        enable_odd_even=odd_even,
                        footer_distance_cm=footer_distance_cm,
                    )
                ]

                # 偶数页页眉/页脚（仅启用奇偶分页时）
                if odd_even and even_header:
                    even_header_size_pt = _parse_font_size(even_header_size)
                    even_header_style_obj = TextStyleInput(
                        font_name=even_header_font or None,
                        font_size=even_header_size_pt,
                        alignment=AlignmentType(even_header_align) if _safe_choice(even_header_align, _ALIGN_CHOICES) else None,
                    )
                    sections[0].even_page_header = HeaderFooterInput(
                        content=even_header,
                        style=even_header_style_obj,
                    )
                if odd_even and page_enabled:
                    even_footer_style = TextStyleInput(
                        alignment=AlignmentType(even_page_align_val) if _safe_choice(even_page_align_val, _ALIGN_CHOICES) else None,
                    )
                    sections[0].even_page_footer = HeaderFooterInput(
                        content=None,
                        style=even_footer_style,
                        page_number=PageNumberInput(
                            enabled=True,
                            format=page_format_val or "{PAGE}",
                        ),
                    )

                # 存储页边距到 JSON（writer 暂未支持）
                margins = {
                    "top_cm": margin_top,
                    "bottom_cm": margin_bottom,
                    "left_cm": margin_left,
                    "right_cm": margin_right,
                }

                # ── 从 Dataframe 构建段落列表 ──
                paragraphs: list[ParagraphModify] = []
                if df_value is not None:
                    import pandas as pd
                    if isinstance(df_value, pd.DataFrame):
                        for _, row in df_value.iterrows():
                            row_list = [str(row.iloc[1]) if pd.notna(row.iloc[1]) else "",
                                         str(row.iloc[2]) if pd.notna(row.iloc[2]) else "",
                                         str(row.iloc[3]) if pd.notna(row.iloc[3]) else ""]
                            _build_paragraph(paragraphs, *row_list)
                    elif isinstance(df_value, list):
                        for row in df_value:
                            if not row or len(row) < 4:
                                continue
                            _build_paragraph(paragraphs, *row[1:4])

                # 合并 margins 到 state_data（存储但不传给 modify_document）
                modify_state = json.loads(json.dumps(_state_data))
                modify_state["margins"] = margins

                # ── 构建完整的修改数据（含段落样式） ──
                title_style_obj = TextStyleInput(
                    font_name=title_font or None,
                    font_size=title_size_pt,
                    bold=title_bold,
                    alignment=AlignmentType(title_align) if _safe_choice(title_align, _ALIGN_CHOICES) else None,
                    first_line_indent=round(title_indent_val * (title_size_pt or 16), 1) if title_indent_val else None,
                )
                subtitle_style_obj = TextStyleInput(
                    font_name=subtitle_font or None,
                    font_size=subtitle_size_pt,
                    bold=subtitle_bold,
                    alignment=AlignmentType(subtitle_align) if _safe_choice(subtitle_align, _ALIGN_CHOICES) else None,
                    first_line_indent=round(subtitle_indent_val * (subtitle_size_pt or 16), 1) if subtitle_indent_val else None,
                )
                heading1_style_obj = TextStyleInput(
                    font_name=heading1_font or None,
                    font_size=heading1_size_pt,
                    bold=heading1_bold,
                    alignment=AlignmentType(heading1_align) if _safe_choice(heading1_align, _ALIGN_CHOICES) else None,
                    first_line_indent=round(heading1_indent_val * (heading1_size_pt or 16), 1) if heading1_indent_val else None,
                )
                heading2_style_obj = TextStyleInput(
                    font_name=heading2_font or None,
                    font_size=heading2_size_pt,
                    bold=heading2_bold,
                    alignment=AlignmentType(heading2_align) if _safe_choice(heading2_align, _ALIGN_CHOICES) else None,
                    first_line_indent=round(heading2_indent_val * (heading2_size_pt or 16), 1) if heading2_indent_val else None,
                )
                heading3_style_obj = TextStyleInput(
                    font_name=heading3_font or None,
                    font_size=heading3_size_pt,
                    bold=heading3_bold,
                    alignment=AlignmentType(heading3_align) if _safe_choice(heading3_align, _ALIGN_CHOICES) else None,
                    first_line_indent=round(heading3_indent_val * (heading3_size_pt or 16), 1) if heading3_indent_val else None,
                )
                heading4_style_obj = TextStyleInput(
                    font_name=heading4_font or None,
                    font_size=heading4_size_pt,
                    bold=heading4_bold,
                    alignment=AlignmentType(heading4_align) if _safe_choice(heading4_align, _ALIGN_CHOICES) else None,
                    first_line_indent=round(heading4_indent_val * (heading4_size_pt or 16), 1) if heading4_indent_val else None,
                )
                body_style_obj = TextStyleInput(
                    font_name=body_font or None,
                    font_size=body_size_pt,
                    alignment=AlignmentType(body_align) if _safe_choice(body_align, _ALIGN_CHOICES) else None,
                    first_line_indent=round(body_indent * (body_size_pt or 16), 1) if body_indent else None,
                    line_spacing=body_line_spacing,
                )

                for para in paragraphs:
                    if para.type == ParagraphType.title:
                        if para.style is None:
                            para.style = title_style_obj
                    elif para.type == ParagraphType.subtitle:
                        if para.style is None:
                            para.style = subtitle_style_obj
                    elif para.type == ParagraphType.heading:
                        if para.style is None:
                            if para.heading_level == 1:
                                para.style = heading1_style_obj
                            elif para.heading_level == 2:
                                para.style = heading2_style_obj
                            elif para.heading_level == 3:
                                para.style = heading3_style_obj
                            elif para.heading_level == 4:
                                para.style = heading4_style_obj
                            else:
                                para.style = heading1_style_obj
                    elif para.type in (ParagraphType.body, ParagraphType.other, ParagraphType.list):
                        if para.style is None:
                            para.style = body_style_obj

                modify_data = ModifyInput(
                    sections=sections,
                    paragraphs=paragraphs,
                )

                # ── 调试：保存修改前的 JSON ──
                if debug_mode:
                    stem = os.path.splitext(os.path.basename(file_path))[0]
                    config_path = os.path.join(os.path.dirname(file_path), f"{stem}_modified_config.json")
                    try:
                        with open(config_path, "w", encoding="utf-8") as f:
                            json.dump(
                                modify_data.model_dump(mode="json"),
                                f,
                                ensure_ascii=False,
                                indent=2,
                            )
                    except OSError:
                        pass

                # ── 输出路径 ──
                stem = os.path.splitext(os.path.basename(file_path))[0]
                output_dir = tempfile.gettempdir()
                output_path = os.path.join(output_dir, f"{stem}_modified.docx")

                modify_document(file_path, output_path, modify_data)

                if not os.path.exists(output_path):
                    raise gr.Error("文档生成失败")

                gr.Info(f"✅ 文档已生成: {os.path.basename(output_path)}")
                return gr.update(value=output_path, visible=True)

            except gr.Error:
                raise
            except Exception as e:
                raise gr.Error(f"处理失败: {e}") from e


        def _on_select_all(select_all: bool, df_value: Any) -> Any:
            """全选/取消全选：更新所有行的复选框列。"""
            rows = _df_value_to_rows(df_value)
            if not rows:
                return gr.update()
            for row in rows:
                row[0] = True if select_all else False
            return gr.update(value=rows)


        def _on_apply_batch(new_type: str, df_value: Any) -> Any:
            """批量应用：将选中段落的类型批量修改。"""
            if not new_type:
                raise gr.Error("请先选择要修改的目标类型")
            rows = _df_value_to_rows(df_value)
            if not rows:
                return gr.update()
            modified = False
            for row in rows:
                if row[0]:
                    if len(row) > 2:
                        row[2] = new_type
                        modified = True
            if not modified:
                raise gr.Error("请先勾选需要修改的段落")
            return gr.update(value=rows)


        _CN_TO_PARAGRAPH_TYPE: dict[str, tuple[ParagraphType, Optional[int]]] = {
            "主标题": (ParagraphType.title, None),
            "副标题": (ParagraphType.subtitle, None),
            "标题一": (ParagraphType.heading, 1),
            "标题二": (ParagraphType.heading, 2),
            "标题三": (ParagraphType.heading, 3),
            "标题四": (ParagraphType.heading, 4),
            "正文": (ParagraphType.body, None),
        }


        def _build_paragraph(
            paragraphs: list[ParagraphModify],
            idx_str: str,
            type_str: str,
            text: str,
        ) -> None:
            """从 Dataframe 行数据构建 ParagraphModify 并追加到列表。"""
            try:
                p_idx = int(idx_str) - 1
            except (ValueError, TypeError):
                return
            p_type: Optional[ParagraphType] = None
            heading_level: Optional[int] = None
            if type_str:
                mapping = _CN_TO_PARAGRAPH_TYPE
                if type_str.strip() in mapping:
                    p_type, heading_level = mapping[type_str.strip()]
                else:
                    p_type = ParagraphType.other
            paragraphs.append(
                ParagraphModify(
                    index=p_idx,
                    type=p_type,
                    heading_level=heading_level,
                    text=text or None,
                )
            )


        # ═════════════════════════════════════
        # 绑定事件
        # ═════════════════════════════════════

        # 识别回调
        _recog_outputs = [
            txt_header_content,
            dd_header_font,
            dd_header_size,
            dd_header_align,
            chk_page_number,
            txt_page_format,
            dd_page_align,
            chk_odd_even,
            txt_even_header,
            dd_even_header_font,
            dd_even_header_size,
            dd_even_header_align,
            dd_even_page_align,
            num_footer_distance,
            num_margin_top,
            num_margin_bottom,
            num_margin_left,
            num_margin_right,
            # 主标题 (5个)
            dd_title_font, dd_title_size, chk_title_bold, dd_title_align, num_title_indent,
            # 副标题 (5个)
            dd_subtitle_font, dd_subtitle_size, chk_subtitle_bold, dd_subtitle_align, num_subtitle_indent,
            # 标题一 (5个)
            dd_heading1_font, dd_heading1_size, chk_heading1_bold, dd_heading1_align, num_heading1_indent,
            # 标题二 (5个)
            dd_heading2_font, dd_heading2_size, chk_heading2_bold, dd_heading2_align, num_heading2_indent,
            # 标题三 (5个)
            dd_heading3_font, dd_heading3_size, chk_heading3_bold, dd_heading3_align, num_heading3_indent,
            # 标题四 (5个)
            dd_heading4_font, dd_heading4_size, chk_heading4_bold, dd_heading4_align, num_heading4_indent,
            # 正文 (5个)
            dd_body_font, dd_body_size, dd_body_align, num_body_indent, num_body_line_spacing,
            df_paragraphs,
            state_data,
            state_path,
        ]

        btn_recognize_raw.click(
            fn=lambda f, dm, sd, sp: _recognize(f, False, dm, sd, sp),
            inputs=[file_input, chk_debug_mode, state_data, state_path],
            outputs=_recog_outputs,
        )
        btn_recognize_smart.click(
            fn=lambda f, dm, sd, sp: _recognize(f, True, dm, sd, sp),
            inputs=[file_input, chk_debug_mode, state_data, state_path],
            outputs=_recog_outputs,
        )

        # 处理回调
        btn_process.click(
            fn=_process,
            inputs=[
                file_input,
                state_data,
                state_path,
                chk_debug_mode,
                txt_header_content,
                dd_header_font,
                dd_header_size,
                dd_header_align,
                chk_page_number,
                txt_page_format,
                dd_page_align,
                chk_odd_even,
                txt_even_header,
                dd_even_header_font,
                dd_even_header_size,
                dd_even_header_align,
                dd_even_page_align,
                num_footer_distance,
                num_margin_top,
                num_margin_bottom,
                num_margin_left,
                num_margin_right,
                dd_title_font, dd_title_size, chk_title_bold, dd_title_align, num_title_indent,
                dd_subtitle_font, dd_subtitle_size, chk_subtitle_bold, dd_subtitle_align, num_subtitle_indent,
                dd_heading1_font, dd_heading1_size, chk_heading1_bold, dd_heading1_align, num_heading1_indent,
                dd_heading2_font, dd_heading2_size, chk_heading2_bold, dd_heading2_align, num_heading2_indent,
                dd_heading3_font, dd_heading3_size, chk_heading3_bold, dd_heading3_align, num_heading3_indent,
                dd_heading4_font, dd_heading4_size, chk_heading4_bold, dd_heading4_align, num_heading4_indent,
                dd_body_font, dd_body_size, dd_body_align, num_body_indent, num_body_line_spacing,
                df_paragraphs,
            ],
            outputs=[file_output],
        )

        # 全选/取消全选
        chk_select_all.change(
            fn=_on_select_all,
            inputs=[chk_select_all, df_paragraphs],
            outputs=[df_paragraphs],
        )

        # 批量应用
        btn_apply.click(
            fn=_on_apply_batch,
            inputs=[dd_batch_type, df_paragraphs],
            outputs=[df_paragraphs],
        )

    return demo


# ─────────────────────────────────────────────
# 入口
# ─────────────────────────────────────────────

if __name__ == "__main__":
    main()


def main() -> None:
    """启动 GUI 服务。"""
    demo = build_ui()
    demo.launch(server_name="127.0.0.1", server_port=7860)
