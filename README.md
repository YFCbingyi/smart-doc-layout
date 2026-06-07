# Smart Doc Layout — Word 文档处理器

一个基于 Python 的 Word 文档自动化处理工具，支持识别文档结构和样式，并根据配置批量修改文档内容、样式、页眉页脚和页码。

## 功能特性

- **文档识别** — 读取 `.docx` 文件，提取段落类型、样式、页眉/页脚、页码等结构化信息
- **智能识别** — 自动识别中文四级标题（一、/（一）/1./(1)）和主标题/副标题
- **样式修改** — 批量修改段落字体、字号、颜色、对齐、行距、缩进等样式
- **页码处理** — 添加或移除页码，支持自定义格式（`{PAGE}` / `{NUMPAGES}`）
- **页眉页脚** — 修改页眉/页脚文本和样式
- **奇偶分页** — 支持奇偶页不同的页眉/页脚设置
- **页脚距设置** — 控制页脚距页面底端的距离
- **图形界面** — 基于 Gradio 的 Web 界面，可视化操作
- **命令行工具** — 支持命令行批量处理

## 安装

```bash
pip install -r requirements.txt
```

依赖项：python-docx、pydantic、lxml、gradio

## 快速开始

### GUI 模式

```bash
python -m word_processor gui
```

浏览器访问 `http://127.0.0.1:7860`，上传 Word 文档后即可识别和修改。

### 命令行模式

```bash
# 识别文档结构
python -m word_processor read input.docx output.json

# 根据配置修改文档
python -m word_processor modify input.docx config.json output.docx
```

### Python API

```python
from word_processor import read_document, modify_document

# 识别文档
data = read_document("input.docx", smart=True)

# 查看段落信息
for p in data.paragraphs:
    print(f"[{p.index}] {p.type}: {p.text[:50]}")

# 修改文档
modify_document("input.docx", "output.docx", data)
```

## 命令行文档

```
python -m word_processor <command> [args]

命令:
  read    识别文档，输出 JSON 格式的结构化数据
  modify  根据 JSON 配置修改文档
  gui     启动 Web 图形界面

read:
  python -m word_processor read <input.docx> [output.json]

modify:
  python -m word_processor modify <input.docx> <modify.json> [output.docx]
```

- `read` 命令识别文档后可将结果保存为 JSON，方便查看或手动编辑
- `modify` 命令读取 JSON 配置并对文档进行修改，未指定的字段保持不变
- `gui` 命令启动 Gradio Web 服务，默认监听 `127.0.0.1:7860`

## Python API

### `read_document(file_path, smart=True)`

| 参数 | 类型 | 说明 |
|------|------|------|
| `file_path` | `str` | `.docx` 文件路径 |
| `smart` | `bool` | 是否启用智能识别（四级标题、主副标题），默认 `True` |

**返回**: `ModifyInput` 对象，包含 `sections`（节列表）和 `paragraphs`（段落列表）。

### `modify_document(file_path, output_path, modify_data)`

| 参数 | 类型 | 说明 |
|------|------|------|
| `file_path` | `str` | 源文档路径 |
| `output_path` | `str` | 输出文档路径 |
| `modify_data` | `ModifyInput` | 修改参数模型 |

## GUI 界面说明

| 面板 | 功能 |
|------|------|
| 文件上传/下载 | 上传 `.docx` 文件，下载处理后的文档 |
| 识别按钮 | "原文格式识别"（基于样式）/"智能辅助识别"（含内容分析） |
| 基本配置 | 页眉内容/字体/字号/对齐、页码启用/格式/对齐、奇偶分页、页边距、页脚距 |
| 样式配置 | 主标题、副标题、标题一~四、正文的字体/字号/加粗/对齐/缩进/行距 |
| 段落列表 | 勾选段落 + 批量修改类型（支持全选/取消全选） |
| 调试模式 | 保存识别和修改的 JSON 配置文件 |

## 数据模型

核心数据模型基于 Pydantic v2，识别输出和修改输入共用同一套模型。

| 模型 | 说明 |
|------|------|
| `ModifyInput` | 顶层模型，包含 `sections` 和 `paragraphs` |
| `SectionModify` | 节数据：页眉/页脚、页码、奇偶分页、页脚距 |
| `ParagraphModify` | 段落数据：索引、类型、标题级别、文本、样式 |
| `HeaderFooterInput` | 页眉/页脚：内容、样式、页码设置 |
| `PageNumberInput` | 页码：启用状态、格式模板 |
| `TextStyleInput` | 样式：字体、字号、加粗、斜体、颜色、对齐、行距、缩进 |

段落类型枚举：`title`、`subtitle`、`heading`、`body`、`list`、`table`、`other`

## 项目结构

```
smart-doc-layout/
├── word_processor/          # 核心包
│   ├── __init__.py          # 包入口，导出主要 API
│   ├── __main__.py          # 运行入口 (python -m word_processor)
│   ├── cli.py               # 命令行接口
│   ├── gui.py               # Gradio Web GUI
│   ├── models.py            # Pydantic 数据模型
│   ├── reader.py            # 文档读取/识别
│   ├── writer.py            # 文档修改/写入
│   └── page_number.py       # 页码 OXML 操作
├── tests/                   # 测试
│   ├── create_sample.py     # 示例文档生成
│   ├── test_reader.py       # 读取测试
│   └── test_writer.py       # 写入测试
├── examples/                # 示例文件
├── requirements.txt         # 依赖
├── README.md                # 本文件
└── .trae/specs/             # 功能规格文档
```

## 开发指南

### 设计说明

- **识别与修改模型统一**：`ModifyInput` 同时作为识别输出和修改输入，形成"识别 → 编辑 → 写入"的工作闭环
- **基于原文档修改**：非从头生成文档，保留未修改的段落和样式
- **OXML 层操作**：页码等复杂功能通过底层 Open XML 操作实现
- **智能识别**：通过正则匹配文本内容识别中文标题，不依赖 Word 样式

## 许可证

MIT License
