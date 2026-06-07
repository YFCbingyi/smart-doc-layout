"""Word文档识别与修改的命令行入口。

用法:
    # 识别文档
    python -m word_processor read <input.docx> [output.json]

    # 修改文档
    python -m word_processor modify <input.docx> <modify.json> [output.docx]
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

from word_processor.models import ModifyInput
from word_processor.reader import read_document
from word_processor.writer import modify_document


def _json_serialize(obj: object) -> object:
    """处理Pydantic模型的JSON序列化。"""
    if isinstance(obj, ModifyInput):
        return obj.model_dump(mode="json")
    return str(obj)


def cmd_read(args: list[str]) -> int:
    """执行文档识别命令。"""
    if len(args) < 1:
        print("用法: python -m word_processor read <input.docx> [output.json]")
        return 1

    input_path = args[0]
    output_path = args[1] if len(args) > 1 else None

    try:
        print(f"正在读取文档: {input_path}")
        data = read_document(input_path)

        json_data = data.model_dump(mode="json", exclude_none=True)

        if output_path:
            with open(output_path, "w", encoding="utf-8") as f:
                json.dump(json_data, f, ensure_ascii=False, indent=2)
            print(f"识别结果已保存到: {output_path}")
        else:
            json.dump(json_data, sys.stdout, ensure_ascii=False, indent=2)
            print()

        print(f"完成: 共识别 {len(data.paragraphs)} 个段落, {len(data.sections)} 个节")

    except FileNotFoundError as e:
        print(f"错误: {e}", file=sys.stderr)
        return 1
    except ValueError as e:
        print(f"错误: {e}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"未知错误: {e}", file=sys.stderr)
        return 1

    return 0


def cmd_modify(args: list[str]) -> int:
    """执行文档修改命令。"""
    if len(args) < 2:
        print("用法: python -m word_processor modify <input.docx> <modify.json> [output.docx]")
        return 1

    input_path = args[0]
    modify_json_path = args[1]
    output_path = args[2] if len(args) > 2 else input_path

    try:
        print(f"正在读取源文档: {input_path}")
        print(f"正在读取修改配置: {modify_json_path}")

        with open(modify_json_path, "r", encoding="utf-8") as f:
            modify_dict = json.load(f)

        modify_data = ModifyInput(**modify_dict)

        modify_document(input_path, output_path, modify_data)
        print(f"修改完成，结果已保存到: {output_path}")

    except FileNotFoundError as e:
        print(f"错误: {e}", file=sys.stderr)
        return 1
    except json.JSONDecodeError as e:
        print(f"JSON格式错误: {e}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"错误: {e}", file=sys.stderr)
        return 1

    return 0


def main() -> int:
    """主入口。"""
    if len(sys.argv) < 2:
        print("用法:")
        print("  python -m word_processor read <input.docx> [output.json]")
        print("  python -m word_processor modify <input.docx> <modify.json> [output.docx]")
        return 1

    command = sys.argv[1]
    args = sys.argv[2:]

    if command == "read":
        return cmd_read(args)
    elif command == "modify":
        return cmd_modify(args)
    else:
        print(f"未知命令: {command}")
        print("可用命令: read, modify")
        return 1


if __name__ == "__main__":
    sys.exit(main())
