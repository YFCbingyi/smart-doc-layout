"""模板管理工具 — 保存/加载/重命名/删除排版配置模板。

模板以 JSON 文件形式存放在项目根目录下的 `config/` 文件夹中。
"""

from __future__ import annotations

import json
import os
from typing import Any


def _config_dir() -> str:
    """返回模板配置目录的绝对路径（项目根目录下的 config/）。"""
    # template.py 位于 word_processor/ 下，config/ 在项目根目录
    package_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(package_dir)
    config_dir = os.path.join(project_root, "config")
    os.makedirs(config_dir, exist_ok=True)
    return config_dir


def list_templates() -> list[str]:
    """扫描 config/ 目录，返回所有模板名称列表（不含 .json 扩展名）。"""
    templates: list[str] = []
    cfg_dir = _config_dir()
    if not os.path.isdir(cfg_dir):
        return templates
    for fname in os.listdir(cfg_dir):
        if fname.endswith(".json") and fname != ".gitkeep":
            name = fname[:-5]  # 去掉 .json
            templates.append(name)
    templates.sort()
    return templates


def save_template(name: str, data: dict[str, Any]) -> str:
    """将配置数据保存为模板文件。

    Args:
        name: 模板名称（不含扩展名）
        data: 配置数据字典

    Returns:
        保存后的模板名称
    """
    cfg_dir = _config_dir()
    filepath = os.path.join(cfg_dir, f"{name}.json")
    with open(filepath, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return name


def load_template(name: str) -> dict[str, Any]:
    """读取指定模板的配置数据。

    Args:
        name: 模板名称（不含扩展名）

    Returns:
        配置数据字典

    Raises:
        FileNotFoundError: 模板不存在
    """
    cfg_dir = _config_dir()
    filepath = os.path.join(cfg_dir, f"{name}.json")
    if not os.path.isfile(filepath):
        raise FileNotFoundError(f"模板 '{name}' 不存在")
    with open(filepath, "r", encoding="utf-8") as f:
        return json.load(f)


def rename_template(old_name: str, new_name: str) -> str:
    """重命名模板文件。

    Args:
        old_name: 原模板名称（不含扩展名）
        new_name: 新模板名称（不含扩展名）

    Returns:
        新模板名称

    Raises:
        FileNotFoundError: 原模板不存在
        FileExistsError: 新名称已存在
    """
    cfg_dir = _config_dir()
    old_path = os.path.join(cfg_dir, f"{old_name}.json")
    new_path = os.path.join(cfg_dir, f"{new_name}.json")
    if not os.path.isfile(old_path):
        raise FileNotFoundError(f"模板 '{old_name}' 不存在")
    if os.path.exists(new_path):
        raise FileExistsError(f"模板 '{new_name}' 已存在")
    os.rename(old_path, new_path)
    return new_name


def delete_template(name: str) -> None:
    """删除模板文件。

    Args:
        name: 模板名称（不含扩展名）

    Raises:
        FileNotFoundError: 模板不存在
    """
    cfg_dir = _config_dir()
    filepath = os.path.join(cfg_dir, f"{name}.json")
    if not os.path.isfile(filepath):
        raise FileNotFoundError(f"模板 '{name}' 不存在")
    os.remove(filepath)
