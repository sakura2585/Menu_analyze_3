# -*- coding: utf-8 -*-
"""
# 標籤資料庫：tag_database.json
順序會保留（影響「主標籤篩選」區塊由上而下的排列）；分析時合併新標籤於清單末尾。
"""

from __future__ import annotations

import json
from pathlib import Path

from app_paths import project_data_dir

_DB_NAME = "tag_database.json"


def _db_path() -> Path:
    return project_data_dir() / _DB_NAME


def _dedupe_preserve_order(items: list[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for x in items:
        t = str(x).strip()
        if t and t not in seen:
            seen.add(t)
            out.append(t)
    return out


def _load_ordered(path: Path) -> list[str]:
    if not path.is_file():
        return []
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
    except (json.JSONDecodeError, OSError):
        return []
    items = data.get("hashtags")
    if not isinstance(items, list):
        return []
    return _dedupe_preserve_order([str(x) for x in items])


def _save_ordered(path: Path, tags: list[str]) -> None:
    clean = _dedupe_preserve_order(tags)
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"hashtags": clean}, f, ensure_ascii=False, indent=2)


def register_hashtags(values: list[str]) -> tuple[int, int]:
    """
    將一批標籤文字合併進資料庫（新項目接在清單後方，不變更既有順序）。
    回傳 (本次新増個數, 合併後總個數)。
    """
    path = _db_path()
    order = _load_ordered(path)
    seen = set(order)
    new_count = 0
    for v in values:
        v = v.strip()
        if v and v not in seen:
            order.append(v)
            seen.add(v)
            new_count += 1
    if new_count:
        _save_ordered(path, order)
    return new_count, len(order)


def database_path() -> Path:
    """給 UI 顯示用。"""
    return _db_path()


def list_hashtags() -> list[str]:
    """標籤庫內所有標籤，順序與檔案一致（供排序／主標籤篩選區塊順序）。"""
    return list(_load_ordered(_db_path()))


def save_hashtag_list(tags: list[str]) -> int:
    """依給定順序寫入標籤庫（去重，保留先出現者）。回傳寫入後筆數。"""
    path = _db_path()
    clean = _dedupe_preserve_order(tags)
    _save_ordered(path, clean)
    return len(clean)


def replace_hashtags_from_text(text: str) -> int:
    """
    以多行文字覆寫標籤庫（每行一個標籤，空白行略過，先出現者保留順序）。
    回傳寫入的不重複標籤數。
    """
    order: list[str] = []
    seen: set[str] = set()
    for line in text.splitlines():
        t = line.strip()
        if t and t not in seen:
            seen.add(t)
            order.append(t)
    _save_ordered(_db_path(), order)
    return len(order)
