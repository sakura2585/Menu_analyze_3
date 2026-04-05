# -*- coding: utf-8 -*-
"""輸入與分析：多頁文字與目前頁／名單檢視選項持久化。"""

from __future__ import annotations

import json
import uuid
from pathlib import Path

from app_paths import project_data_dir

_PATH = project_data_dir() / "input_pages.json"

ROSTER_VIEW_ALL = "【全部頁】"


def _new_page_id() -> str:
    return uuid.uuid4().hex[:12]


def allocate_page_id() -> str:
    """產生新輸入頁 id（供 app 新增頁時使用）。"""
    return _new_page_id()


def default_pages_state() -> dict:
    pid = _new_page_id()
    return {
        "current_page_id": pid,
        "roster_view": ROSTER_VIEW_ALL,
        "pages": [{"id": pid, "name": "預設頁", "text": ""}],
    }


def load_input_pages_state() -> dict:
    if not _PATH.is_file():
        return default_pages_state()
    try:
        with open(_PATH, encoding="utf-8") as f:
            data = json.load(f)
    except (json.JSONDecodeError, OSError):
        return default_pages_state()
    pages = data.get("pages")
    if not isinstance(pages, list) or not pages:
        return default_pages_state()
    norm = []
    for p in pages:
        if not isinstance(p, dict):
            continue
        pid = str(p.get("id") or "").strip() or _new_page_id()
        name = str(p.get("name") or "").strip() or "未命名"
        text = str(p.get("text") or "")
        norm.append({"id": pid, "name": name, "text": text})
    if not norm:
        return default_pages_state()
    cur = str(data.get("current_page_id") or "").strip()
    if not any(p["id"] == cur for p in norm):
        cur = norm[0]["id"]
    rv = str(data.get("roster_view") or ROSTER_VIEW_ALL).strip() or ROSTER_VIEW_ALL
    return {"current_page_id": cur, "roster_view": rv, "pages": norm}


def save_input_pages_state(state: dict) -> None:
    pages = state.get("pages") or []
    clean_pages = []
    for p in pages:
        if not isinstance(p, dict):
            continue
        clean_pages.append(
            {
                "id": str(p.get("id") or _new_page_id()),
                "name": str(p.get("name") or "未命名").strip() or "未命名",
                "text": str(p.get("text") or ""),
            }
        )
    if not clean_pages:
        d = default_pages_state()
        clean_pages = d["pages"]
        cur = d["current_page_id"]
        rv = d["roster_view"]
    else:
        cur = str(state.get("current_page_id") or clean_pages[0]["id"])
        if not any(p["id"] == cur for p in clean_pages):
            cur = clean_pages[0]["id"]
        rv = str(state.get("roster_view") or ROSTER_VIEW_ALL)
    out = {"current_page_id": cur, "roster_view": rv, "pages": clean_pages}
    with open(_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
