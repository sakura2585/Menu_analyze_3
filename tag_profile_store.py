# -*- coding: utf-8 -*-
"""
標籤組（profile）：多組標籤庫與主標籤篩選狀態分檔儲存；
切換作用中組別時，讀寫改為對應的 tag_database__*.json、primary_filter_selection__*.json。
舊版單一路徑 tag_database.json / primary_filter_selection.json 僅在首次讀取「小狀元」組時複製遷移。
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from app_paths import project_data_dir

PREFS_NAME = "tag_profile_prefs.json"

# 舊版單一檔遷移目標組（與預設第一組名稱一致）
LEGACY_MIGRATION_PROFILE_ID = "小狀元"

DEFAULT_PROFILE_DISPLAY_ORDER: tuple[str, ...] = ("小狀元", "組二", "組三")

_FORBIDDEN = '\\/:*?"<>|\n\r\t'

_state: dict[str, Any] | None = None


def sanitize_profile_id(name: str) -> str:
    s = (name or "").strip()
    for c in _FORBIDDEN:
        s = s.replace(c, "_")
    return s.strip(" ._")


def prefs_file_path() -> Path:
    return project_data_dir() / PREFS_NAME


def _default_state() -> dict[str, Any]:
    return {
        "active_profile_id": LEGACY_MIGRATION_PROFILE_ID,
        "profiles": list(DEFAULT_PROFILE_DISPLAY_ORDER),
    }


def _load_raw_file() -> dict[str, Any]:
    p = prefs_file_path()
    if not p.is_file():
        return {}
    try:
        with open(p, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (json.JSONDecodeError, OSError):
        return {}


def _dedupe_preserve(items: list[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for x in items:
        if x not in seen:
            seen.add(x)
            out.append(x)
    return out


def _persist(st: dict[str, Any]) -> None:
    p = prefs_file_path()
    p.parent.mkdir(parents=True, exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(
            {"active_profile_id": st["active_profile_id"], "profiles": st["profiles"]},
            f,
            ensure_ascii=False,
            indent=2,
        )


def _ensure_state() -> dict[str, Any]:
    global _state
    if _state is not None:
        return _state
    raw = _load_raw_file()
    created = False
    if not raw.get("profiles"):
        st = _default_state()
        created = True
    else:
        profs = [str(x).strip() for x in raw["profiles"] if str(x).strip()]
        profs = _dedupe_preserve(profs)
        if not profs:
            st = _default_state()
            created = True
        else:
            active = str(raw.get("active_profile_id") or "").strip()
            if active not in profs:
                active = profs[0]
            st = {"active_profile_id": active, "profiles": profs}
    _state = st
    if created:
        _persist(st)
    return _state


def get_active_profile_id() -> str:
    return str(_ensure_state()["active_profile_id"])


def list_profiles() -> list[str]:
    return list(_ensure_state()["profiles"])


def set_active_profile(profile_id: str) -> None:
    st = _ensure_state()
    pid = (profile_id or "").strip()
    if pid not in st["profiles"]:
        raise ValueError(f"未知的標籤組：{profile_id}")
    st["active_profile_id"] = pid
    _persist(st)


def add_profile(display_name: str) -> str:
    pid = sanitize_profile_id(display_name)
    if not pid:
        raise ValueError("請輸入有效的標籤組名稱（不可僅為標點或空白）。")
    st = _ensure_state()
    if pid in st["profiles"]:
        raise ValueError(f"已有「{pid}」這組，請換個名稱。")
    st["profiles"].append(pid)
    _persist(st)
    _ensure_empty_tag_file(pid)
    return pid


def rename_profile(old_id: str, new_display_name: str) -> str:
    """更名組別並重新命名磁碟上對應檔案。回傳新的組 id。"""
    old_id = (old_id or "").strip()
    new_id = sanitize_profile_id(new_display_name)
    if not old_id:
        raise ValueError("無效的標籤組。")
    if not new_id:
        raise ValueError("請輸入有效的標籤組名稱（不可僅為標點或空白）。")
    if new_id == old_id:
        return old_id
    st = _ensure_state()
    if old_id not in st["profiles"]:
        raise ValueError(f"未知的標籤組：{old_id}")
    if new_id in st["profiles"]:
        raise ValueError(f"已有「{new_id}」這組，請換個名稱。")

    root = project_data_dir()
    tag_old = root / f"tag_database__{old_id}.json"
    tag_new = root / f"tag_database__{new_id}.json"
    fil_old = root / f"primary_filter_selection__{old_id}.json"
    fil_new = root / f"primary_filter_selection__{new_id}.json"

    if tag_new.is_file() or fil_new.is_file():
        raise ValueError("目標檔名已存在，請換個名稱。")

    if tag_old.is_file():
        tag_old.rename(tag_new)
    else:
        _ensure_empty_tag_file(new_id)

    if fil_old.is_file():
        fil_old.rename(fil_new)

    i = st["profiles"].index(old_id)
    st["profiles"][i] = new_id
    if st["active_profile_id"] == old_id:
        st["active_profile_id"] = new_id
    _persist(st)
    return new_id


def remove_profile(profile_id: str) -> str:
    """
    刪除標籤組及其專用檔案，回傳刪除後的作用中組 id。
    至少需保留一組。
    """
    pid = (profile_id or "").strip()
    if not pid:
        raise ValueError("無效的標籤組。")
    st = _ensure_state()
    if pid not in st["profiles"]:
        raise ValueError(f"未知的標籤組：{pid}")
    if len(st["profiles"]) <= 1:
        raise ValueError("至少需保留一組標籤組。")

    root = project_data_dir()
    for fn in (f"tag_database__{pid}.json", f"primary_filter_selection__{pid}.json"):
        p = root / fn
        if p.is_file():
            try:
                p.unlink()
            except OSError as e:
                raise ValueError(f"無法刪除 {p.name}：{e}") from e

    st["profiles"] = [x for x in st["profiles"] if x != pid]
    if st["active_profile_id"] == pid:
        st["active_profile_id"] = st["profiles"][0]
    _persist(st)
    return str(st["active_profile_id"])


def _ensure_empty_tag_file(profile_id: str) -> None:
    p = project_data_dir() / f"tag_database__{profile_id}.json"
    if p.is_file():
        return
    p.parent.mkdir(parents=True, exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        json.dump({"hashtags": []}, f, ensure_ascii=False, indent=2)
