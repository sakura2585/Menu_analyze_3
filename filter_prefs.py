# -*- coding: utf-8 -*-
"""主標籤篩選：已勾選標籤、顯示規則、區塊匯出勾選、匯出自訂範本，以及交叉表欄標籤持久化。"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from app_paths import project_data_dir

_PATH = project_data_dir() / "primary_filter_selection.json"

# 預設顯示規則（與舊版「序號+姓名+(大/小)」一致）
# disposable：勾選後，資料含「拋棄式」的列才對姓名加外框（不可只用「拋」字判斷）
DEFAULT_DISPLAY_RULE: dict[str, bool] = {
    "serial": True,
    "page_tag": False,
    "name": True,
    "size_label": True,
    "disposable": False,
}

# 匯出「自訂」排列：區塊標題／每一筆（可用 {tag}{count}{serial}{page}{name}{size}）
DEFAULT_EXPORT_CUSTOM_TEMPLATES: dict[str, str] = {
    "custom_block": "【{tag}】（{count} 筆）",
    "custom_row": "{serial}\t{page}\t{name}\t{size}",
}


def normalize_display_rule(d: Any) -> dict[str, bool]:
    """序號／人數標籤／資料頁標籤全不勾時，強制顯示姓名。"""
    if not isinstance(d, dict):
        d = {}
    s = bool(d.get("serial", True))
    p = bool(d.get("page_tag", False))
    n = bool(d.get("name", True))
    z = bool(d.get("size_label", True))
    disp = bool(d.get("disposable", False))
    if not s and not n and not z and not p:
        n = True
    return {"serial": s, "page_tag": p, "name": n, "size_label": z, "disposable": disp}


def normalize_export_templates(raw: Any) -> dict[str, str]:
    out = dict(DEFAULT_EXPORT_CUSTOM_TEMPLATES)
    if not isinstance(raw, dict):
        return out
    cb = raw.get("custom_block")
    cr = raw.get("custom_row")
    if isinstance(cb, str) and cb.strip():
        out["custom_block"] = cb.replace("\r\n", "\n").strip()
    if isinstance(cr, str) and cr.strip():
        out["custom_row"] = cr.replace("\r\n", "\n").strip()
    return out


def _normalize_crosstab_col_tags(raw: Any) -> list[str]:
    if not isinstance(raw, list):
        return []
    return [str(x).strip() for x in raw if str(x).strip()]


def _normalize_export_map(raw: Any) -> dict[str, bool]:
    if not isinstance(raw, dict):
        return {}
    out: dict[str, bool] = {}
    for k, v in raw.items():
        key = str(k).strip()
        if key:
            out[key] = bool(v)
    return out


def _load_raw() -> dict[str, Any]:
    if not _PATH.is_file():
        return {}
    try:
        with open(_PATH, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (json.JSONDecodeError, OSError):
        return {}


def load_filter_prefs() -> tuple[
    list[str], dict[str, dict[str, bool]], dict[str, bool], dict[str, str], list[str]
]:
    data = _load_raw()
    tags = data.get("selected_tags")
    if not isinstance(tags, list):
        tags = []
    clean_tags = [str(x).strip() for x in tags if str(x).strip()]
    rules_out: dict[str, dict[str, bool]] = {}
    raw_rules = data.get("display_rules")
    if isinstance(raw_rules, dict):
        for k, v in raw_rules.items():
            key = str(k).strip()
            if key:
                rules_out[key] = normalize_display_rule(v)
    export_inc = _normalize_export_map(data.get("block_export_include"))
    templates = normalize_export_templates(data.get("export_custom_templates"))
    crosstab_cols = _normalize_crosstab_col_tags(data.get("crosstab_col_tags"))
    return clean_tags, rules_out, export_inc, templates, crosstab_cols


def save_filter_prefs(
    tags: list[str],
    rules: dict[str, dict[str, bool] | Any],
    export_include: dict[str, bool] | None = None,
    export_templates: dict[str, str] | None = None,
    *,
    crosstab_col_tags: list[str] | None = None,
) -> None:
    data = _load_raw()
    clean_tags = [str(t).strip() for t in tags if str(t).strip()]
    clean_rules: dict[str, dict[str, bool]] = {}
    for k, v in (rules or {}).items():
        key = str(k).strip()
        if key:
            clean_rules[key] = normalize_display_rule(v)
    if export_include is None:
        export_include = _normalize_export_map(data.get("block_export_include"))
    clean_exp = {str(k).strip(): bool(v) for k, v in (export_include or {}).items() if str(k).strip()}
    if export_templates is None:
        export_templates = normalize_export_templates(data.get("export_custom_templates"))
    else:
        export_templates = normalize_export_templates(export_templates)
    if crosstab_col_tags is None:
        crosstab_clean = _normalize_crosstab_col_tags(data.get("crosstab_col_tags"))
    else:
        crosstab_clean = [str(x).strip() for x in crosstab_col_tags if str(x).strip()]
    out = {
        "selected_tags": clean_tags,
        "display_rules": clean_rules,
        "block_export_include": clean_exp,
        "export_custom_templates": export_templates,
        "crosstab_col_tags": crosstab_clean,
    }
    with open(_PATH, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)


def selection_file_path() -> Path:
    return _PATH


# --- 向後相容（僅讀寫 selected_tags，不建議新程式使用）---
def load_primary_filter_selection() -> list[str]:
    tags, _, _, _, _ = load_filter_prefs()
    return tags


def save_primary_filter_selection(tags: list[str]) -> None:
    _, rules, exp, tpl, ctab = load_filter_prefs()
    save_filter_prefs(tags, rules, exp, tpl, crosstab_col_tags=ctab)
