# -*- coding: utf-8 -*-
"""主標籤篩選：已勾選標籤、顯示規則、區塊匯出勾選、匯出自訂範本、交叉表欄標籤。
依標籤組分檔：primary_filter_selection__組名.json（預設組「小狀元」會自舊檔 primary_filter_selection.json 遷移）。"""

from __future__ import annotations

import json
import shutil
from pathlib import Path
from typing import Any

from app_paths import project_data_dir

_LEGACY_SELECTION = "primary_filter_selection.json"


def _selection_path_for_profile(profile_id: str | None) -> Path:
    from tag_profile_store import LEGACY_MIGRATION_PROFILE_ID, get_active_profile_id

    pid = (profile_id if profile_id is not None else get_active_profile_id()).strip()
    p = project_data_dir() / f"primary_filter_selection__{pid}.json"
    if not p.is_file():
        legacy = project_data_dir() / _LEGACY_SELECTION
        if legacy.is_file() and pid == LEGACY_MIGRATION_PROFILE_ID:
            try:
                shutil.copy2(legacy, p)
            except OSError:
                pass
    return p

# 預設顯示規則（與舊版「序號+姓名+(大/小)」一致）
# disposable：勾選後，資料含「拋棄式」的列才對姓名加外框（不可只用「拋」字判斷）
# utensil：勾選後，資料含「自備餐具」的列對姓名加外框（與拋棄式並列標示；同列兩者皆有時以拋棄式為準）
DEFAULT_DISPLAY_RULE: dict[str, bool] = {
    "serial": True,
    "page_tag": False,
    "name": True,
    "size_label": True,
    "disposable": False,
    "utensil": False,
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
    ut = bool(d.get("utensil", False))
    if not s and not n and not z and not p:
        n = True
    return {
        "serial": s,
        "page_tag": p,
        "name": n,
        "size_label": z,
        "disposable": disp,
        "utensil": ut,
    }


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


def _load_raw(profile_id: str | None = None) -> dict[str, Any]:
    path = _selection_path_for_profile(profile_id)
    if not path.is_file():
        return {}
    try:
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (json.JSONDecodeError, OSError):
        return {}


def _normalize_order_list(raw: Any) -> list[str]:
    if not isinstance(raw, list):
        return []
    out: list[str] = []
    seen: set[str] = set()
    for x in raw:
        s = str(x).strip()
        if s and s not in seen:
            seen.add(s)
            out.append(s)
    return out


def load_filter_prefs(profile_id: str | None = None) -> tuple[
    list[str],
    dict[str, dict[str, bool]],
    dict[str, bool],
    dict[str, str],
    list[str],
    list[str],
    list[str],
    list[str],
]:
    data = _load_raw(profile_id)
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
    primary_tag_order = _normalize_order_list(data.get("primary_tag_order"))
    crosstab_tag_order = _normalize_order_list(data.get("crosstab_tag_order"))
    export_tag_order = _normalize_order_list(data.get("export_tag_order"))
    return (
        clean_tags,
        rules_out,
        export_inc,
        templates,
        crosstab_cols,
        primary_tag_order,
        crosstab_tag_order,
        export_tag_order,
    )


def save_filter_prefs(
    tags: list[str],
    rules: dict[str, dict[str, bool] | Any],
    export_include: dict[str, bool] | None = None,
    export_templates: dict[str, str] | None = None,
    *,
    crosstab_col_tags: list[str] | None = None,
    primary_tag_order: list[str] | None = None,
    crosstab_tag_order: list[str] | None = None,
    export_tag_order: list[str] | None = None,
    profile_id: str | None = None,
) -> None:
    data = _load_raw(profile_id)
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
    if primary_tag_order is None:
        primary_order_clean = _normalize_order_list(data.get("primary_tag_order"))
    else:
        primary_order_clean = _normalize_order_list(primary_tag_order)
    if crosstab_tag_order is None:
        crosstab_order_clean = _normalize_order_list(data.get("crosstab_tag_order"))
    else:
        crosstab_order_clean = _normalize_order_list(crosstab_tag_order)
    if export_tag_order is None:
        export_order_clean = _normalize_order_list(data.get("export_tag_order"))
    else:
        export_order_clean = _normalize_order_list(export_tag_order)
    out = {
        "selected_tags": clean_tags,
        "display_rules": clean_rules,
        "block_export_include": clean_exp,
        "export_custom_templates": export_templates,
        "crosstab_col_tags": crosstab_clean,
        "primary_tag_order": primary_order_clean,
        "crosstab_tag_order": crosstab_order_clean,
        "export_tag_order": export_order_clean,
    }
    path = _selection_path_for_profile(profile_id)
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)


def selection_file_path(profile_id: str | None = None) -> Path:
    return _selection_path_for_profile(profile_id)


# --- 向後相容（僅讀寫 selected_tags，不建議新程式使用）---
def load_primary_filter_selection() -> list[str]:
    tags, _, _, _, _, _, _, _ = load_filter_prefs()
    return tags


def save_primary_filter_selection(tags: list[str]) -> None:
    _, rules, exp, tpl, ctab, pord, cord, eord = load_filter_prefs()
    save_filter_prefs(
        tags,
        rules,
        exp,
        tpl,
        crosstab_col_tags=ctab,
        primary_tag_order=pord,
        crosstab_tag_order=cord,
        export_tag_order=eord,
    )
