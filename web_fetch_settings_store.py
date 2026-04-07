# -*- coding: utf-8 -*-
from __future__ import annotations

import json
from dataclasses import asdict, dataclass

from app_paths import project_data_dir

_PATH = project_data_dir() / "web_fetch_settings.json"


@dataclass
class WebFetchSettings:
    profile_id: str = "little_champion_home"
    base_url: str = ""
    login_account: str = "a0824"
    login_password: str = ""
    # 若為空字串，UI／抓取 Flow 改採用內建 profile 預設 XPath
    source_xpath: str = ""
    date_xpath: str = ""
    date_prev_xpath: str = ""
    date_next_xpath: str = ""
    pre_click_xpath: str = ""
    # 表格為四欄時捨棄最後一欄（備註）；舊站備註常雜亂且不需分析
    omit_notes_column: bool = True
    # 抓取視窗尺寸（記住上次調整）
    ui_width: int = 760
    ui_height: int = 560


def load_web_fetch_settings() -> WebFetchSettings:
    try:
        data = json.loads(_PATH.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return WebFetchSettings()
    except json.JSONDecodeError:
        return WebFetchSettings()
    if not isinstance(data, dict):
        return WebFetchSettings()
    return WebFetchSettings(
        profile_id=str(data.get("profile_id") or "little_champion_home"),
        base_url=str(data.get("base_url") or ""),
        login_account=str(data.get("login_account") or "a0824"),
        login_password=str(data.get("login_password") or ""),
        source_xpath=str(data.get("source_xpath") or ""),
        date_xpath=str(data.get("date_xpath") or ""),
        date_prev_xpath=str(data.get("date_prev_xpath") or ""),
        date_next_xpath=str(data.get("date_next_xpath") or ""),
        pre_click_xpath=str(data.get("pre_click_xpath") or ""),
        omit_notes_column=bool(data.get("omit_notes_column", True)),
        ui_width=int(data.get("ui_width") or 760),
        ui_height=int(data.get("ui_height") or 560),
    )


def save_web_fetch_settings(settings: WebFetchSettings) -> None:
    _PATH.write_text(json.dumps(asdict(settings), ensure_ascii=False, indent=2), encoding="utf-8")

