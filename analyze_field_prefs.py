# -*- coding: utf-8 -*-
"""分析時要掃描的欄位開關（使用者可自訂）。"""
from __future__ import annotations

import json
from dataclasses import asdict, dataclass

from app_paths import project_data_dir

_PATH = project_data_dir() / "analyze_fields.json"


@dataclass
class AnalyzeFieldSet:
    """控制標籤規則要掃哪些欄位（序號欄不參與）。預設全開。"""

    name: bool = True
    plan: bool = True
    notes: bool = True
    full_line: bool = True

    @classmethod
    def all_on(cls) -> AnalyzeFieldSet:
        return cls(True, True, True, True)


def load_analyze_field_set() -> AnalyzeFieldSet:
    try:
        data = json.loads(_PATH.read_text(encoding="utf-8"))
    except FileNotFoundError:
        return AnalyzeFieldSet.all_on()
    except json.JSONDecodeError:
        return AnalyzeFieldSet.all_on()
    if not isinstance(data, dict):
        return AnalyzeFieldSet.all_on()
    return AnalyzeFieldSet(
        name=bool(data.get("name", True)),
        plan=bool(data.get("plan", True)),
        notes=bool(data.get("notes", True)),
        full_line=bool(data.get("full_line", True)),
    )


def save_analyze_field_set(fs: AnalyzeFieldSet) -> None:
    _PATH.write_text(json.dumps(asdict(fs), ensure_ascii=False, indent=2), encoding="utf-8")
