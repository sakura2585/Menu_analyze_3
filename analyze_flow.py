# -*- coding: utf-8 -*-
"""
分析流程（Application 層第一版）。

後續可在 apply_row_enrichers 依 profile／欄位掛載不同分析器，不必再塞進 app.py。
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable

from analyze_field_prefs import AnalyzeFieldSet
from order_note_parser import ParsedRow, parse_bulk
from tag_store import register_hashtags


@dataclass
class AnalyzeRequest:
    pages: list[dict[str, Any]]


@dataclass
class AnalyzeResult:
    rows: list[ParsedRow]
    n_new: int
    n_total: int
    page_tag_error: str = ""
    hashtag_error: str = ""


# 列級擴充：預留給「姓名欄專用／備註欄專用」等分析器（目前為 no-op）
RowEnricher = Callable[[ParsedRow, str], None]


def _default_row_enrichers() -> tuple[RowEnricher, ...]:
    return ()


def apply_row_enrichers(row: ParsedRow, page_name: str, enrichers: tuple[RowEnricher, ...] | None = None) -> None:
    """解析完成後逐列呼叫；預設無額外規則。"""
    for fn in enrichers or _default_row_enrichers():
        fn(row, page_name)


def sync_page_names_to_hashtag_db(page_names: list[str]) -> str:
    try:
        register_hashtags(page_names)
        return ""
    except OSError as e:
        return str(e)


def parse_all_pages(
    pages: list[dict[str, Any]],
    *,
    enrichers: tuple[RowEnricher, ...] | None = None,
    field_set: AnalyzeFieldSet | None = None,
) -> list[ParsedRow]:
    all_rows: list[ParsedRow] = []
    for p in pages:
        raw = (p.get("text") or "").strip()
        if not raw:
            continue
        pname = str(p.get("name") or "未命名").strip() or "未命名"
        chunk = parse_bulk(p.get("text") or "", field_set=field_set)
        for r in chunk:
            r.source_page = pname
            r.tags.append({"category": "manual", "value": pname})
            apply_row_enrichers(r, pname, enrichers)
        all_rows.extend(chunk)
    return all_rows


def sync_hashtags_from_rows(rows: list[ParsedRow]) -> tuple[int, int, str]:
    hashtag_values = [t["value"] for r in rows for t in r.tags if t.get("category") == "hashtag"]
    try:
        n_new, n_total = register_hashtags(hashtag_values)
        return n_new, n_total, ""
    except OSError as e:
        return 0, 0, str(e)


def run_analyze(
    req: AnalyzeRequest,
    *,
    enrichers: tuple[RowEnricher, ...] | None = None,
    field_set: AnalyzeFieldSet | None = None,
) -> AnalyzeResult:
    pages = req.pages or []
    page_names = [str(p.get("name") or "未命名").strip() or "未命名" for p in pages]

    page_tag_error = sync_page_names_to_hashtag_db(page_names)
    all_rows = parse_all_pages(pages, enrichers=enrichers, field_set=field_set)
    n_new, n_total, hashtag_error = sync_hashtags_from_rows(all_rows)

    return AnalyzeResult(
        rows=all_rows,
        n_new=n_new,
        n_total=n_total,
        page_tag_error=page_tag_error,
        hashtag_error=hashtag_error,
    )
