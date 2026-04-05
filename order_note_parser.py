# -*- coding: utf-8 -*-
"""
訂餐／備註文字解析：欄位拆分、括號標籤、人數與常見關鍵字抽取。
比對時英文大小寫不敏感（casefold）；輸出之標籤中，以 no 開頭者統一為 NO 前綴，其餘中文／符號維持原樣。
詞典命中之 promo／swap 等則輸出詞典既定字串。
文中 `#詞`（至空白或下一個 # 前）會自動成為 category「hashtag」的標籤，供匯出或寫入資料庫。
"""

from __future__ import annotations

import csv
import io
import json
import re
from dataclasses import dataclass, field
from typing import Any

# 欄位至少要有：序號、姓名區、方案區、備註區（不足則補空字串）
EXPECTED_COLS = 4

# 括號內容（支援全形括號可再擴充）
PAREN_RE = re.compile(r"\(([^)]*)\)")
# 人數區間：2～3人、3～4人（～ 或 ~）
HEADCOUNT_RE = re.compile(r"(\d[～~]\d人)")


def headcount_size_label(headcount: str | None) -> str:
    """由人數字串判斷：2～3人（含 ~）→「小」，3～4人 →「大」；無法判斷回傳空字串。"""
    if not headcount:
        return ""
    m = HEADCOUNT_RE.search(headcount.strip())
    if not m:
        return ""
    token = m.group(1)
    if token.startswith("2"):
        return "小"
    if token.startswith("3"):
        return "大"
    return ""


def format_order_serial(serial: str) -> str:
    """序號為純數字時至少兩位（01、05）；非數字則原樣。"""
    s = str(serial).strip()
    if s.isdigit():
        return f"{int(s):02d}"
    return s
# 方案／加購常見片語
PROMO_SNIPPETS = ("優惠（月）", "優惠", "加購+", "加購")
SWAP_SNIPPETS = ("湯換菜", "湯換水果", "湯換飯", "海鮮換+", "NO湯換菜", "NO湯換水果", "NO湯換飯")
UTENSIL_SNIPPETS = ("拋棄式", "自備餐具")
RICE_RE = re.compile(r"白飯\s*\*\s*\d+|飯\s*\*\s*\d+", re.IGNORECASE)
# #標籤：# 後接至少一字元，至空白、Tab 或下一個 # 為止（不含 #）
_HASHTAG_RE = re.compile(r"#([^\s#\t]+)")


def _fold(s: str) -> str:
    """比對用：僅對有大小寫的字母（多為英文）做 casefold；中文與符號不變，須與關鍵字完全一致。"""
    return s.casefold()


_NO_PREFIX_RE = re.compile(r"(?i)no")


def normalize_leading_no(s: str) -> str:
    """
    輸出用：若字串開頭為 ASCII 的 no（不分大小寫），統一成 NO + 後綴（中文等維持原樣）。
    比對仍用 _fold；此函式僅正規化寫入標籤／匯出的文字。
    """
    s = s.strip()
    if not s:
        return s
    m = _NO_PREFIX_RE.match(s)
    if m:
        return "NO" + s[m.end() :]
    return s


def extract_hashtags(text: str) -> list[str]:
    """
    從整段文字擷取所有 #開頭詞句，去掉 # 後正規化（含 NO 前綴），同一列內英文大小寫視為重複則只保留一筆。
    回傳順序為文中出現順序。
    """
    seen: set[str] = set()
    out: list[str] = []
    for m in _HASHTAG_RE.finditer(text):
        raw_val = m.group(1).strip()
        if not raw_val:
            continue
        val = normalize_leading_no(raw_val)
        if not val:
            continue
        key = _fold(val)
        if key in seen:
            continue
        seen.add(key)
        out.append(val)
    return out


@dataclass
class ParsedRow:
    """單筆解析結果（可序列化給 GUI / API / DB）。"""

    line_no: int
    raw_line: str
    serial: str
    name_block: str
    plan_block: str
    notes_block: str
    customer_name: str
    paren_tags: list[str] = field(default_factory=list)
    headcount: str | None = None
    tags: list[dict[str, str]] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    # 來自「輸入與分析」多頁時的頁名（主標籤）；單頁或未設定則為 None
    source_page: str | None = None

    def to_dict(self) -> dict[str, Any]:
        d = {
            "line_no": self.line_no,
            "serial": self.serial,
            "customer_name": self.customer_name,
            "name_block": self.name_block,
            "plan_block": self.plan_block,
            "notes_block": self.notes_block,
            "headcount": self.headcount,
            "paren_tags": list(self.paren_tags),
            "tags": list(self.tags),
            "errors": list(self.errors),
            "source_page": self.source_page,
        }
        return d


def _split_columns(line: str) -> list[str]:
    """以 Tab 為主；若沒有 Tab，退回以兩個以上空白分割。"""
    line = line.rstrip("\r\n")
    if "\t" in line:
        parts = line.split("\t")
    else:
        parts = re.split(r"\s{2,}", line.strip())
    parts = [p.strip() for p in parts]
    while len(parts) < EXPECTED_COLS:
        parts.append("")
    if len(parts) > EXPECTED_COLS:
        # 多出的欄位併入最後一欄備註
        extra = parts[EXPECTED_COLS - 1 :]
        parts = parts[: EXPECTED_COLS - 1] + ["\t".join(extra)]
    return parts[:EXPECTED_COLS]


def _extract_customer_name(name_block: str) -> str:
    name_block = name_block.strip()
    if not name_block:
        return ""
    idx = name_block.find("(")
    if idx == -1:
        return name_block
    return name_block[:idx].strip()


def _headcount_from_text(text: str) -> str | None:
    m = HEADCOUNT_RE.search(text)
    return m.group(1) if m else None


def _add_tag(tags: list[dict[str, str]], category: str, value: str) -> None:
    value = value.strip()
    if not value:
        return
    tags.append({"category": category, "value": value})


def _scan_snippets(text: str, snippets: tuple[str, ...], category: str, tags: list[dict[str, str]]) -> None:
    folded_text = _fold(text)
    for s in snippets:
        if _fold(s) in folded_text:
            _add_tag(tags, category, s)


def _extract_no_like_segments(notes: str, tags: list[dict[str, str]]) -> None:
    """從備註中粗分 NO／排除敘述（以常見分隔符切塊）。"""
    if not notes.strip():
        return
    # 保留整段原文作為 searchable 備註
    _add_tag(tags, "notes_raw", notes.strip())
    chunks = re.split(r"[\.\-／、,，;；]+", notes)
    for ch in chunks:
        ch = ch.strip()
        if not ch:
            continue
        if _fold(ch).startswith("no"):
            _add_tag(tags, "restriction", normalize_leading_no(ch))
        elif "換" in ch and any(k in ch for k in ("湯", "海鮮", "菜", "水果", "飯")):
            _add_tag(tags, "swap_mention", ch)


def _build_tags(
    paren_tags: list[str],
    plan_block: str,
    notes_block: str,
    full_line: str,
) -> list[dict[str, str]]:
    tags: list[dict[str, str]] = []
    for p in paren_tags:
        _add_tag(tags, "paren", normalize_leading_no(p))

    _scan_snippets(plan_block, PROMO_SNIPPETS, "promo", tags)
    _scan_snippets(plan_block + notes_block, SWAP_SNIPPETS, "swap", tags)
    _scan_snippets(plan_block + notes_block, UTENSIL_SNIPPETS, "utensil", tags)

    for m in RICE_RE.finditer(plan_block + notes_block):
        _add_tag(tags, "rice", m.group(0).replace(" ", ""))

    _extract_no_like_segments(notes_block, tags)

    # 若備註很長且沒被細拆，full_line 仍可供後續 AI 使用
    if len(notes_block.strip()) > 30 and sum(1 for t in tags if t["category"] == "restriction") <= 1:
        _add_tag(tags, "needs_review", notes_block.strip()[:200])

    for ht in extract_hashtags(full_line):
        _add_tag(tags, "hashtag", ht)

    return tags


def parse_line(line: str, line_no: int) -> ParsedRow:
    raw = line.rstrip("\r\n")
    errors: list[str] = []
    cols = _split_columns(raw)
    serial, name_block, plan_block, notes_block = cols

    if not raw.strip():
        return ParsedRow(
            line_no=line_no,
            raw_line=raw,
            serial="",
            name_block="",
            plan_block="",
            notes_block="",
            customer_name="",
            errors=["空行"],
        )

    customer_name = _extract_customer_name(name_block)
    paren_tags = [normalize_leading_no(x.strip()) for x in PAREN_RE.findall(name_block) if x.strip()]

    combined = f"{name_block}\t{plan_block}\t{notes_block}"
    headcount = _headcount_from_text(combined)

    tags = _build_tags(paren_tags, plan_block, notes_block, raw)

    if headcount:
        # 避免重複：若 tags 裡尚無相同人數
        if not any(t["category"] == "headcount" for t in tags):
            tags.insert(0, {"category": "headcount", "value": headcount})

    if not serial and raw.strip():
        errors.append("序號欄為空")

    return ParsedRow(
        line_no=line_no,
        raw_line=raw,
        serial=serial.strip(),
        name_block=name_block,
        plan_block=plan_block,
        notes_block=notes_block,
        customer_name=customer_name,
        paren_tags=paren_tags,
        headcount=headcount,
        tags=tags,
        errors=errors,
    )


def parse_bulk(text: str) -> list[ParsedRow]:
    lines = text.splitlines()
    out: list[ParsedRow] = []
    for i, line in enumerate(lines, start=1):
        if not line.strip():
            continue
        out.append(parse_line(line, i))
    return out


def rows_to_json(rows: list[ParsedRow], ensure_ascii: bool = False) -> str:
    data = [r.to_dict() for r in rows]
    return json.dumps(data, ensure_ascii=ensure_ascii, indent=2)


def rows_to_csv_text(rows: list[ParsedRow]) -> str:
    """CSV 字串（UTF-8），供寫檔；含標籤摘要與備註預覽。"""
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["serial", "customer_name", "headcount", "source_page", "tag_summary", "notes_preview"])
    for r in rows:
        summary = ";".join(f'{t["category"]}:{t["value"]}' for t in r.tags[:30])
        if len(r.tags) > 30:
            summary += "…"
        prev = (r.notes_block or "")[:120]
        w.writerow([r.serial, r.customer_name, r.headcount or "", r.source_page or "", summary, prev])
    return buf.getvalue()
