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

from analyze_field_prefs import AnalyzeFieldSet

# 欄位：排序／序號、姓名、品項（方案）、備註（與小狀元表頭一致）
EXPECTED_COLS = 4

# 抓取結果常見首行：僅日期（與資料列區隔）
_META_DATE_ONLY_RE = re.compile(r"^\s*\d{4}年\d{1,2}月\d{1,2}日\s*$")

# 括號內容（支援全形括號可再擴充）
PAREN_RE = re.compile(r"\(([^)]*)\)")
# 人數區間：2～3人、3～4人（全形～、半形 ~、波浪號 〜 等）
HEADCOUNT_RE = re.compile(r"(\d(?:[～~〜]|\u301c)\d人)")
# 品項欄僅「○～○人」、與「○～○人優惠（月）」視為同一類欄位內容
_PLAN_HEADCOUNT_ONLY_RE = re.compile(r"^\d(?:[～~〜]|\u301c)\d人$")


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
# 方案／加購常見片語（品項欄經 _normalize_plan_for_promo_match 後比對）
PROMO_SNIPPETS = ("優惠（月）", "優惠", "加購+", "加購")
SWAP_SNIPPETS = ("湯換菜", "湯換水果", "湯換飯", "海鮮換+", "NO湯換菜", "NO湯換水果", "NO湯換飯")
UTENSIL_SNIPPETS = ("拋棄式", "自備餐具")
RICE_RE = re.compile(r"白飯\s*\*\s*\d+|飯\s*\*\s*\d+", re.IGNORECASE)
# #標籤：# 後接至少一字元，至空白／Tab／全形‧不休空白／下一個 # 為止（不含 #）
_HASHTAG_RE = re.compile(r"#([^\s#\t\u00a0\u3000]+)")


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


def merge_hashtags_from_fields(*fields: str) -> list[str]:
    """多個欄位各自抽 hashtag 後合併；跨欄依欄位順序、欄內維持原文順序，fold 去重。"""
    seen: set[str] = set()
    out: list[str] = []
    for field in fields:
        for ht in extract_hashtags(field or ""):
            k = _fold(ht)
            if k in seen:
                continue
            seen.add(k)
            out.append(ht)
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


def _is_skippable_non_data_line(line: str) -> bool:
    """日期單行、表頭列等不應當成一筆訂單。"""
    s = line.strip()
    if not s:
        return True
    if _META_DATE_ONLY_RE.match(s):
        return True
    if "\t" in s:
        parts = [p.strip() for p in s.split("\t")]
        if len(parts) >= 2:
            h0, h1 = parts[0], parts[1]
            if h0 in ("排序", "序號", "編號", "No", "NO", "No.") and h1 in ("姓名", "名字", "名稱"):
                return True
    return False


def _normalize_serial_cell(raw: str) -> str:
    """排序欄：去掉常見星號／圖示前綴，保留數字序號。"""
    s = (raw or "").strip()
    if not s:
        return ""
    s = re.sub(r"^[\s\u2605\u2730\u2726\u2736\u2734\u2739\*\u2022]+", "", s)
    return s.strip()


def _split_columns(line: str) -> list[str]:
    """以 Tab 為主（與 Selenium 擷取 table 一致）；若沒有 Tab，退回以兩個以上空白分割。"""
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
    out = parts[:EXPECTED_COLS]
    out[0] = _normalize_serial_cell(out[0])
    return out


def _extract_customer_name(name_block: str) -> str:
    name_block = name_block.strip()
    if not name_block:
        return ""
    idx = name_block.find("(")
    base = name_block if idx == -1 else name_block[:idx].strip()
    if not base:
        return ""
    # 若姓名後以空白接常見備註片段（如 NO蝦.蛋.自備餐具），只取第一段作為姓名。
    # 這可避免 customer_name 吃進限制條件，並讓後續以 name_block 補掃標籤。
    m = re.match(r"^(\S+)\s+(.+)$", base)
    if not m:
        return base
    first, tail = m.group(1), m.group(2).strip()
    if (
        re.match(r"(?i)^no", tail)
        or any(ch in tail for ch in ".。／/、,，;；|｜:：")
        or any(k in tail for k in ("自備餐具", "拋棄式", "換", "不", "不要"))
    ):
        return first
    return base


def _headcount_from_text(text: str) -> str | None:
    m = HEADCOUNT_RE.search(text)
    return m.group(1) if m else None


def _normalize_plan_for_promo_match(plan_block: str) -> str:
    """半形括號改全形、去掉空白，讓「優惠 ( 月 )」與「優惠（月）」能命中同一組片語。"""
    s = (plan_block or "").strip()
    s = s.replace("(", "（").replace(")", "）")
    s = re.sub(r"\s+", "", s)
    return s


def _infer_promo_when_plan_is_headcount_only(plan_norm: str, tags: list[dict[str, str]]) -> None:
    """品項欄只有「2～3人」「3～4人」等時，與帶「優惠」的寫法視為同欄，補上 promo 優惠。"""
    if any(t.get("category") == "promo" for t in tags):
        return
    if plan_norm and _PLAN_HEADCOUNT_ONLY_RE.match(plan_norm):
        _add_tag(tags, "promo", "優惠")


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


def _extract_no_like_segments(
    text: str,
    tags: list[dict[str, str]],
    *,
    verbatim_category: str | None = "notes_raw",
) -> None:
    """
    粗分 NO／排除敘述（以常見分隔符切塊）。
    verbatim_category：若指定則寫入該欄原文（備註用 notes_raw、姓名欄用 name_column_raw）；
    僅掃描 restriction 時可傳 None。
    """
    if not text.strip():
        return
    if verbatim_category:
        _add_tag(tags, verbatim_category, text.strip())
    # 略增分隔符（含半形/全形句點、直線、冒號、斜線、換行）有利長句切塊掃描
    chunks = re.split(r"[\.。\-／、,，;；\|｜:：/\r\n]+", text)
    for ch in chunks:
        ch = ch.strip()
        if not ch:
            continue
        # 允許「姓名 NO蝦」這種 NO 不在片段起首的格式，抽出片段中的 NO詞。
        m_no = re.search(r"(?i)(?:^|\s)(no[^\s,，;；:：/／\|｜\.。]+)", ch)
        if m_no:
            _add_tag(tags, "restriction", normalize_leading_no(m_no.group(1)))
        if "換" in ch and any(k in ch for k in ("湯", "海鮮", "菜", "水果", "飯")):
            _add_tag(tags, "swap_mention", ch)


def _hashtag_fold_set(tags: list[dict[str, str]]) -> set[str]:
    return {_fold(t.get("value", "")) for t in tags if t.get("category") == "hashtag" and t.get("value")}


def _apply_tag_library(search_text: str, tags: list[dict[str, str]]) -> None:
    """
    以 # 標籤庫（tag_database.json）詞彙與全文做子字串比對；長詞先比，减少短詞誤傷。
    命中者補上 category「hashtag」，字型以標籤庫為準；與既有 hashtag（含 #詞 已抽出者）依 fold 去重。
    """
    if not (search_text or "").strip():
        return
    try:
        from tag_store import list_hashtags
    except ImportError:
        return
    lib = list_hashtags()
    if not lib:
        return
    seen = _hashtag_fold_set(tags)
    folded_full = _fold(search_text)
    # 長詞優先：同時命中「副換菜」與「菜」時先掛長片語；短詞若已為長詞子字串仍可能重複命中，靠 fold 去重
    for phrase in sorted((p.strip() for p in lib if p.strip()), key=len, reverse=True):
        fp = _fold(phrase)
        if not fp or fp in seen:
            continue
        if fp in folded_full or phrase in search_text:
            _add_tag(tags, "hashtag", phrase)
            seen.add(fp)


def _build_tags(
    paren_tags: list[str],
    name_block: str,
    plan_block: str,
    notes_block: str,
    full_line: str,
    field_set: AnalyzeFieldSet | None = None,
) -> list[dict[str, str]]:
    fs = field_set or AnalyzeFieldSet.all_on()
    tags: list[dict[str, str]] = []
    for p in paren_tags:
        _add_tag(tags, "paren", normalize_leading_no(p))

    if fs.plan:
        plan_norm = _normalize_plan_for_promo_match(plan_block)
        _scan_snippets(plan_norm, PROMO_SNIPPETS, "promo", tags)
        _infer_promo_when_plan_is_headcount_only(plan_norm, tags)

    combined_text = (
        (name_block if fs.name else "")
        + (plan_block if fs.plan else "")
        + (notes_block if fs.notes else "")
    )
    if combined_text.strip():
        _scan_snippets(combined_text, SWAP_SNIPPETS, "swap", tags)
        _scan_snippets(combined_text, UTENSIL_SNIPPETS, "utensil", tags)
        for m in RICE_RE.finditer(combined_text):
            _add_tag(tags, "rice", m.group(0).replace(" ", ""))

    if fs.notes and (notes_block or "").strip():
        _extract_no_like_segments(notes_block, tags, verbatim_category="notes_raw")
    if fs.name and (name_block or "").strip():
        _extract_no_like_segments(name_block, tags, verbatim_category="name_column_raw")

    review_parts: list[str] = []
    if fs.name and (name_block or "").strip():
        review_parts.append(name_block)
    if fs.notes and (notes_block or "").strip():
        review_parts.append(notes_block)
    combined_review = "\n".join(review_parts)
    if (
        combined_review.strip()
        and len(combined_review.strip()) > 30
        and sum(1 for t in tags if t["category"] == "restriction") <= 1
    ):
        _add_tag(tags, "needs_review", combined_review.strip()[:200])

    ht_parts: list[str] = []
    if fs.name:
        ht_parts.append(name_block)
    if fs.plan:
        ht_parts.append(plan_block)
    if fs.notes:
        ht_parts.append(notes_block)
    if fs.full_line:
        ht_parts.append(full_line)
    for ht in merge_hashtags_from_fields(*ht_parts):
        _add_tag(tags, "hashtag", ht)

    lib_parts: list[str] = []
    if fs.name and (name_block or "").strip():
        lib_parts.append(name_block)
    if fs.plan and (plan_block or "").strip():
        lib_parts.append(plan_block)
    if fs.notes and (notes_block or "").strip():
        lib_parts.append(notes_block)
    if fs.full_line and (full_line or "").strip():
        lib_parts.append(full_line)
    lib_scan = "\n".join(lib_parts)
    if lib_scan.strip():
        _apply_tag_library(lib_scan, tags)

    return tags


def parse_line(line: str, line_no: int, field_set: AnalyzeFieldSet | None = None) -> ParsedRow:
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

    fs = field_set or AnalyzeFieldSet.all_on()
    customer_name = _extract_customer_name(name_block)
    paren_tags = [normalize_leading_no(x.strip()) for x in PAREN_RE.findall(name_block) if x.strip()]
    if not fs.name:
        paren_tags = []

    combined = f"{name_block}\t{plan_block}\t{notes_block}"
    hc_bits: list[str] = []
    if fs.name:
        hc_bits.append(name_block)
    if fs.plan:
        hc_bits.append(plan_block)
    if fs.notes:
        hc_bits.append(notes_block)
    headcount = _headcount_from_text("\t".join(hc_bits) if hc_bits else combined)

    tags = _build_tags(paren_tags, name_block, plan_block, notes_block, raw, field_set=fs)

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


def parse_bulk(text: str, field_set: AnalyzeFieldSet | None = None) -> list[ParsedRow]:
    lines = text.splitlines()
    out: list[ParsedRow] = []
    fs = field_set or AnalyzeFieldSet.all_on()
    data_row_no = 0
    for line in lines:
        if not line.strip():
            continue
        if _is_skippable_non_data_line(line):
            continue
        data_row_no += 1
        out.append(parse_line(line, data_row_no, field_set=fs))
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
