# -*- coding: utf-8 -*-
"""A4 直式 PDF 列印（主篩選名單、交叉表）。

需安裝：pip install reportlab
中文：優先註冊 Windows 內建 msjh.ttc（微軟正黑體）。
"""

from __future__ import annotations

import html
import os
from pathlib import Path
from typing import Any

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    KeepTogether,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

_FONT_NAME = "PrintCJK"
_FONT_REGISTERED = False


def _register_cjk_font() -> str:
    global _FONT_REGISTERED
    if _FONT_REGISTERED:
        return _FONT_NAME
    windir = os.environ.get("WINDIR", r"C:\Windows")
    candidates = [
        Path(windir) / "Fonts" / "msjh.ttc",
        Path(windir) / "Fonts" / "msjhbd.ttc",
        Path(windir) / "Fonts" / "msjhl.ttc",
        Path(windir) / "Fonts" / "kaiu.ttf",
        Path(windir) / "Fonts" / "mingliu.ttc",
    ]
    for p in candidates:
        if not p.is_file():
            continue
        try:
            suf = p.suffix.lower()
            if suf == ".ttc":
                pdfmetrics.registerFont(TTFont(_FONT_NAME, str(p), subfontIndex=0))
            else:
                pdfmetrics.registerFont(TTFont(_FONT_NAME, str(p)))
            _FONT_REGISTERED = True
            return _FONT_NAME
        except Exception:
            continue
    raise RuntimeError(
        "找不到可用的中文字型（已嘗試 Windows Fonts 內 msjh.ttc 等）。"
        "請確認系統已安裝微軟正黑體，或手動指定字型路徑。"
    )


def _para(text: str, style: ParagraphStyle) -> Paragraph:
    return Paragraph(html.escape(text).replace("\n", "<br/>"), style)


def save_primary_filter_pdf(
    app: Any,
    dest: str | Path,
    tags: list[str],
    *,
    name_cols: int = 7,
    name_font_size: float = 7.8,
) -> None:
    """主標籤篩選：依 tags 順序輸出各區塊（標題、筆數、分量統計、名單表格）。"""
    from filter_prefs import normalize_display_rule

    _register_cjk_font()
    dest = Path(dest)
    tags = [str(t).strip() for t in tags if str(t).strip()]
    if not tags:
        raise ValueError("沒有可列印的標籤。")

    margin = 9 * mm
    doc = SimpleDocTemplate(
        str(dest),
        pagesize=A4,
        leftMargin=margin,
        rightMargin=margin,
        topMargin=margin,
        bottomMargin=margin,
        title="主標籤篩選名單",
    )
    styles = getSampleStyleSheet()
    title_st = ParagraphStyle(
        "t",
        parent=styles["Heading1"],
        fontName=_FONT_NAME,
        fontSize=12,
        leading=15,
        textColor=colors.HexColor("#0D47A1"),
    )
    h2_st = ParagraphStyle(
        "h2",
        parent=styles["Heading2"],
        fontName=_FONT_NAME,
        fontSize=10,
        leading=12,
        textColor=colors.HexColor("#1565C0"),
    )
    body_st = ParagraphStyle(
        "b",
        parent=styles["Normal"],
        fontName=_FONT_NAME,
        fontSize=7.5,
        leading=9.5,
    )
    small_st = ParagraphStyle(
        "s",
        parent=styles["Normal"],
        fontName=_FONT_NAME,
        fontSize=6.8,
        leading=8.4,
        textColor=colors.HexColor("#424242"),
    )
    fenji_st = ParagraphStyle(
        "fenji",
        parent=styles["Normal"],
        fontName=_FONT_NAME,
        fontSize=8.4,
        leading=10.6,
        textColor=colors.HexColor("#0D47A1"),
    )

    story: list = []
    story.append(_para("主標籤篩選名單（A4 直式）", title_st))
    story.append(Spacer(1, 4 * mm))

    for tag in tags:
        matches = app._sort_primary_filter_matches(app._rows_matching_tag_value(tag))
        if not matches:
            continue
        rule = normalize_display_rule(app._get_display_rule(tag))
        n = len(matches)
        block: list = []
        block.append(_para(f"【{tag}】　共 {n} 筆", h2_st))
        block.append(Spacer(1, 1 * mm))
        for line in app._primary_filter_block_stats_lines(matches, rule):
            if str(line).strip().startswith("分量分計"):
                block.append(_para(line, fenji_st))
                block.append(Spacer(1, 0.8 * mm))
            else:
                block.append(_para(line, small_st))
        block.append(Spacer(1, 2 * mm))

        # 名單改為多欄併排（不顯示序號），減少列印頁數。
        # 內容以「姓名(小/大)」為主，保留區塊清楚分隔即可。
        # 依資料頁分段：每個資料頁在同一標籤內獨立排列，第一列顯示資料頁名。
        ncols = min(10, max(3, int(name_cols)))  # 人名欄數（不含左側資料頁標示欄）
        by_page: dict[str, list[tuple[str, str | None]]] = {}
        for r in matches:
            pg = (getattr(r, "source_page", None) or "").strip() or "（無頁名）"
            nm = (getattr(r, "customer_name", "") or "").strip() or "（無姓名）"
            sz = app._row_headcount_str(r)
            sz_lab = app._crosstab_row_category(r) if sz else ""
            frame_kind: str | None = None
            fk = app._fenji_stat_bucket(r, rule)
            if fk == "disp":
                frame_kind = "disp"
            elif fk == "ut":
                frame_kind = "ut"
            mark = "(拋)" if frame_kind == "disp" else ("(自)" if frame_kind == "ut" else "")
            name_text = f"{nm}{mark}({sz_lab})" if sz_lab in ("小", "大") else f"{nm}{mark}"
            by_page.setdefault(pg, []).append((name_text, frame_kind))

        rows: list[list[str]] = []
        frame_cells: dict[tuple[int, int], str] = {}
        page_rows: set[int] = set()
        rr = 0
        for pg in sorted(by_page.keys()):
            cur_page = by_page.get(pg, [])
            if not cur_page:
                continue
            ci = 0
            row_cells = [pg] + [""] * ncols
            for name_text, fk in cur_page:
                row_cells[1 + ci] = name_text
                if fk:
                    frame_cells[(rr, 1 + ci)] = fk
                ci += 1
                if ci >= ncols:
                    rows.append(row_cells)
                    page_rows.add(rr)
                    rr += 1
                    ci = 0
                    row_cells = [""] + [""] * ncols
            if any(x for x in row_cells[1:]) or row_cells[0]:
                rows.append(row_cells)
                page_rows.add(rr)  # 第一行有頁名；後續會是空字串
                rr += 1
            # 各資料頁段落間插入一個小空行，便於辨識
            rows.append([""] * (ncols + 1))
            rr += 1

        while rows and not any(rows[-1]):
            rows.pop()
            rr -= 1
        if not rows:
            rows = [[""] * (ncols + 1)]

        tw = doc.width
        page_col_w = tw * 0.14
        name_col_w = (tw - page_col_w) / max(ncols, 1)
        col_w = [page_col_w] + [name_col_w] * ncols
        t = Table(rows, colWidths=col_w)
        t.setStyle(
            TableStyle(
                [
                    ("FONT", (0, 0), (-1, -1), _FONT_NAME, min(12.0, max(6.0, float(name_font_size)))),
                    ("GRID", (0, 0), (-1, -1), 0.2, colors.HexColor("#CFD8DC")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 1.2),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 1.2),
                    ("TOPPADDING", (0, 0), (-1, -1), 1.0),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 1.0),
                ]
            )
        )
        for r0 in sorted(page_rows):
            # 資料頁標示列：左欄淡底與加粗，方便快速分段。
            t.setStyle(
                TableStyle(
                    [
                        ("FONT", (0, r0), (0, r0), _FONT_NAME, min(12.0, max(6.0, float(name_font_size))) + 0.2),
                        ("BACKGROUND", (0, r0), (0, r0), colors.HexColor("#E8EAF6")),
                    ]
                )
            )
        # 拋棄式（藍框）／自備餐具（綠框）姓名外框
        for (r0, c0), fk in frame_cells.items():
            if fk == "disp":
                col = colors.HexColor("#0D47A1")
            elif fk == "ut":
                col = colors.HexColor("#2E7D32")
            else:
                continue
            t.setStyle(TableStyle([("BOX", (c0, r0), (c0, r0), 0.8, col)]))
        block.append(t)
        block.append(Spacer(1, 1.6 * mm))
        story.append(KeepTogether(block))

    if len(story) <= 2:
        raise ValueError("所選標籤皆無資料可列印。")

    doc.build(story)


def save_crosstab_pdf(app: Any, dest: str | Path) -> None:
    """交叉表：與畫面相同結構（分資料頁區塊＋全資料頁合計列）。"""
    _register_cjk_font()
    dest = Path(dest)

    data, err = app._compute_crosstab_matrix()
    if err or data is None:
        raise ValueError(
            "無法產生交叉表 PDF（請確認已分析且已選欄標籤）。"
            if err == "no_cols"
            else "無法產生交叉表 PDF。"
        )

    row_kinds, col_tags, page_blocks, grand_col_totals, grand = data
    if not col_tags:
        raise ValueError("交叉表沒有可見欄位。")

    margin = 8 * mm
    doc = SimpleDocTemplate(
        str(dest),
        pagesize=landscape(A4),
        leftMargin=margin,
        rightMargin=margin,
        topMargin=margin,
        bottomMargin=margin,
        title="交叉表（A4橫印）",
    )

    styles = getSampleStyleSheet()
    title_st = ParagraphStyle(
        "ct",
        parent=styles["Heading1"],
        fontName=_FONT_NAME,
        fontSize=11.5,
        leading=14,
        textColor=colors.HexColor("#0D47A1"),
    )
    sect_st = ParagraphStyle(
        "cs",
        parent=styles["Normal"],
        fontName=_FONT_NAME,
        fontSize=8,
        leading=10,
        backColor=colors.HexColor("#E8EAF6"),
        borderPadding=4,
        textColor=colors.HexColor("#1A237E"),
    )

    story: list = []
    story.append(_para("交叉表（A4 直式）", title_st))
    story.append(Spacer(1, 3 * mm))

    nloc = len(col_tags)
    label_col_w = doc.width * 0.11
    data_col_total = doc.width - label_col_w
    per_col = data_col_total / max(nloc + 1, 1)

    def _page_table(
        page_title: str,
        mat: list[list[int]],
        row_sums: list[int],
        col_sums: list[int],
        corner: int,
    ) -> Table:
        hdr = ["分量", *col_tags, "列合計"]
        rows_pdf: list[list[str]] = [hdr]
        for i, rk in enumerate(row_kinds):
            rows_pdf.append(
                [rk, *[str(mat[i][j]) for j in range(nloc)], str(row_sums[i])]
            )
        rows_pdf.append(
            ["欄合計", *[str(col_sums[j]) for j in range(nloc)], str(corner)]
        )
        cw = [label_col_w] + [per_col] * (nloc + 1)
        tbl = Table(rows_pdf, colWidths=cw, repeatRows=1)
        tbl.setStyle(
            TableStyle(
                [
                    ("FONT", (0, 0), (-1, -1), _FONT_NAME, 6.0),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E3F2FD")),
                    ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#FFF8E1")),
                    ("GRID", (0, 0), (-1, -1), 0.2, colors.HexColor("#CFD8DC")),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("ALIGN", (1, 0), (-1, -1), "CENTER"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 1.2),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 1.2),
                    ("TOPPADDING", (0, 0), (-1, -1), 1.0),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 1.0),
                ]
            )
        )
        return tbl

    for _bi, (page, mat_f, row_sums_f, col_sums_f, blk_corner) in enumerate(page_blocks):
        block: list = []
        block.append(_para(f"【{page}】", sect_st))
        block.append(Spacer(1, 1.2 * mm))
        block.append(_page_table(page, mat_f, row_sums_f, col_sums_f, blk_corner))
        block.append(Spacer(1, 2 * mm))
        story.append(KeepTogether(block))

    grand_hdr = ["標籤合計", *[str(x) for x in grand_col_totals], str(grand)]
    g_tbl = Table(
        [grand_hdr],
        colWidths=[label_col_w] + [per_col] * (nloc + 1),
    )
    g_tbl.setStyle(
        TableStyle(
            [
                ("FONT", (0, 0), (-1, -1), _FONT_NAME, 7),
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#E1F5FE")),
                ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#81D4FA")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (1, 0), (-1, -1), "CENTER"),
            ]
        )
    )
    story.append(g_tbl)
    story.append(Spacer(1, 3 * mm))
    story.append(
        _para(
            f"全表儲存格加總：{grand}（同人多標籤會重複計入欄位，非去重人數）",
            ParagraphStyle(
                "fn",
                parent=styles["Normal"],
                fontName=_FONT_NAME,
                fontSize=7,
                leading=10,
                textColor=colors.HexColor("#616161"),
            ),
        )
    )

    doc.build(story)
