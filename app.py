# -*- coding: utf-8 -*-
"""
訂餐備註分析小工具：分頁式桌面介面。
依賴：標準庫（tkinter）；匯出 Word 時需安裝 python-docx（pip install python-docx）。
執行：python app.py
"""

from __future__ import annotations

from collections import Counter
import csv
import json
import os
import shutil
import subprocess
import sys
import tempfile
import tkinter as tk
from pathlib import Path
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk

from order_note_parser import (
    format_order_serial,
    headcount_size_label,
    normalize_leading_no,
    parse_bulk,
    rows_to_csv_text,
    rows_to_json,
)
from filter_prefs import (
    DEFAULT_DISPLAY_RULE,
    DEFAULT_EXPORT_CUSTOM_TEMPLATES,
    load_filter_prefs,
    normalize_display_rule,
    normalize_export_templates,
    save_filter_prefs,
)
from input_pages_store import (
    ROSTER_VIEW_ALL,
    allocate_page_id,
    load_input_pages_state,
    save_input_pages_state,
)
from tag_store import (
    database_path,
    list_hashtags,
    register_hashtags,
    replace_hashtags_from_text,
    save_hashtag_list,
)

# 主標籤篩選：捲動區底色、各區塊輪替底色
FILTER_INNER_BG = "#ECECEC"
# 主標籤篩選：名單／統計字級
FILTER_ROSTER_FONT = ("Microsoft JhengHei UI", 12)
FILTER_STAT_FONT = ("Microsoft JhengHei UI", 14, "bold")
# 主標籤篩選：區塊「輸出」勾選（加大加粗以利辨識）
FILTER_EXPORT_CHECK_FONT = ("Microsoft JhengHei UI", 15, "bold")
# 流式名單：邊距、欄距（左緣到左緣固定步進 = 該區最寬字寬 + 此值）、列距
FILTER_ROSTER_PAD = 10
FILTER_ROSTER_ITEM_GAP_H = 16
FILTER_ROSTER_ROW_GAP_V = 8
# 姓名加「拋棄式」外框時，欄寬預留的額外像素（左右各半）
FILTER_DISPOSABLE_WIDTH_PAD = 10
# 資料判定須為完整詞「拋棄式」（僅「拋」不算）
_DISPOSABLE_MARKER = "拋棄式"
FILTER_BLOCK_BGS = (
    "#E3F2FD",
    "#E8F5E9",
    "#FFF8E1",
    "#F3E5F5",
    "#E0F7FA",
    "#FFEBEE",
    "#F1F8E9",
    "#FFF3E0",
)

# 匯出預覽排列：(內部 key, 下拉顯示文字)
_EXPORT_LAYOUT_OPTIONS: tuple[tuple[str, str], ...] = (
    ("screen", "與篩選區塊相同（標題・流式欄距・統計）"),
    ("tsv", "Tab 分欄（表頭，試算表／直印）"),
    ("print_cols", "直印對齊（等寬欄・空格）"),
    ("flow3", "橫向三欄（省紙，│ 分隔）"),
    ("flow4", "橫向四欄（省紙，│ 分隔）"),
    ("names", "僅姓名（每筆一行）"),
    ("custom", "自訂（下方格式字串）"),
)

# 分頁列：選中與未選（vista 主題無法改分頁底色，故改用可自訂的 clam）
_NOTEBOOK_TAB_BG = "#D8D8D8"
_NOTEBOOK_TAB_BG_SELECTED = "#1565C0"
_NOTEBOOK_TAB_FG_SELECTED = "#FFFFFF"
_NOTEBOOK_TAB_BG_ACTIVE = "#42A5F5"
_NOTEBOOK_TAB_FG = "#1A1A1A"


class OrderNoteApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("訂餐備註分析／標籤")
        self.root.minsize(900, 640)
        self._style_notebook_tabs()
        self._rows: list = []
        self._row_by_iid: dict[str, int] = {}
        (
            self._filter_selected_tags,
            self._filter_display_rules,
            self._filter_export_blocks,
            self._export_custom_templates,
            self._crosstab_col_tags,
        ) = load_filter_prefs()
        self._filter_export_vars: dict[str, tk.IntVar] = {}
        # 交叉表欄標籤由 filter_prefs（primary_filter_selection.json）載入／儲存
        self._pages_state: dict = load_input_pages_state()
        self._page_list_updating = False
        self._roster_view_var = tk.StringVar(
            value=str(self._pages_state.get("roster_view") or ROSTER_VIEW_ALL)
        )

        outer = ttk.Frame(self.root, padding=8)
        outer.pack(fill=tk.BOTH, expand=True)

        self._nb = ttk.Notebook(outer)
        self._nb.pack(fill=tk.BOTH, expand=True)

        self._build_tab_input()
        self._build_tab_roster()
        self._build_tab_primary_filter()
        self._build_tab_crosstab()
        self._build_tab_export()
        self._build_tab_hashtag_db()

        self._status = tk.StringVar(value="就緒")
        ttk.Label(outer, textvariable=self._status).pack(anchor=tk.W, pady=(8, 0))

        self._refresh_page_listbox()
        self._apply_current_page_to_txt_in()
        self._sync_roster_page_choices(set_combo=True)
        if self._rows and self._crosstab_col_tags:
            self._refresh_crosstab_table()
        self.root.protocol("WM_DELETE_WINDOW", self._on_app_close)

    def _style_notebook_tabs(self) -> None:
        st = ttk.Style(self.root)
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass
        st.configure("TNotebook", background="#ECECEC", borderwidth=0)
        st.configure(
            "TNotebook.Tab",
            padding=(11, 5),
            background=_NOTEBOOK_TAB_BG,
            foreground=_NOTEBOOK_TAB_FG,
        )
        st.map(
            "TNotebook.Tab",
            background=[
                ("selected", _NOTEBOOK_TAB_BG_SELECTED),
                ("active", _NOTEBOOK_TAB_BG_ACTIVE),
                ("!selected", _NOTEBOOK_TAB_BG),
            ],
            foreground=[
                ("selected", _NOTEBOOK_TAB_FG_SELECTED),
                ("active", _NOTEBOOK_TAB_FG_SELECTED),
                ("!selected", _NOTEBOOK_TAB_FG),
            ],
            expand=[("selected", [1, 1, 1, 0])],
        )

    def _on_app_close(self) -> None:
        self._sync_txt_in_to_current_page()
        self._pages_state["roster_view"] = self._roster_view_var.get()
        try:
            save_input_pages_state(self._pages_state)
        except OSError:
            pass
        try:
            self._export_custom_templates = self._get_export_templates_live()
            save_filter_prefs(
                self._filter_selected_tags,
                self._filter_display_rules,
                self._filter_export_blocks,
                self._export_custom_templates,
                crosstab_col_tags=self._crosstab_col_tags,
            )
        except OSError:
            pass
        self.root.destroy()

    def _page_by_id(self, page_id: str) -> dict | None:
        for p in self._pages_state.get("pages") or []:
            if p.get("id") == page_id:
                return p
        return None

    def _sync_txt_in_to_current_page(self) -> None:
        p = self._page_by_id(self._pages_state.get("current_page_id") or "")
        if p is not None:
            p["text"] = self.txt_in.get("1.0", tk.END)

    def _apply_current_page_to_txt_in(self) -> None:
        p = self._page_by_id(self._pages_state.get("current_page_id") or "")
        self.txt_in.delete("1.0", tk.END)
        if p is not None:
            self.txt_in.insert("1.0", p.get("text") or "")

    def _refresh_page_listbox(self) -> None:
        self._page_list_updating = True
        try:
            self._page_list.delete(0, tk.END)
            for p in self._pages_state.get("pages") or []:
                self._page_list.insert(tk.END, p.get("name") or "未命名")
            cur = self._pages_state.get("current_page_id")
            pages = self._pages_state.get("pages") or []
            for i, p in enumerate(pages):
                if p.get("id") == cur:
                    self._page_list.selection_clear(0, tk.END)
                    self._page_list.selection_set(i)
                    self._page_list.activate(i)
                    self._page_list.see(i)
                    break
        finally:
            self._page_list_updating = False

    def _on_page_list_select(self, _evt=None) -> None:
        if self._page_list_updating:
            return
        sel = self._page_list.curselection()
        if not sel:
            return
        pages = self._pages_state.get("pages") or []
        i = int(sel[0])
        if i < 0 or i >= len(pages):
            return
        self._sync_txt_in_to_current_page()
        self._pages_state["current_page_id"] = pages[i]["id"]
        self._apply_current_page_to_txt_in()
        try:
            save_input_pages_state(self._pages_state)
        except OSError:
            pass

    def _new_input_page(self) -> None:
        from tkinter.simpledialog import askstring

        name = askstring(
            "新增頁",
            "頁名（將視為主標籤並寫入標籤庫）：",
            parent=self.root,
        )
        name = (name or "").strip()
        if not name:
            return
        for p in self._pages_state.get("pages") or []:
            if (p.get("name") or "").strip() == name:
                messagebox.showwarning("新增頁", "已存在相同頁名，請使用不同名稱。", parent=self.root)
                return
        try:
            register_hashtags([name])
        except OSError as e:
            messagebox.showerror("標籤庫", f"無法寫入標籤庫：{e}", parent=self.root)
            return
        self._sync_txt_in_to_current_page()
        new_id = allocate_page_id()
        self._pages_state.setdefault("pages", []).append({"id": new_id, "name": name, "text": ""})
        self._pages_state["current_page_id"] = new_id
        self._refresh_page_listbox()
        self._apply_current_page_to_txt_in()
        self._sync_roster_page_choices(set_combo=True)
        try:
            save_input_pages_state(self._pages_state)
        except OSError as e:
            messagebox.showwarning("儲存", f"無法寫入輸入頁設定：{e}", parent=self.root)
        self._refresh_hashtag_db_view()
        self._status.set(f"已新增頁「{name}」並寫入標籤庫。")

    def _delete_input_page(self) -> None:
        pages = self._pages_state.get("pages") or []
        if len(pages) <= 1:
            messagebox.showwarning("刪除頁", "至少需要保留一頁資料。", parent=self.root)
            return

        sel = self._page_list.curselection()
        if sel:
            idx = int(sel[0])
        else:
            cur_id = self._pages_state.get("current_page_id")
            idx = next((i for i, p in enumerate(pages) if p.get("id") == cur_id), 0)
        if idx < 0 or idx >= len(pages):
            return

        victim = pages[idx]
        vname = str(victim.get("name") or "未命名").strip() or "未命名"
        if not messagebox.askyesno(
            "刪除頁",
            f"確定刪除頁「{vname}」？\n此頁內的文字將一併移除；名單若曾分析過，請重新分析以更新合併結果。",
            parent=self.root,
        ):
            return

        self._sync_txt_in_to_current_page()
        cur_id_before = self._pages_state.get("current_page_id")
        del pages[idx]
        self._pages_state["pages"] = pages

        if cur_id_before == victim.get("id"):
            new_idx = min(idx, len(pages) - 1)
            self._pages_state["current_page_id"] = pages[new_idx]["id"]
        # 若名單檢視正選被刪的頁，改回「全部頁」
        if self._roster_view_var.get() == vname:
            self._roster_view_var.set(ROSTER_VIEW_ALL)
            self._pages_state["roster_view"] = ROSTER_VIEW_ALL

        self._refresh_page_listbox()
        self._apply_current_page_to_txt_in()
        self._sync_roster_page_choices(set_combo=True)
        self._refresh_roster_tree()
        try:
            save_input_pages_state(self._pages_state)
        except OSError as e:
            messagebox.showwarning("儲存", f"無法寫入輸入頁設定：{e}", parent=self.root)
        self._status.set(f"已刪除頁「{vname}」。")

    def _sync_roster_page_choices(self, set_combo: bool = True) -> None:
        pages = self._pages_state.get("pages") or []
        names = [str(p.get("name") or "未命名") for p in pages]
        vals = [ROSTER_VIEW_ALL] + names
        if hasattr(self, "_roster_view_cb"):
            self._roster_view_cb["values"] = vals
        cur = self._roster_view_var.get()
        if cur not in vals:
            cur = ROSTER_VIEW_ALL
            self._roster_view_var.set(cur)
            self._pages_state["roster_view"] = cur
        if set_combo and hasattr(self, "_roster_view_cb"):
            self._roster_view_cb.set(cur)

    def _on_roster_view_change(self, _evt=None) -> None:
        v = self._roster_view_var.get()
        self._pages_state["roster_view"] = v
        self._sync_txt_in_to_current_page()
        try:
            save_input_pages_state(self._pages_state)
        except OSError:
            pass
        self._refresh_roster_tree()

    def _refresh_roster_tree(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._row_by_iid.clear()
        self.txt_detail.delete("1.0", tk.END)
        if not self._rows:
            return
        mode = self._roster_view_var.get()
        if mode == ROSTER_VIEW_ALL or not mode:
            visible = list(range(len(self._rows)))
        else:
            visible = [i for i, r in enumerate(self._rows) if (getattr(r, "source_page", None) or "") == mode]
        for display_pos, gi in enumerate(visible):
            r = self._rows[gi]
            iid = str(display_pos)
            self._row_by_iid[iid] = gi
            prev = (r.notes_block or r.plan_block)[:60].replace("\n", " ")
            if len((r.notes_block or r.plan_block)) > 60:
                prev += "…"
            self.tree.insert(
                "",
                tk.END,
                iid=iid,
                values=(
                    r.line_no,
                    r.serial,
                    r.customer_name,
                    r.headcount or "",
                    len(r.tags),
                    r.source_page or "",
                    prev,
                ),
            )

    # --- 分頁：輸入與分析 ---
    def _build_tab_input(self) -> None:
        tab = ttk.Frame(self._nb, padding=8)
        self._nb.add(tab, text="輸入與分析")

        ttk.Label(
            tab,
            text="左側可新增多頁；每頁可貼上原始資料（Tab 分欄）。「分析」會合併所有頁一併解析；頁名視為主標籤並寫入標籤庫。",
            wraplength=820,
        ).pack(anchor=tk.W)

        pan = ttk.PanedWindow(tab, orient=tk.HORIZONTAL)
        pan.pack(fill=tk.BOTH, expand=True, pady=(8, 8))

        left = ttk.Frame(pan, width=200)
        pan.add(left, weight=0)
        ttk.Label(left, text="資料頁").pack(anchor=tk.W)
        self._page_list = tk.Listbox(left, height=14, exportselection=False)
        self._page_list.pack(fill=tk.BOTH, expand=True, pady=(4, 6))
        self._page_list.bind("<<ListboxSelect>>", self._on_page_list_select)
        btn_pages = ttk.Frame(left)
        btn_pages.pack(fill=tk.X)
        ttk.Button(btn_pages, text="新增頁", command=self._new_input_page).pack(fill=tk.X, pady=(0, 4))
        ttk.Button(btn_pages, text="刪除頁", command=self._delete_input_page).pack(fill=tk.X)

        right = ttk.Frame(pan)
        pan.add(right, weight=1)

        self.txt_in = tk.Text(
            right,
            height=18,
            wrap=tk.NONE,
            font=("Consolas", 10),
            undo=True,
        )
        sy = ttk.Scrollbar(right, orient=tk.VERTICAL, command=self.txt_in.yview)
        sx = ttk.Scrollbar(right, orient=tk.HORIZONTAL, command=self.txt_in.xview)
        self.txt_in.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.txt_in.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        right.rowconfigure(0, weight=1)
        right.columnconfigure(0, weight=1)

        btn_row = ttk.Frame(tab)
        btn_row.pack(fill=tk.X)
        ttk.Button(btn_row, text="分析並產生標籤", command=self._analyze).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_row, text="清空本頁輸入", command=self._clear_input).pack(side=tk.LEFT)

    # --- 分頁：名單與標籤 ---
    def _build_tab_roster(self) -> None:
        tab = ttk.Frame(self._nb, padding=8)
        self._tab_roster = tab
        self._nb.add(tab, text="名單與標籤")

        view_row = ttk.Frame(tab)
        view_row.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(view_row, text="檢視資料頁：").pack(side=tk.LEFT, padx=(0, 6))
        self._roster_view_cb = ttk.Combobox(
            view_row,
            textvariable=self._roster_view_var,
            state="readonly",
            width=28,
        )
        self._roster_view_cb.pack(side=tk.LEFT)
        self._roster_view_cb.bind("<<ComboboxSelected>>", self._on_roster_view_change)

        top = ttk.Frame(tab)
        top.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(top, text="解析結果（點列可檢視 JSON；可手動新增標籤）。").pack(side=tk.LEFT, anchor=tk.W)
        ttk.Button(top, text="新增標籤", command=self._add_manual_tag).pack(side=tk.RIGHT)

        paned = ttk.PanedWindow(tab, orient=tk.HORIZONTAL)
        paned.pack(fill=tk.BOTH, expand=True)

        tree_frame = ttk.Frame(paned)
        paned.add(tree_frame, weight=3)

        cols = ("line_no", "serial", "name", "headcount", "tag_count", "page", "preview")
        self.tree = ttk.Treeview(tree_frame, columns=cols, show="headings", height=16)
        headings = [
            ("line_no", "行"),
            ("serial", "序號"),
            ("name", "姓名"),
            ("headcount", "人數"),
            ("tag_count", "標籤數"),
            ("page", "資料頁"),
            ("preview", "備註摘要"),
        ]
        for cid, text in headings:
            self.tree.heading(cid, text=text)
        self.tree.column("line_no", width=40, stretch=False)
        self.tree.column("serial", width=50, stretch=False)
        self.tree.column("name", width=100)
        self.tree.column("headcount", width=80, stretch=False)
        self.tree.column("tag_count", width=60, stretch=False)
        self.tree.column("page", width=90, stretch=False)
        self.tree.column("preview", width=260)

        sy_tree = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=sy_tree.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sy_tree.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.bind("<<TreeviewSelect>>", self._on_select_row)

        detail_frame = ttk.LabelFrame(paned, text="本列標籤 JSON", padding=4)
        paned.add(detail_frame, weight=2)

        self.txt_detail = tk.Text(detail_frame, height=14, wrap=tk.WORD, font=("Consolas", 9))
        sy_d = ttk.Scrollbar(detail_frame, orient=tk.VERTICAL, command=self.txt_detail.yview)
        self.txt_detail.configure(yscrollcommand=sy_d.set)
        self.txt_detail.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sy_d.pack(side=tk.RIGHT, fill=tk.Y)

    # --- 分頁：主標籤篩選 ---
    def _build_tab_primary_filter(self) -> None:
        tab = ttk.Frame(self._nb, padding=8)
        self._tab_primary_filter = tab
        self._nb.add(tab, text="主標籤篩選")

        top_bar = ttk.Frame(tab)
        top_bar.pack(fill=tk.X, pady=(0, 8))

        bar = ttk.Frame(top_bar)
        bar.pack(side=tk.LEFT, anchor=tk.NW)
        ttk.Button(bar, text="選擇主標籤…", command=self._open_primary_tag_picker).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(bar, text="依目前選取更新名單", command=self._refresh_primary_filter_results).pack(side=tk.LEFT)
        ttk.Button(bar, text="匯出…", command=self._open_export_from_primary_filter).pack(side=tk.LEFT, padx=(12, 0))

        # 合計表置於右上（與按鈕同一列）
        self._filter_summary_host = tk.Frame(
            top_bar, bg="#E8EAF6", highlightbackground="#9FA8DA", highlightthickness=1
        )
        self._filter_summary_host.pack(side=tk.RIGHT, anchor=tk.NE, padx=(12, 0))

        outer_scroll = ttk.Frame(tab)
        outer_scroll.pack(fill=tk.BOTH, expand=True)

        self._filter_canvas = tk.Canvas(outer_scroll, highlightthickness=0, bg=FILTER_INNER_BG)
        self._filter_vsb = ttk.Scrollbar(outer_scroll, orient=tk.VERTICAL, command=self._filter_canvas.yview)
        self._filter_canvas.configure(yscrollcommand=self._filter_vsb.set)
        self._filter_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._filter_vsb.pack(side=tk.RIGHT, fill=tk.Y)

        self._filter_inner = tk.Frame(self._filter_canvas, bg=FILTER_INNER_BG, padx=4, pady=4)
        self._filter_win_id = self._filter_canvas.create_window((0, 0), window=self._filter_inner, anchor="nw")

        def _inner_cfg(_e: tk.Event | None = None) -> None:
            bb = self._filter_canvas.bbox("all")
            if bb:
                self._filter_canvas.configure(scrollregion=bb)

        def _canvas_cfg(e: tk.Event) -> None:
            self._filter_canvas.itemconfigure(self._filter_win_id, width=e.width)

        self._filter_inner.bind("<Configure>", _inner_cfg)
        self._filter_canvas.bind("<Configure>", _canvas_cfg)

        tk.Label(
            self._filter_inner,
            text="請至「# 標籤庫」維護主標籤，再按「選擇主標籤…」從庫中勾選；完成「輸入與分析」後可檢視各標籤名單。",
            bg=FILTER_INNER_BG,
            fg="#333333",
            justify=tk.LEFT,
            wraplength=780,
        ).pack(anchor=tk.W, pady=8)

        self._filter_canvas.bind("<MouseWheel>", self._on_filter_mousewheel)
        self._filter_vsb.bind("<MouseWheel>", self._on_filter_mousewheel)

    def _rows_matching_tag_value(self, tag_value: str) -> list:
        """該筆任一標籤的 value 與 tag_value 完全相同即命中（不限 hashtag／manual）。"""
        out = []
        for r in self._rows:
            if any(t.get("value") == tag_value for t in r.tags):
                out.append(r)
        return out

    @staticmethod
    def _row_headcount_str(r) -> str | None:
        if getattr(r, "headcount", None):
            return str(r.headcount).strip() or None
        for t in r.tags:
            if t.get("category") == "headcount":
                v = t.get("value")
                if isinstance(v, str) and v.strip():
                    return v.strip()
        return None

    def _roster_cell_and_size(self, r) -> tuple[str, str]:
        """名單顯示字串與 小/大/空（供統計）。"""
        sz = headcount_size_label(self._row_headcount_str(r))
        ser = format_order_serial(r.serial)
        nm = (r.customer_name or "").strip() or "（無姓名）"
        cell = f"{ser} {nm}({sz})" if sz else f"{ser} {nm}"
        return cell, sz

    def _get_display_rule(self, tag: str) -> dict[str, bool]:
        return normalize_display_rule(self._filter_display_rules.get(tag, DEFAULT_DISPLAY_RULE))

    @staticmethod
    def _row_has_disposable_in_data(r) -> bool:
        """原始資料或姓名／備註／標籤等是否含「拋棄式」（須完整詞，不可只比對「拋」）。"""
        chunks: list[str] = [
            getattr(r, "raw_line", "") or "",
            getattr(r, "customer_name", "") or "",
            getattr(r, "name_block", "") or "",
            getattr(r, "notes_block", "") or "",
            getattr(r, "plan_block", "") or "",
        ]
        for t in getattr(r, "tags", None) or []:
            if isinstance(t, dict):
                v = t.get("value")
                if isinstance(v, str):
                    chunks.append(v)
        hay = "\n".join(chunks)
        return _DISPOSABLE_MARKER in hay

    def _roster_segments(self, r, rule: dict[str, bool]) -> list[tuple[str, bool]]:
        """(片段文字, 是否為姓名且因「勾選拋棄式＋資料含拋棄式」而套外框)。"""
        rule = normalize_display_rule(rule)
        sz = headcount_size_label(self._row_headcount_str(r))
        frame_name = (
            bool(rule["name"])
            and bool(rule["disposable"])
            and self._row_has_disposable_in_data(r)
        )
        segs: list[tuple[str, bool]] = []
        if rule["serial"]:
            segs.append((format_order_serial(r.serial), False))
        if rule.get("page_tag"):
            pg = (getattr(r, "source_page", None) or "").strip()
            if pg:
                segs.append((pg, False))
        if rule["name"]:
            nm = (r.customer_name or "").strip() or "（無姓名）"
            segs.append((nm, frame_name))
        if rule["size_label"] and sz:
            segs.append((f"({sz})", False))
        return segs

    def _roster_plain_width(self, r, rule: dict[str, bool]) -> float:
        fnt = tkfont.Font(font=FILTER_ROSTER_FONT)
        segs = self._roster_segments(r, rule)
        if not segs:
            return float(fnt.measure(" "))
        sp = float(fnt.measure(" "))
        w = 0.0
        for i, (text, use_disposable_frame) in enumerate(segs):
            if i > 0:
                w += sp
            w += float(fnt.measure(text))
            if use_disposable_frame:
                w += float(FILTER_DISPOSABLE_WIDTH_PAD)
        return w

    def _filter_unified_roster_pitch(self, tags: list[str]) -> float:
        """本次篩選內所有會顯示的名單，共用同一欄距（與最寬一筆對齊）。"""
        fnt = tkfont.Font(font=FILTER_ROSTER_FONT)
        widths: list[float] = []
        for tag in tags:
            rule = self._get_display_rule(tag)
            for r in self._rows_matching_tag_value(tag):
                widths.append(self._roster_plain_width(r, rule))
        if not widths:
            return max(48 + FILTER_ROSTER_ITEM_GAP_H, 80)
        max_tw = max(widths)
        return float(max(max_tw + FILTER_ROSTER_ITEM_GAP_H, 80))

    def _on_filter_mousewheel(self, event: tk.Event) -> None:
        if getattr(event, "delta", 0):
            self._filter_canvas.yview_scroll(int(-1 * event.delta / 120), "units")

    def _filter_bind_mousewheel_recursive(self, widget: tk.Widget) -> None:
        widget.bind("<MouseWheel>", self._on_filter_mousewheel)
        for c in widget.winfo_children():
            self._filter_bind_mousewheel_recursive(c)

    def _render_compact_roster(
        self, parent: tk.Frame, matches: list, bg: str, roster_pitch: float, rule: dict[str, bool]
    ) -> tuple[int, int, int, int]:
        """依寬度自動換行；欄距 roster_pitch 與其他區塊相同。回傳 (小計, 大計, 小含拋棄式, 大含拋棄式)。"""
        rule = normalize_display_rule(rule)
        small_n = large_n = small_disp = large_disp = other_n = 0
        rows_segs: list[list[tuple[str, bool]]] = []
        use_disp = bool(rule["disposable"])
        for r in matches:
            sz = headcount_size_label(self._row_headcount_str(r))
            hit = use_disp and self._row_has_disposable_in_data(r)
            if sz == "小":
                small_n += 1
                if hit:
                    small_disp += 1
            elif sz == "大":
                large_n += 1
                if hit:
                    large_disp += 1
            else:
                other_n += 1
            rows_segs.append(self._roster_segments(r, rule))

        pad = FILTER_ROSTER_PAD
        gap_v = FILTER_ROSTER_ROW_GAP_V
        pitch = roster_pitch
        fnt = tkfont.Font(font=FILTER_ROSTER_FONT)
        sp_w = float(fnt.measure(" "))

        cv = tk.Canvas(parent, bg=bg, highlightthickness=0, height=24)
        cv.pack(fill=tk.X, expand=False)

        def layout(_event: tk.Event | None = None) -> None:
            cv.delete("all")
            cv.update_idletasks()
            ww = cv.winfo_width()
            if ww < 30:
                parent.update_idletasks()
                ww = max(cv.winfo_width(), parent.winfo_width() - 4, 280)

            if not rows_segs:
                cv.configure(height=24)
                return

            th = fnt.metrics("linespace")

            x = float(pad)
            y = float(pad)
            row_h = 0
            for segs in rows_segs:
                if x > pad and x + pitch > ww - pad:
                    x = float(pad)
                    y += row_h + gap_v
                    row_h = 0
                cx = x
                line_h = th
                for i, (text, disposable_border) in enumerate(segs):
                    if not text:
                        continue
                    if i > 0:
                        cx += sp_w
                    tid = cv.create_text(
                        cx,
                        y,
                        text=text,
                        anchor=tk.NW,
                        font=FILTER_ROSTER_FONT,
                        fill="#1a1a1a",
                    )
                    bb = cv.bbox(tid)
                    if bb and disposable_border:
                        p = 2
                        rid = cv.create_rectangle(
                            bb[0] - p,
                            bb[1] - p,
                            bb[2] + p,
                            bb[3] + p,
                            outline="#0D47A1",
                            width=1,
                            fill="",
                        )
                        cv.tag_lower(rid, tid)
                    adv = float(bb[2] - bb[0]) if bb else float(fnt.measure(text))
                    cx += adv
                    if bb:
                        line_h = max(line_h, bb[3] - bb[1])
                x += pitch
                row_h = max(row_h, line_h)

            total_h = max(int(y + row_h + pad), 24)
            cv.configure(height=total_h)

        cv.bind("<Configure>", layout)
        parent.after_idle(layout)

        stat_frame = tk.Frame(parent, bg=bg)
        stat_frame.pack(anchor=tk.W, pady=(6, 0), fill=tk.X)

        def _stat_lbl(t: str, *, sub: bool = False) -> None:
            tk.Label(
                stat_frame,
                text=t,
                bg=bg,
                fg="#37474F" if sub else "#0D47A1",
                font=("Microsoft JhengHei UI", 11) if sub else FILTER_STAT_FONT,
                anchor=tk.W,
                justify=tk.LEFT,
            ).pack(anchor=tk.W)

        if rule.get("page_tag"):
            _stat_lbl(self._format_page_distribution_line_for_matches(matches), sub=True)
        tk.Label(
            stat_frame,
            text=self._format_block_fenji_one_line(matches, rule),
            bg=bg,
            fg="#0D47A1",
            font=FILTER_STAT_FONT,
            anchor=tk.W,
            justify=tk.LEFT,
            wraplength=0,
        ).pack(anchor=tk.W, fill=tk.X)

        return small_n, large_n, small_disp, large_disp

    def _filter_apply_scroll(self) -> None:
        self._filter_canvas.update_idletasks()
        bb = self._filter_canvas.bbox("all")
        if bb:
            self._filter_canvas.configure(scrollregion=bb)
        self._filter_canvas.yview_moveto(0)
        self._filter_bind_mousewheel_recursive(self._filter_inner)

    def _clear_primary_filter_panels(self) -> None:
        for w in self._filter_inner.winfo_children():
            w.destroy()

    def _clear_primary_filter_summary(self) -> None:
        h = getattr(self, "_filter_summary_host", None)
        if h is None:
            return
        for w in h.winfo_children():
            w.destroy()

    def _render_primary_filter_top_summary(
        self,
        uniq: dict,
        tags: list[str],
        keys_by_tag: dict[str, set],
        person_key,
    ) -> None:
        """置頂合計：資料頁為欄、小／大為列（訂單去重，與原底部合計口徑相同）。"""
        host = getattr(self, "_filter_summary_host", None)
        if host is None or not uniq:
            return

        def _footer_disposable_applies(r) -> bool:
            if not self._row_has_disposable_in_data(r):
                return False
            k = person_key(r)
            for t in tags:
                if not self._get_display_rule(t).get("disposable"):
                    continue
                if k in keys_by_tag.get(t, set()):
                    return True
            return False

        page_keys = sorted(
            {
                (getattr(r, "source_page", None) or "").strip() or "（無頁名）"
                for r in uniq.values()
            }
        )
        spg: dict[str, int] = {p: 0 for p in page_keys}
        sdg: dict[str, int] = {p: 0 for p in page_keys}
        lpg: dict[str, int] = {p: 0 for p in page_keys}
        ldg: dict[str, int] = {p: 0 for p in page_keys}
        og: dict[str, int] = {p: 0 for p in page_keys}

        for r in uniq.values():
            pg = (getattr(r, "source_page", None) or "").strip() or "（無頁名）"
            sz = headcount_size_label(self._row_headcount_str(r))
            disp = _footer_disposable_applies(r)
            if sz == "小":
                if disp:
                    sdg[pg] = sdg.get(pg, 0) + 1
                else:
                    spg[pg] = spg.get(pg, 0) + 1
            elif sz == "大":
                if disp:
                    ldg[pg] = ldg.get(pg, 0) + 1
                else:
                    lpg[pg] = lpg.get(pg, 0) + 1
            else:
                og[pg] = og.get(pg, 0) + 1

        hdr_bg = "#E3F2FD"
        cell_bg = "#FAFAFA"
        sum_bg = "#FFF8E1"
        font_hdr = ("Microsoft JhengHei UI", 10, "bold")
        font_cell = ("Microsoft JhengHei UI", 10)
        host.configure(bg="#E8EAF6")

        def _cell(parent, r: int, c: int, text: str, *, hdr: bool = False, sumc: bool = False) -> None:
            bg = hdr_bg if hdr else (sum_bg if sumc else cell_bg)
            bold = hdr or sumc
            tk.Label(
                parent,
                text=text,
                font=font_hdr if bold else font_cell,
                bg=bg,
                fg="#0D47A1" if bold else "#212121",
                padx=10,
                pady=6,
                relief=tk.FLAT,
                borderwidth=1,
                highlightthickness=1,
                highlightbackground="#CFD8DC",
            ).grid(row=r, column=c, sticky=tk.NSEW, padx=1, pady=1)

        gridf = tk.Frame(host, bg="#B0BEC5", padx=1, pady=1)
        gridf.pack(side=tk.RIGHT, anchor=tk.N, padx=2, pady=2)

        npg = len(page_keys)
        _cell(gridf, 0, 0, "", hdr=True)
        for j, pk in enumerate(page_keys):
            _cell(gridf, 0, j + 1, pk, hdr=True)
        _cell(gridf, 0, npg + 1, "合計", hdr=True)

        def _pn_cell(sp: int, sd: int) -> str:
            """一般+拋棄式筆數；後段數字後加 (拋)。"""
            return f"{sp}+{sd}(拋)"

        gs_plain = sum(spg.get(pk, 0) for pk in page_keys)
        gs_disp = sum(sdg.get(pk, 0) for pk in page_keys)
        gl_plain = sum(lpg.get(pk, 0) for pk in page_keys)
        gl_disp = sum(ldg.get(pk, 0) for pk in page_keys)

        _cell(gridf, 1, 0, "小", hdr=True)
        for j, pk in enumerate(page_keys):
            sp, sd = spg.get(pk, 0), sdg.get(pk, 0)
            _cell(gridf, 1, j + 1, _pn_cell(sp, sd))
        _cell(gridf, 1, npg + 1, _pn_cell(gs_plain, gs_disp), sumc=True)

        _cell(gridf, 2, 0, "大", hdr=True)
        for j, pk in enumerate(page_keys):
            lp, ld = lpg.get(pk, 0), ldg.get(pk, 0)
            _cell(gridf, 2, j + 1, _pn_cell(lp, ld))
        _cell(gridf, 2, npg + 1, _pn_cell(gl_plain, gl_disp), sumc=True)

        go_total = sum(og.get(pk, 0) for pk in page_keys)
        if go_total > 0:
            ri = 3
            _cell(gridf, ri, 0, "未標", hdr=True)
            for j, pk in enumerate(page_keys):
                o = og.get(pk, 0)
                _cell(gridf, ri, j + 1, str(o))
            _cell(gridf, ri, npg + 1, str(go_total), sumc=True)

        for c in range(npg + 2):
            gridf.grid_columnconfigure(c, weight=1)

    def _open_primary_tag_picker(self) -> None:
        values = list_hashtags()
        if not values:
            messagebox.showinfo(
                "主標籤篩選",
                "「# 標籤庫」目前為空。\n請至「# 標籤庫」分頁新增標籤並儲存，或在資料中使用 #單詞後執行分析以寫入庫。",
                parent=self.root,
            )
            return

        prev = set(self._filter_selected_tags) & set(values)
        dlg = tk.Toplevel(self.root)
        dlg.title("選擇主標籤")
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.geometry("840x560")
        dlg.minsize(560, 320)

        btnf = ttk.Frame(dlg, padding=8)
        btnf.pack(fill=tk.X, side=tk.BOTTOM)
        mid = ttk.Frame(dlg)
        mid.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(mid, highlightthickness=0)
        sb = ttk.Scrollbar(mid, orient=tk.VERTICAL, command=canvas.yview)
        inner = ttk.Frame(canvas, padding=6)
        win_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _ic(_e: tk.Event | None = None) -> None:
            bb = canvas.bbox("all")
            if bb:
                canvas.configure(scrollregion=bb)

        def _cc(e: tk.Event) -> None:
            canvas.itemconfigure(win_id, width=e.width)

        inner.bind("<Configure>", _ic)
        canvas.bind("<Configure>", _cc)
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        def _picker_mousewheel(event: tk.Event) -> None:
            d = getattr(event, "delta", 0)
            if d:
                canvas.yview_scroll(int(-1 * d / 120), "units")
            elif getattr(event, "num", None) == 4:
                canvas.yview_scroll(-1, "units")
            elif getattr(event, "num", None) == 5:
                canvas.yview_scroll(1, "units")

        def _bind_picker_wheel(w: tk.Widget) -> None:
            w.bind("<MouseWheel>", _picker_mousewheel)
            w.bind("<Button-4>", _picker_mousewheel)
            w.bind("<Button-5>", _picker_mousewheel)
            for c in w.winfo_children():
                _bind_picker_wheel(c)

        # 主勾選＋顯示規則；「拋棄式」勾選且資料含拋棄式時才對姓名加外框
        vars_by: dict[str, tk.IntVar] = {}
        serial_vars: dict[str, tk.IntVar] = {}
        page_tag_vars: dict[str, tk.IntVar] = {}
        name_vars: dict[str, tk.IntVar] = {}
        size_vars: dict[str, tk.IntVar] = {}
        disposable_vars: dict[str, tk.IntVar] = {}
        # 固定欄位：序號／資料頁標籤／姓名／人數標籤／拋棄式 垂直對齊
        inner.grid_columnconfigure(0, weight=0)
        inner.grid_columnconfigure(1, weight=0, minsize=52)
        for col in range(2, 7):
            inner.grid_columnconfigure(col, uniform="picker_disp", minsize=76, weight=0)

        for ri, v in enumerate(values):
            vars_by[v] = tk.IntVar(value=1 if v in prev else 0)
            r0 = self._get_display_rule(v)
            serial_vars[v] = tk.IntVar(value=1 if r0["serial"] else 0)
            page_tag_vars[v] = tk.IntVar(value=1 if r0.get("page_tag") else 0)
            name_vars[v] = tk.IntVar(value=1 if r0["name"] else 0)
            size_vars[v] = tk.IntVar(value=1 if r0["size_label"] else 0)
            disposable_vars[v] = tk.IntVar(value=1 if r0.get("disposable") else 0)
            py = 2
            tk.Checkbutton(
                inner,
                text=v,
                variable=vars_by[v],
                onvalue=1,
                offvalue=0,
                anchor=tk.W,
            ).grid(row=ri, column=0, sticky=tk.W, pady=py, padx=(0, 8))
            ttk.Label(inner, text="顯示：").grid(row=ri, column=1, sticky=tk.W, pady=py)
            for col, (lab, dct) in enumerate(
                (
                    ("序號", serial_vars),
                    ("資料頁標籤", page_tag_vars),
                    ("姓名", name_vars),
                    ("人數標籤", size_vars),
                    ("拋棄式", disposable_vars),
                ),
                start=2,
            ):
                tk.Checkbutton(
                    inner,
                    text=lab,
                    variable=dct[v],
                    onvalue=1,
                    offvalue=0,
                    anchor=tk.W,
                ).grid(row=ri, column=col, sticky=tk.W, pady=py, padx=(4, 0))

        _bind_picker_wheel(inner)
        for w, evs in (
            (canvas, ("<MouseWheel>", "<Button-4>", "<Button-5>")),
            (sb, ("<MouseWheel>", "<Button-4>", "<Button-5>")),
            (mid, ("<MouseWheel>", "<Button-4>", "<Button-5>")),
        ):
            for ev in evs:
                w.bind(ev, _picker_mousewheel)

        def _sel_all() -> None:
            for iv in vars_by.values():
                iv.set(1)

        def _clr_all() -> None:
            for iv in vars_by.values():
                iv.set(0)

        def _ok() -> None:
            chosen = [v for v in values if int(vars_by[v].get() or 0) == 1]
            new_rules: dict[str, dict[str, bool]] = dict(self._filter_display_rules)
            for v in values:
                new_rules[v] = normalize_display_rule(
                    {
                        "serial": int(serial_vars[v].get() or 0) == 1,
                        "page_tag": int(page_tag_vars[v].get() or 0) == 1,
                        "name": int(name_vars[v].get() or 0) == 1,
                        "size_label": int(size_vars[v].get() or 0) == 1,
                        "disposable": int(disposable_vars[v].get() or 0) == 1,
                    }
                )
            self._filter_selected_tags = chosen
            self._filter_display_rules = new_rules
            try:
                save_filter_prefs(
                    chosen,
                    new_rules,
                    self._filter_export_blocks,
                    self._get_export_templates_live(),
                )
            except OSError as e:
                messagebox.showwarning("主標籤篩選", f"無法儲存勾選與顯示規則：{e}", parent=self.root)
            dlg.destroy()
            self._nb.select(self._tab_primary_filter)
            self._rebuild_primary_filter_results()
            if not chosen:
                messagebox.showinfo(
                    "主標籤篩選",
                    "未勾選任何標籤。請至少勾選一項後按「確定」。",
                    parent=self.root,
                )
            self._status.set(f"主標籤篩選：已選 {len(self._filter_selected_tags)} 個標籤。")

        def _cancel() -> None:
            dlg.destroy()

        ttk.Button(btnf, text="全選", command=_sel_all).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btnf, text="全不選", command=_clr_all).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Button(btnf, text="確定", command=_ok).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(btnf, text="取消", command=_cancel).pack(side=tk.RIGHT)
        dlg.bind("<Escape>", lambda _e: _cancel())

    def _refresh_primary_filter_results(self) -> None:
        if not self._filter_selected_tags:
            messagebox.showinfo("主標籤篩選", "請先按「選擇主標籤…」從標籤庫勾選要檢視的標籤。", parent=self.root)
            return
        if not self._rows:
            messagebox.showinfo("主標籤篩選", "請先在「輸入與分析」完成分析。", parent=self.root)
            return
        self._rebuild_primary_filter_results()
        self._status.set("已依目前選取更新名單。")

    def _rebuild_primary_filter_results(self) -> None:
        self._clear_primary_filter_panels()
        self._clear_primary_filter_summary()
        self._filter_export_vars.clear()
        if not self._filter_selected_tags:
            tk.Label(
                self._filter_inner,
                text="請按「選擇主標籤…」從「# 標籤庫」勾選主標籤後，將在此顯示各標籤對應名單。",
                bg=FILTER_INNER_BG,
                fg="#333333",
                justify=tk.LEFT,
                wraplength=780,
            ).pack(anchor=tk.W, pady=8)
            self._filter_apply_scroll()
            return

        if not self._rows:
            tk.Label(
                self._filter_inner,
                text="已選主標籤，但尚無分析資料。請至「輸入與分析」貼上資料並執行分析後，再按「依目前選取更新名單」。",
                bg=FILTER_INNER_BG,
                fg="#333333",
                justify=tk.LEFT,
                wraplength=780,
            ).pack(anchor=tk.W, pady=8)
            self._filter_apply_scroll()
            return

        lib_list = list_hashtags()
        lib = set(lib_list)
        order_map = {name: i for i, name in enumerate(lib_list)}
        tags = [t for t in self._filter_selected_tags if t in lib]
        tags.sort(key=lambda t: order_map.get(t, 10**9))
        if not tags:
            tk.Label(
                self._filter_inner,
                text="所勾選的標籤已不在「# 標籤庫」中，請按「選擇主標籤…」重新勾選。",
                bg=FILTER_INNER_BG,
                fg="#333333",
                justify=tk.LEFT,
                wraplength=780,
            ).pack(anchor=tk.W, pady=8)
            self._filter_apply_scroll()
            return

        if all(len(self._rows_matching_tag_value(t)) == 0 for t in tags):
            tk.Label(
                self._filter_inner,
                text="提示：所選標籤皆為 0 筆時，請確認「標籤庫」內文字與解析結果 JSON 裡 tags[].value 完全一致（含全形／半形括號、空白）。",
                bg=FILTER_INNER_BG,
                fg="#333333",
                justify=tk.LEFT,
                wraplength=780,
            ).pack(anchor=tk.W, pady=(0, 10))

        unified_pitch = self._filter_unified_roster_pitch(tags)

        def _person_key(r) -> tuple[str, str]:
            return (str(r.serial).strip(), (r.customer_name or "").strip())

        uniq: dict[tuple[str, str], object] = {}
        for tag in tags:
            for r in self._rows_matching_tag_value(tag):
                uniq[_person_key(r)] = r
        keys_by_tag: dict[str, set[tuple[str, str]]] = {
            t: {_person_key(r) for r in self._rows_matching_tag_value(t)} for t in tags
        }
        if uniq:
            self._render_primary_filter_top_summary(uniq, tags, keys_by_tag, _person_key)

        bi = 0
        for tag in tags:
            matches = self._rows_matching_tag_value(tag)
            if not matches:
                continue
            bg = FILTER_BLOCK_BGS[bi % len(FILTER_BLOCK_BGS)]
            bi += 1
            block = tk.Frame(
                self._filter_inner,
                bg=bg,
                highlightbackground="#B0BEC5",
                highlightthickness=1,
            )
            block.pack(fill=tk.X, expand=False, pady=(0, 12), padx=2)

            head = tk.Frame(block, bg=bg)
            head.pack(fill=tk.X, padx=8, pady=(10, 6))
            init_export = bool(self._filter_export_blocks.get(tag, True))
            ex_var = tk.IntVar(value=1 if init_export else 0)
            self._filter_export_vars[tag] = ex_var

            def _sync_export(tn: str = tag, v: tk.IntVar = ex_var) -> None:
                self._filter_export_blocks[tn] = int(v.get() or 0) == 1
                try:
                    save_filter_prefs(
                        self._filter_selected_tags,
                        self._filter_display_rules,
                        self._filter_export_blocks,
                        self._get_export_templates_live(),
                    )
                except OSError:
                    pass

            tk.Checkbutton(
                head,
                text="輸出",
                variable=ex_var,
                onvalue=1,
                offvalue=0,
                command=_sync_export,
                font=FILTER_EXPORT_CHECK_FONT,
                bg=bg,
                fg="#0D47A1",
                anchor=tk.W,
                activebackground=bg,
                activeforeground="#0D47A1",
                selectcolor="#FFFFFF",
            ).pack(side=tk.LEFT, padx=(4, 10))
            tk.Label(
                head,
                text=tag,
                bg=bg,
                fg="#0D47A1",
                font=("Microsoft JhengHei UI", 13, "bold"),
                anchor=tk.W,
            ).pack(side=tk.LEFT, anchor=tk.W)

            body = tk.Frame(block, bg=bg)
            body.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 6))
            self._render_compact_roster(
                body, matches, bg, unified_pitch, self._get_display_rule(tag)
            )

        self._filter_apply_scroll()

    def _visible_primary_filter_tags(self) -> list[str]:
        """目前會畫出的主標籤（有命中筆數），順序同標籤庫。"""
        if not self._rows or not self._filter_selected_tags:
            return []
        lib_list = list_hashtags()
        lib = set(lib_list)
        order_map = {name: i for i, name in enumerate(lib_list)}
        tags = [t for t in self._filter_selected_tags if t in lib]
        tags.sort(key=lambda t: order_map.get(t, 10**9))
        return [t for t in tags if self._rows_matching_tag_value(t)]

    def _tags_checked_for_export(self, visible: list[str]) -> list[str]:
        out: list[str] = []
        for t in visible:
            iv = self._filter_export_vars.get(t)
            if iv is not None:
                if int(iv.get() or 0) == 1:
                    out.append(t)
            elif self._filter_export_blocks.get(t, True):
                out.append(t)
        return out

    def _current_export_tags_subset(self) -> list[str]:
        """與目前匯出預覽一致的標籤範圍（未勾「輸出」時視同輸出全部可見區塊）。"""
        vis = self._visible_primary_filter_tags()
        if not vis:
            return []
        picked = self._tags_checked_for_export(vis)
        return picked if picked else list(vis)

    def _ask_export_all_or_cancel(self) -> str | None:
        """全未勾時：回傳 'all' 表示輸出全部有資料區塊，None 為取消。"""
        result: list[str | None] = [None]
        dlg = tk.Toplevel(self.root)
        dlg.title("匯出")
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)
        ttk.Label(
            dlg,
            text="尚未勾選任何「輸出」區塊。\n是否改為輸出目前畫面上「所有有資料的區塊」？",
            wraplength=380,
        ).pack(padx=16, pady=(16, 8))
        bf = ttk.Frame(dlg, padding=(8, 0, 8, 16))

        def _all() -> None:
            result[0] = "all"
            dlg.destroy()

        def _cancel() -> None:
            result[0] = None
            dlg.destroy()

        ttk.Button(bf, text="輸出全部有資料區塊", command=_all).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(bf, text="取消", command=_cancel).pack(side=tk.LEFT)
        bf.pack()
        dlg.bind("<Escape>", lambda _e: _cancel())
        dlg.wait_window()
        return result[0]

    def _open_export_from_primary_filter(self) -> None:
        vis = self._visible_primary_filter_tags()
        if not vis:
            messagebox.showinfo(
                "匯出",
                "請先完成「輸入與分析」，並在「主標籤篩選」顯示至少一個有資料的區塊。",
                parent=self.root,
            )
            return
        picked = self._tags_checked_for_export(vis)
        if not picked:
            if self._ask_export_all_or_cancel() == "all":
                picked = list(vis)
            else:
                return
        self._refresh_export_tab_preview(picked)
        self._nb.select(self._tab_export)

    @staticmethod
    def _format_page_distribution_line_for_matches(matches: list) -> str:
        """依資料頁（source_page）分計筆數；各頁加總＝本區名單筆數。"""
        if not matches:
            return "資料頁分計：（無）"
        c = Counter(
            (getattr(r, "source_page", None) or "").strip() or "（無頁名）" for r in matches
        )
        parts = [f"「{k}」{v} 筆" for k, v in sorted(c.items(), key=lambda x: x[0])]
        n = len(matches)
        return "資料頁分計：" + "　".join(parts) + f"　（小計 {n} 筆）"

    def _count_size_breakdown(
        self, matches: list, rule: dict[str, bool]
    ) -> tuple[int, int, int, int, int]:
        """回傳 (小筆數, 大筆數, 小含拋棄式, 大含拋棄式, 未標示份量筆數)。"""
        rule = normalize_display_rule(rule)
        small_n = large_n = small_disp = large_disp = other_n = 0
        use_disp = bool(rule["disposable"])
        for r in matches:
            sz = headcount_size_label(self._row_headcount_str(r))
            hit = use_disp and self._row_has_disposable_in_data(r)
            if sz == "小":
                small_n += 1
                if hit:
                    small_disp += 1
            elif sz == "大":
                large_n += 1
                if hit:
                    large_disp += 1
            else:
                other_n += 1
        return small_n, large_n, small_disp, large_disp, other_n

    def _format_block_fenji_one_line(self, matches: list, rule: dict[str, bool]) -> str:
        """區塊內分量單行：小／大／合計（一般+拋），與主畫面色系一致之 #0D47A1。"""
        small_n, large_n, small_disp, large_disp, other_n = self._count_size_breakdown(
            matches, rule
        )
        sp = small_n - small_disp
        lp = large_n - large_disp
        sd, ld = small_disp, large_disp
        tot_p, tot_d = sp + lp, sd + ld
        s = (
            f"分量分計 : 小 : {sp}+{sd}(拋)    大 : {lp}+{ld}(拋)    "
            f"合計: {tot_p}+{tot_d}(拋)"
        )
        if other_n:
            s += f"    未標:{other_n}"
        return s

    def _primary_filter_block_stats_lines(
        self, matches: list, rule: dict[str, bool]
    ) -> list[str]:
        """區塊統計：（可選）資料頁、分量單行。"""
        rule = normalize_display_rule(rule)
        out: list[str] = []
        if rule.get("page_tag"):
            out.append(self._format_page_distribution_line_for_matches(matches))
        out.append(self._format_block_fenji_one_line(matches, rule))
        return out

    def _primary_filter_block_stat_text(self, matches: list, rule: dict[str, bool]) -> str:
        """多行合併字串（匯出純文字／相容舊呼叫）。"""
        return "\n".join(self._primary_filter_block_stats_lines(matches, rule))

    def _export_footer_text_for_tags(
        self, tags: list[str], *, visible_tags: list[str] | None = None
    ) -> str:
        if not tags:
            return ""

        def _person_key(r) -> tuple[str, str]:
            return (str(r.serial).strip(), (r.customer_name or "").strip())

        uniq: dict[tuple[str, str], object] = {}
        for tag in tags:
            for r in self._rows_matching_tag_value(tag):
                uniq[_person_key(r)] = r
        keys_by_tag = {t: {_person_key(r) for r in self._rows_matching_tag_value(t)} for t in tags}

        def _footer_disposable_applies(r) -> bool:
            if not self._row_has_disposable_in_data(r):
                return False
            k = _person_key(r)
            for t in tags:
                if not self._get_display_rule(t).get("disposable"):
                    continue
                if k in keys_by_tag.get(t, set()):
                    return True
            return False

        page_keys = sorted(
            {
                (getattr(r, "source_page", None) or "").strip() or "（無頁名）"
                for r in uniq.values()
            }
        )
        spg: dict[str, int] = {p: 0 for p in page_keys}
        sdg: dict[str, int] = {p: 0 for p in page_keys}
        lpg: dict[str, int] = {p: 0 for p in page_keys}
        ldg: dict[str, int] = {p: 0 for p in page_keys}
        og: dict[str, int] = {p: 0 for p in page_keys}

        for r in uniq.values():
            pg = (getattr(r, "source_page", None) or "").strip() or "（無頁名）"
            sz = headcount_size_label(self._row_headcount_str(r))
            disp = _footer_disposable_applies(r)
            if sz == "小":
                if disp:
                    sdg[pg] += 1
                else:
                    spg[pg] += 1
            elif sz == "大":
                if disp:
                    ldg[pg] += 1
                else:
                    lpg[pg] += 1
            else:
                og[pg] += 1

        gs_plain = sum(spg.get(pk, 0) for pk in page_keys)
        gs_disp = sum(sdg.get(pk, 0) for pk in page_keys)
        gl_plain = sum(lpg.get(pk, 0) for pk in page_keys)
        gl_disp = sum(ldg.get(pk, 0) for pk in page_keys)
        go_total = sum(og.get(pk, 0) for pk in page_keys)

        def _pn(sp: int, sd: int) -> str:
            return f"{sp}+{sd}(拋)"

        out_lines = [
            "\t".join(["", *page_keys, "合計"]),
            "\t".join(["小"] + [_pn(spg[pk], sdg[pk]) for pk in page_keys] + [_pn(gs_plain, gs_disp)]),
            "\t".join(["大"] + [_pn(lpg[pk], ldg[pk]) for pk in page_keys] + [_pn(gl_plain, gl_disp)]),
        ]
        if go_total > 0:
            out_lines.append(
                "\t".join(
                    ["未標"] + [str(og.get(pk, 0)) for pk in page_keys] + [str(go_total)],
                )
            )
        return "\n".join(out_lines)

    @staticmethod
    def _export_cell_visual_width(s: str) -> int:
        return sum(2 if ord(ch) > 127 else 1 for ch in s)

    def _export_pad_visual(self, s: str, min_w: int) -> str:
        w = self._export_cell_visual_width(s)
        if w >= min_w:
            return s
        return s + " " * (min_w - w)

    def _export_row_cells(self, r, rule: dict[str, bool]) -> tuple[str, str, str, str]:
        rule = normalize_display_rule(rule)
        ser = format_order_serial(r.serial) if rule["serial"] else ""
        pg = ""
        if rule.get("page_tag"):
            pg = (getattr(r, "source_page", None) or "").strip()
        nm = ""
        if rule["name"]:
            nm = (r.customer_name or "").strip() or "（無姓名）"
        sz_lab = headcount_size_label(self._row_headcount_str(r))
        szs = f"({sz_lab})" if (rule["size_label"] and sz_lab) else ""
        return ser, pg, nm, szs

    def _export_row_display_line(self, r, rule: dict[str, bool]) -> str:
        segs = self._roster_segments(r, rule)
        return " ".join(tx for tx, _ in segs)

    def _export_roster_plain_width(self, r, rule: dict[str, bool]) -> float:
        """與主篩選 Canvas 相同之顯示寬度（含拋棄式外框預留寬），供流式匯出欄距。"""
        return self._roster_plain_width(r, rule)

    def _filter_unified_roster_pitch_for_plain_export(self, tags: list[str]) -> float:
        widths: list[float] = []
        for tag in tags:
            rule = self._get_display_rule(tag)
            for r in self._rows_matching_tag_value(tag):
                widths.append(self._export_roster_plain_width(r, rule))
        if not widths:
            return max(48 + FILTER_ROSTER_ITEM_GAP_H, 80)
        max_tw = max(widths)
        return float(max(max_tw + FILTER_ROSTER_ITEM_GAP_H, 80))

    def _export_filter_block_content_width_px(self) -> int:
        """與主標籤篩選 Canvas 接近的寬度，供流式排版換行用。"""
        try:
            self.root.update_idletasks()
            for w in (getattr(self, "_filter_canvas", None), getattr(self, "_filter_inner", None)):
                if w is None:
                    continue
                ww = int(w.winfo_width())
                if ww > 80:
                    return ww
        except tk.TclError:
            pass
        return 780

    @staticmethod
    def _export_pixel_grid_line(cells: list[str], pitch: float, pad: int, fnt: tkfont.Font) -> str:
        """將同一視覺列上的名單儲格左對齊在固定像素間距（與篩選區 Canvas 邏輯一致）。"""
        line = ""
        for i, cell in enumerate(cells):
            target_x = float(pad) + float(i) * pitch
            while float(fnt.measure(line)) < target_x:
                line += " "
            line += cell
        return line.rstrip()

    def _export_filter_roster_table_row_groups(
        self, matches: list, rule: dict[str, bool], roster_pitch: float, content_w: int
    ) -> list[list]:
        """與篩選區流式名單相同的儲格分列（每格一筆 ParsedRow），供 Word 表格排版。"""
        if not matches:
            return []
        pad = int(FILTER_ROSTER_PAD)
        pitch = float(roster_pitch)
        ww = max(int(content_w), 120)
        rows: list[list] = []
        cur: list = []
        x = float(pad)
        for r in matches:
            if x > pad and x + pitch > float(ww) - float(pad):
                rows.append(cur)
                cur = []
                x = float(pad)
            cur.append(r)
            x += pitch
        if cur:
            rows.append(cur)
        return rows

    def _export_filter_like_roster_lines(
        self, matches: list, rule: dict[str, bool], roster_pitch: float, content_w: int
    ) -> list[str]:
        """模擬 _render_compact_roster 的換行與欄距，產出純文字列。"""
        fnt = tkfont.Font(font=FILTER_ROSTER_FONT)
        pad = int(FILTER_ROSTER_PAD)
        pitch = float(roster_pitch)
        groups = self._export_filter_roster_table_row_groups(matches, rule, roster_pitch, content_w)
        return [
            self._export_pixel_grid_line(
                [self._export_row_display_line(r, rule) for r in grp],
                pitch,
                pad,
                fnt,
            )
            for grp in groups
        ]

    @staticmethod
    def _export_format_safe(template: str, mapping: dict[str, str]) -> str:
        class _Safe(dict):
            def __missing__(self, key: str) -> str:
                return ""

        try:
            return str(template).format_map(_Safe(mapping))
        except ValueError:
            return str(template) + "〈格式錯誤：請檢查大括號是否成對〉"

    def _get_export_templates_live(self) -> dict[str, str]:
        win = getattr(self, "_win_export_custom", None)
        if win is not None and win.winfo_exists():
            tb = getattr(self, "_txt_export_custom_block", None)
            tr = getattr(self, "_txt_export_custom_row", None)
            if tb is not None and tr is not None:
                block = tb.get("1.0", "end-1c")
                row = tr.get("1.0", "end-1c")
                return normalize_export_templates({"custom_block": block, "custom_row": row})
        return normalize_export_templates(self._export_custom_templates)

    def _maybe_refresh_custom_export_preview(self) -> None:
        if not hasattr(self, "_var_export_layout_ui"):
            return
        layout = self._export_layout_label_to_key.get(self._var_export_layout_ui.get(), "screen")
        if layout == "custom":
            self._refresh_export_tab_preview(None)

    def _export_custom_insert(self, target: tk.Text, fragment: str) -> None:
        target.insert(tk.INSERT, fragment)
        target.see(tk.INSERT)
        target.focus_set()
        self._maybe_refresh_custom_export_preview()

    def _export_custom_replace_all(self, target: tk.Text, content: str) -> None:
        target.delete("1.0", tk.END)
        target.insert("1.0", content)
        target.focus_set()
        self._maybe_refresh_custom_export_preview()

    def _build_export_custom_toolbar(self, parent: tk.Widget, target: tk.Text, *, mode: str) -> None:
        """mode: 'block' | 'row' — 插入占位符與一鍵範本（兩列：插入／範本）。"""
        wrap = ttk.Frame(parent)
        wrap.pack(fill=tk.X, pady=(0, 4))
        bar_ins = ttk.Frame(wrap)
        bar_ins.pack(fill=tk.X, pady=(0, 3))
        bar_pre = ttk.Frame(wrap)
        bar_pre.pack(fill=tk.X)
        ttk.Label(bar_ins, text="插入：").pack(side=tk.LEFT, padx=(0, 4))
        if mode == "block":
            insert_pairs: tuple[tuple[str, str], ...] = (
                ("{tag}", "{tag}"),
                ("{count}", "{count}"),
                ("Tab", "\t"),
                ("換行", "\n"),
            )
            presets: tuple[tuple[str, str], ...] = (
                ("預設標題", DEFAULT_EXPORT_CUSTOM_TEMPLATES["custom_block"]),
                ("簡式", "{tag}：共 {count} 筆"),
                ("僅標籤行", "{tag}"),
                ("清空", ""),
            )
        elif mode == "row":
            insert_pairs = (
                ("{serial}", "{serial}"),
                ("{page}", "{page}"),
                ("{name}", "{name}"),
                ("{size}", "{size}"),
                ("Tab", "\t"),
                ("換行", "\n"),
            )
            presets = (
                ("預設列", DEFAULT_EXPORT_CUSTOM_TEMPLATES["custom_row"]),
                ("CSV", "{serial},{page},{name},{size}"),
                ("直線", "{serial} | {page} | {name} | {size}"),
                ("僅姓名", "{name}"),
                ("清空", ""),
            )
        else:
            return
        for cap, frag in insert_pairs:
            ttk.Button(
                bar_ins,
                text=cap,
                command=lambda t=target, fr=frag: self._export_custom_insert(t, fr),
            ).pack(side=tk.LEFT, padx=(0, 3))
        ttk.Label(bar_pre, text="整段範本：").pack(side=tk.LEFT, padx=(0, 4))
        for cap, body in presets:
            ttk.Button(
                bar_pre,
                text=cap,
                command=lambda t=target, c=body: self._export_custom_replace_all(t, c),
            ).pack(side=tk.LEFT, padx=(0, 3))

    def _save_custom_export_templates(self, *, parent: tk.Misc | None = None) -> None:
        self._export_custom_templates = self._get_export_templates_live()
        par = parent if parent is not None else self.root
        try:
            save_filter_prefs(
                self._filter_selected_tags,
                self._filter_display_rules,
                self._filter_export_blocks,
                self._export_custom_templates,
            )
        except OSError as e:
            messagebox.showerror("匯出", f"無法儲存自訂格式：{e}", parent=par)
            return
        self._status.set("已儲存自訂匯出格式。")
        layout = self._export_layout_label_to_key.get(self._var_export_layout_ui.get(), "screen")
        if layout == "custom":
            self._refresh_export_tab_preview(None)

    def _sync_export_custom_dialog_to_memory(self) -> None:
        win = getattr(self, "_win_export_custom", None)
        if win is None or not win.winfo_exists():
            return
        tb = getattr(self, "_txt_export_custom_block", None)
        tr = getattr(self, "_txt_export_custom_row", None)
        if tb is None or tr is None:
            return
        self._export_custom_templates = normalize_export_templates(
            {
                "custom_block": tb.get("1.0", "end-1c"),
                "custom_row": tr.get("1.0", "end-1c"),
            }
        )

    def _close_export_custom_dialog(self) -> None:
        win = getattr(self, "_win_export_custom", None)
        if win is None or not win.winfo_exists():
            return
        self._sync_export_custom_dialog_to_memory()
        self._win_export_custom = None
        win.destroy()
        self._maybe_refresh_custom_export_preview()

    def _open_export_custom_format_dialog(self) -> None:
        existing = getattr(self, "_win_export_custom", None)
        if existing is not None and existing.winfo_exists():
            existing.lift()
            existing.focus_force()
            return

        win = tk.Toplevel(self.root)
        win.title("自訂匯出格式")
        win.transient(self.root)
        win.minsize(520, 380)
        self._win_export_custom = win

        outer = ttk.Frame(win, padding=10)
        outer.pack(fill=tk.BOTH, expand=True)
        ttk.Label(
            outer,
            text=(
                "排列方式選「自訂」時會套用此處內容。按鈕可插入占位符或整段範本；字面大括號請寫 {{ 與 }}。"
                "欄位依「主標籤篩選」顯示規則，未勾欄位為空。"
                "拋棄式姓名在 Word 表格為藍色字元框線（與主篩選外框一致）；純文字／自訂 {name} 為一般姓名不加括號。"
            ),
            wraplength=640,
        ).pack(anchor=tk.W, pady=(0, 10))

        ttk.Label(outer, text="區塊標題（每區一則，可多行）").pack(anchor=tk.W)
        self._txt_export_custom_block = tk.Text(outer, height=3, wrap=tk.WORD, font=("Consolas", 10))
        cb0 = self._export_custom_templates.get("custom_block", DEFAULT_EXPORT_CUSTOM_TEMPLATES["custom_block"])
        self._txt_export_custom_block.insert("1.0", cb0)
        self._build_export_custom_toolbar(outer, self._txt_export_custom_block, mode="block")
        self._txt_export_custom_block.pack(fill=tk.BOTH, expand=False, pady=(0, 12))

        ttk.Label(outer, text="每一筆資料（可多行）").pack(anchor=tk.W)
        self._txt_export_custom_row = tk.Text(outer, height=5, wrap=tk.WORD, font=("Consolas", 10))
        cr0 = self._export_custom_templates.get("custom_row", DEFAULT_EXPORT_CUSTOM_TEMPLATES["custom_row"])
        self._txt_export_custom_row.insert("1.0", cr0)
        self._build_export_custom_toolbar(outer, self._txt_export_custom_row, mode="row")
        self._txt_export_custom_row.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        btnf = ttk.Frame(outer)
        btnf.pack(fill=tk.X)
        ttk.Button(
            btnf,
            text="儲存到設定檔",
            command=lambda: self._save_custom_export_templates(parent=win),
        ).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btnf, text="關閉", command=self._close_export_custom_dialog).pack(side=tk.RIGHT)

        win.protocol("WM_DELETE_WINDOW", self._close_export_custom_dialog)

    def _build_primary_filter_export_text(
        self,
        tags_subset: list[str],
        *,
        include_block_stats: bool,
        include_footer: bool,
        layout: str,
    ) -> str:
        lib_list = list_hashtags()
        order_map = {name: i for i, name in enumerate(lib_list)}
        tags = sorted(tags_subset, key=lambda t: order_map.get(t, 10**9))
        lines: list[str] = []
        tpl_custom: dict[str, str] | None = None
        if layout == "custom":
            tpl_custom = self._get_export_templates_live()

        unified_pitch_screen = 80.0
        content_w_screen = 780
        if layout == "screen":
            pitch_basis = self._visible_primary_filter_tags() or list(tags_subset)
            unified_pitch_screen = self._filter_unified_roster_pitch_for_plain_export(pitch_basis)
            content_w_screen = self._export_filter_block_content_width_px()

        for tag in tags:
            matches = self._rows_matching_tag_value(tag)
            if not matches:
                continue
            rule = self._get_display_rule(tag)

            if layout == "custom" and tpl_custom is not None:
                cb = tpl_custom.get("custom_block") or DEFAULT_EXPORT_CUSTOM_TEMPLATES["custom_block"]
                cr = tpl_custom.get("custom_row") or DEFAULT_EXPORT_CUSTOM_TEMPLATES["custom_row"]
                lines.append(
                    self._export_format_safe(
                        cb,
                        {"tag": tag, "count": str(len(matches))},
                    )
                )
                for r in matches:
                    a, b, c, d = self._export_row_cells(r, rule)
                    lines.append(
                        self._export_format_safe(
                            cr,
                            {"serial": a, "page": b, "name": c, "size": d},
                        )
                    )
            else:
                if layout == "screen":
                    lines.append(tag)
                    lines.extend(
                        self._export_filter_like_roster_lines(
                            matches, rule, unified_pitch_screen, content_w_screen
                        )
                    )
                else:
                    lines.append(f"【{tag}】")

                if layout == "tsv":
                    lines.append("序號\t資料頁\t姓名\t人數標籤")
                    for r in matches:
                        a, b, c, d = self._export_row_cells(r, rule)
                        lines.append("\t".join((a, b, c, d)))
                elif layout == "print_cols":
                    lines.append(
                        "  "
                        + self._export_pad_visual("序號", 6)
                        + "  "
                        + self._export_pad_visual("資料頁", 14)
                        + "  "
                        + self._export_pad_visual("姓名", 18)
                        + "  "
                        + "人數"
                    )
                    lines.append("  " + "-" * 52)
                    for r in matches:
                        a, b, c, d = self._export_row_cells(r, rule)
                        lines.append(
                            "  "
                            + self._export_pad_visual(a, 6)
                            + "  "
                            + self._export_pad_visual(b, 14)
                            + "  "
                            + self._export_pad_visual(c, 18)
                            + "  "
                            + d
                        )
                elif layout in ("flow3", "flow4"):
                    ncols = 3 if layout == "flow3" else 4
                    chunks = [self._export_row_display_line(r, rule) for r in matches]
                    for i in range(0, len(chunks), ncols):
                        lines.append("  │  ".join(chunks[i : i + ncols]))
                elif layout == "names":
                    for r in matches:
                        rule_n = normalize_display_rule(rule)
                        if not rule_n["name"]:
                            continue
                        nm = (r.customer_name or "").strip() or "（無姓名）"
                        lines.append("  " + nm)
                elif layout != "screen":
                    for r in matches:
                        lines.append("  " + self._export_row_display_line(r, rule))

            if include_block_stats:
                stat_lines = self._primary_filter_block_stats_lines(matches, rule)
                if layout == "screen":
                    lines.extend(stat_lines)
                else:
                    for sl in stat_lines:
                        lines.append("  " + sl)
            lines.append("")
        if include_footer and tags:
            vis_foot = self._visible_primary_filter_tags()
            lines.append(
                self._export_footer_text_for_tags(
                    tags,
                    visible_tags=vis_foot if vis_foot else None,
                )
            )
            lines.append("")
        return "\n".join(lines).rstrip() + "\n"

    @staticmethod
    def _open_path_in_default_app(path: Path) -> None:
        p = str(path.resolve())
        if sys.platform == "win32":
            os.startfile(p)
        elif sys.platform == "darwin":
            subprocess.run(["open", p], check=False)
        else:
            subprocess.run(["xdg-open", p], check=False)

    @staticmethod
    def _reveal_path_in_file_manager(path: Path) -> None:
        p = str(path.resolve())
        if sys.platform == "win32":
            subprocess.run(["explorer", "/select,", p], check=False)
        elif sys.platform == "darwin":
            subprocess.run(["open", "-R", p], check=False)
        else:
            subprocess.run(["xdg-open", str(path.parent)], check=False)

    def _write_export_docx_file(
        self,
        dest: str | Path,
        *,
        tags_subset: list[str],
        layout_key: str,
        include_block_stats: bool,
        include_footer: bool,
    ) -> None:
        from export_preview_docx import save_preview_text_as_docx, save_screen_layout_docx

        dest = Path(dest)
        if layout_key == "screen":
            save_screen_layout_docx(
                dest,
                self,
                tags_subset,
                include_block_stats=include_block_stats,
                include_footer=include_footer,
            )
        else:
            text = self._build_primary_filter_export_text(
                tags_subset,
                include_block_stats=include_block_stats,
                include_footer=include_footer,
                layout=layout_key,
            )
            save_preview_text_as_docx(dest, text)

    def _export_preview_set_idle(self, message: str, *, error: bool = False) -> None:
        self._export_preview_docx_path = None
        self._export_last_preview_tags = None
        self._export_last_preview_layout = None
        self._export_last_preview_inc_stats = None
        self._export_last_preview_inc_foot = None
        if hasattr(self, "_lbl_export_preview_status"):
            self._lbl_export_preview_status.configure(
                text=message,
                fg="#C62828" if error else "#333333",
            )
        if hasattr(self, "_lbl_export_preview_path"):
            self._lbl_export_preview_path.configure(text="", fg="#555555")
        if hasattr(self, "_export_preview_action_buttons"):
            for b in self._export_preview_action_buttons:
                b.configure(state="disabled")

    def _export_preview_set_ready(self, path: Path, tags_subset: list[str], layout_key: str) -> None:
        self._export_preview_docx_path = path
        self._export_last_preview_tags = list(tags_subset)
        self._export_last_preview_layout = layout_key
        self._export_last_preview_inc_stats = bool(int(self._var_export_block_stats.get() or 0))
        self._export_last_preview_inc_foot = bool(int(self._var_export_footer.get() or 0))
        if hasattr(self, "_lbl_export_preview_status"):
            self._lbl_export_preview_status.configure(
                text="預覽已產生。請按「用預設程式開啟」在 Word 中檢視，或另存到你的資料夾。",
                fg="#1B5E20",
            )
        if hasattr(self, "_lbl_export_preview_path"):
            self._lbl_export_preview_path.configure(text=str(path.resolve()), fg="#0D47A1")
        if hasattr(self, "_export_preview_action_buttons"):
            for b in self._export_preview_action_buttons:
                b.configure(state="normal")

    def _on_export_tab_option_changed(self) -> None:
        if self._export_preview_docx_path is not None and self._export_preview_docx_path.is_file():
            self._refresh_export_tab_preview(None)

    def _refresh_export_tab_preview(self, tags_subset: list[str] | None = None) -> None:
        try:
            from export_preview_docx import save_screen_layout_docx  # noqa: F401
        except ImportError:
            self._export_preview_set_idle(
                "無法產生 Word 預覽：請先安裝 python-docx\n（在終端執行：pip install python-docx）",
                error=True,
            )
            return

        if tags_subset is None:
            vis = self._visible_primary_filter_tags()
            tags_subset = self._tags_checked_for_export(vis)
            if not vis:
                self._export_preview_set_idle(
                    "請先完成分析並在「主標籤篩選」顯示區塊後，再按「產生／更新 Word 預覽」。"
                )
                return
            if not tags_subset:
                if self._ask_export_all_or_cancel() == "all":
                    tags_subset = list(vis)
                else:
                    return

        inc_stats = bool(int(self._var_export_block_stats.get() or 0))
        inc_foot = bool(int(self._var_export_footer.get() or 0))
        layout_key = self._export_layout_label_to_key.get(
            self._var_export_layout_ui.get(),
            "screen",
        )

        tmp = Path(tempfile.gettempdir()) / f"manut_export_preview_{os.getpid()}.docx"
        try:
            self._write_export_docx_file(
                tmp,
                tags_subset=tags_subset,
                layout_key=layout_key,
                include_block_stats=inc_stats,
                include_footer=inc_foot,
            )
        except OSError as e:
            messagebox.showerror("匯出", f"無法寫入預覽檔：{e}", parent=self.root)
            return
        except Exception as e:
            messagebox.showerror("匯出", f"無法產生 Word 預覽：{e}", parent=self.root)
            return

        self._export_preview_set_ready(tmp, tags_subset, layout_key)
        self._status.set(f"已更新 Word 預覽：{tmp}")

    def _export_open_docx_preview(self) -> None:
        p = self._export_preview_docx_path
        if p is None or not p.is_file():
            messagebox.showinfo("匯出", "請先按「產生／更新 Word 預覽」。", parent=self.root)
            return
        try:
            self._open_path_in_default_app(p)
        except OSError as e:
            messagebox.showerror("匯出", f"無法開啟檔案：{e}", parent=self.root)

    def _export_reveal_docx_preview(self) -> None:
        p = self._export_preview_docx_path
        if p is None or not p.is_file():
            messagebox.showinfo("匯出", "請先按「產生／更新 Word 預覽」。", parent=self.root)
            return
        try:
            self._reveal_path_in_file_manager(p)
        except OSError as e:
            messagebox.showerror("匯出", f"無法開啟檔案總管：{e}", parent=self.root)

    def _export_save_docx_copy_as(self) -> None:
        p = self._export_preview_docx_path
        if p is None or not p.is_file():
            messagebox.showinfo("匯出", "請先按「產生／更新 Word 預覽」。", parent=self.root)
            return
        dest = filedialog.asksaveasfilename(
            parent=self.root,
            defaultextension=".docx",
            filetypes=[("Word 文件 (*.docx)", "*.docx"), ("全部", "*.*")],
            title="另存 Word 預覽為…",
        )
        if not dest:
            return
        try:
            shutil.copy2(p, dest)
        except OSError as e:
            messagebox.showerror("匯出", f"無法另存：{e}", parent=self.root)
            return
        self._status.set(f"已另存 Word：{dest}")

    def _export_save_txt_from_last_preview(self) -> None:
        tags = self._export_last_preview_tags
        if not tags:
            messagebox.showinfo("匯出", "請先按「產生／更新 Word 預覽」。", parent=self.root)
            return
        layout_key = self._export_last_preview_layout or "screen"
        inc_stats = bool(self._export_last_preview_inc_stats)
        inc_foot = bool(self._export_last_preview_inc_foot)
        text = self._build_primary_filter_export_text(
            tags,
            include_block_stats=inc_stats,
            include_footer=inc_foot,
            layout=layout_key,
        )
        if not text.strip():
            messagebox.showinfo("匯出", "無可輸出的文字。", parent=self.root)
            return
        dest = filedialog.asksaveasfilename(
            parent=self.root,
            defaultextension=".txt",
            filetypes=[("文字檔 (*.txt)", "*.txt"), ("全部", "*.*")],
            title="另存為純文字…",
        )
        if not dest:
            return
        try:
            with open(dest, "w", encoding="utf-8-sig", newline="\n") as f:
                f.write(text)
        except OSError as e:
            messagebox.showerror("匯出", f"無法寫入：{e}", parent=self.root)
            return
        self._status.set(f"已儲存文字檔：{dest}")

    # --- 分頁：交叉表（人數 × 欄標籤）---
    def _crosstab_row_category(self, r) -> str:
        """列標籤：小／大／未標示（與主標籤篩選人數區間一致）。"""
        sz = headcount_size_label(self._row_headcount_str(r))
        return sz if sz in ("小", "大") else "未標示"

    def _open_crosstab_column_picker(self) -> None:
        values = list_hashtags()
        if not values:
            messagebox.showinfo(
                "交叉表",
                "「# 標籤庫」目前為空。\n請至「# 標籤庫」分頁新增標籤並儲存，或在資料中使用 #單詞後執行分析以寫入庫。",
                parent=self.root,
            )
            return

        prev = set(self._crosstab_col_tags) & set(values)
        order_map = {name: i for i, name in enumerate(values)}
        dlg = tk.Toplevel(self.root)
        dlg.title("交叉表：選擇欄標籤")
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.geometry("480x420")
        dlg.minsize(360, 240)

        btnf = ttk.Frame(dlg, padding=8)
        btnf.pack(fill=tk.X, side=tk.BOTTOM)
        mid = ttk.Frame(dlg)
        mid.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(mid, highlightthickness=0)
        sb = ttk.Scrollbar(mid, orient=tk.VERTICAL, command=canvas.yview)
        inner = ttk.Frame(canvas, padding=8)
        win_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _ic(_e: tk.Event | None = None) -> None:
            bb = canvas.bbox("all")
            if bb:
                canvas.configure(scrollregion=bb)

        def _cc(e: tk.Event) -> None:
            canvas.itemconfigure(win_id, width=e.width)

        inner.bind("<Configure>", _ic)
        canvas.bind("<Configure>", _cc)
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        def _picker_mousewheel(event: tk.Event) -> None:
            d = getattr(event, "delta", 0)
            if d:
                canvas.yview_scroll(int(-1 * d / 120), "units")
            elif getattr(event, "num", None) == 4:
                canvas.yview_scroll(-1, "units")
            elif getattr(event, "num", None) == 5:
                canvas.yview_scroll(1, "units")

        def _bind_picker_wheel(w: tk.Widget) -> None:
            w.bind("<MouseWheel>", _picker_mousewheel)
            w.bind("<Button-4>", _picker_mousewheel)
            w.bind("<Button-5>", _picker_mousewheel)
            for c in w.winfo_children():
                _bind_picker_wheel(c)

        vars_by: dict[str, tk.IntVar] = {}
        for i, v in enumerate(values):
            vars_by[v] = tk.IntVar(value=1 if v in prev else 0)
            tk.Checkbutton(inner, text=v, variable=vars_by[v], anchor=tk.W).grid(
                row=i, column=0, sticky=tk.W, pady=2
            )

        _bind_picker_wheel(inner)
        for w, evs in (
            (canvas, ("<MouseWheel>", "<Button-4>", "<Button-5>")),
            (sb, ("<MouseWheel>", "<Button-4>", "<Button-5>")),
            (mid, ("<MouseWheel>", "<Button-4>", "<Button-5>")),
        ):
            for ev in evs:
                w.bind(ev, _picker_mousewheel)

        def _sel_all() -> None:
            for iv in vars_by.values():
                iv.set(1)

        def _clr_all() -> None:
            for iv in vars_by.values():
                iv.set(0)

        def _ok() -> None:
            chosen = [v for v in values if int(vars_by[v].get() or 0) == 1]
            chosen.sort(key=lambda t: order_map.get(t, 10**9))
            self._crosstab_col_tags = chosen
            try:
                save_filter_prefs(
                    self._filter_selected_tags,
                    self._filter_display_rules,
                    self._filter_export_blocks,
                    self._get_export_templates_live(),
                    crosstab_col_tags=chosen,
                )
            except OSError as e:
                messagebox.showwarning("交叉表", f"無法儲存欄標籤至設定檔：{e}", parent=self.root)
            dlg.destroy()
            self._update_crosstab_cols_summary()
            self._nb.select(self._tab_crosstab)
            self._refresh_crosstab_table()
            self._status.set(f"交叉表：已選 {len(chosen)} 個欄標籤（已儲存）。")

        def _cancel() -> None:
            dlg.destroy()

        ttk.Button(btnf, text="全選", command=_sel_all).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btnf, text="全不選", command=_clr_all).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Button(btnf, text="確定", command=_ok).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(btnf, text="取消", command=_cancel).pack(side=tk.RIGHT)
        dlg.bind("<Escape>", lambda _e: _cancel())

    def _compute_crosstab_matrix(
        self,
    ) -> tuple[
        tuple[tuple[str, ...], list[str], list[list[int]], list[int], list[int], int] | None,
        str | None,
    ]:
        """
        與畫面表格相同口徑（隱藏欄／列合計為 0 者）。
        成功回傳 (row_labels, col_tags, mat, row_sums, col_sums, grand), None；
        失敗回傳 None, 錯誤碼。
        """
        row_labels_all = ("小", "大", "未標示")
        col_tags = list(self._crosstab_col_tags)

        if not self._rows:
            return None, "no_rows"
        if not col_tags:
            return None, "no_cols"

        n_rows, n_cols = len(row_labels_all), len(col_tags)
        mat: list[list[int]] = [[0] * n_cols for _ in range(n_rows)]
        for ci, tag in enumerate(col_tags):
            for r in self._rows_matching_tag_value(tag):
                cat = self._crosstab_row_category(r)
                ri = row_labels_all.index(cat)
                mat[ri][ci] += 1

        row_sums = [sum(mat[ri][j] for j in range(n_cols)) for ri in range(n_rows)]
        col_sums = [sum(mat[ri][ci] for ri in range(n_rows)) for ci in range(n_cols)]

        keep_col_idx = [j for j in range(n_cols) if col_sums[j] > 0]
        keep_row_idx = [i for i in range(n_rows) if row_sums[i] > 0]

        if not keep_col_idx:
            return None, "all_zero_cols"

        col_tags_f = [col_tags[j] for j in keep_col_idx]
        row_labels = tuple(row_labels_all[i] for i in keep_row_idx)
        mat_f = [[mat[i][j] for j in keep_col_idx] for i in keep_row_idx]
        nr, nc = len(row_labels), len(col_tags_f)
        row_sums_f = [sum(mat_f[ri][j] for j in range(nc)) for ri in range(nr)]
        col_sums_f = [sum(mat_f[ri][ci] for ri in range(nr)) for ci in range(nc)]
        grand = sum(row_sums_f)
        return (row_labels, col_tags_f, mat_f, row_sums_f, col_sums_f, grand), None

    def _export_crosstab_spreadsheet(self) -> None:
        """匯出為 CSV（UTF-8 BOM），可用 Excel／試算表開啟。"""
        data, err = self._compute_crosstab_matrix()
        msgs = {
            "no_rows": "尚無分析資料。請至「輸入與分析」執行分析後再匯出。",
            "no_cols": "尚未選擇欄標籤。請按「選擇欄標籤…」勾選後再按「更新表格」或匯出。",
            "all_zero_cols": "所選欄標籤目前皆為 0 筆，無表可匯出。",
        }
        if err:
            messagebox.showinfo("交叉表", msgs.get(err, "無法匯出。"), parent=self.root)
            return

        row_labels, col_tags, mat, row_sums, col_sums, grand = data
        dest = filedialog.asksaveasfilename(
            parent=self.root,
            defaultextension=".csv",
            filetypes=[("CSV 試算表 (*.csv)", "*.csv"), ("全部", "*.*")],
            title="匯出交叉表…",
        )
        if not dest:
            return
        try:
            with open(dest, "w", encoding="utf-8-sig", newline="") as f:
                w = csv.writer(f)
                w.writerow(["人數＼標籤", *col_tags, "列合計"])
                for i, rl in enumerate(row_labels):
                    w.writerow([rl, *[str(mat[i][j]) for j in range(len(col_tags))], str(row_sums[i])])
                w.writerow(["欄合計", *[str(col_sums[j]) for j in range(len(col_tags))], str(grand)])
        except OSError as e:
            messagebox.showerror("交叉表", f"無法寫入檔案：{e}", parent=self.root)
            return
        self._status.set(f"交叉表已匯出：{dest}")

    def _update_crosstab_cols_summary(self) -> None:
        if not getattr(self, "_crosstab_cols_summary", None):
            return
        tags = self._crosstab_col_tags
        if not tags:
            self._crosstab_cols_summary.set("目前欄標籤：（尚未選擇）")
            return
        preview = "、".join(tags[:10])
        if len(tags) > 10:
            preview += f"…（共 {len(tags)} 個）"
        self._crosstab_cols_summary.set(f"目前欄標籤（{len(tags)} 個）：{preview}")

    def _refresh_crosstab_table(self) -> None:
        host = getattr(self, "_crosstab_grid_host", None)
        if host is None:
            return
        for w in host.winfo_children():
            w.destroy()

        data, err = self._compute_crosstab_matrix()
        if err == "no_rows":
            ttk.Label(
                host,
                text="尚無分析資料。請至「輸入與分析」貼上資料並執行分析後，再按「更新表格」。",
            ).pack(anchor=tk.W)
            return
        if err == "no_cols":
            ttk.Label(
                host,
                text="尚未選擇欄標籤。請按「選擇欄標籤…」從「# 標籤庫」勾選（例如：台南、高雄），再按「更新表格」。",
            ).pack(anchor=tk.W)
            return
        if err == "all_zero_cols":
            ttk.Label(
                host,
                text=(
                    "所勾選的欄標籤在目前分析結果中，欄合計皆為 0，故不顯示表格欄位。"
                    "請確認標籤與解析結果是否一致，或至「選擇欄標籤…」調整勾選後再按「更新表格」。"
                ),
                wraplength=780,
            ).pack(anchor=tk.W)
            self._status.set("交叉表：所選欄標籤皆為 0 筆。")
            return
        assert data is not None
        row_labels, col_tags, mat, row_sums, col_sums, grand = data
        n_rows, n_cols = len(row_labels), len(col_tags)

        hdr_bg = "#E3F2FD"
        cell_bg = "#FAFAFA"
        sum_bg = "#FFF8E1"
        font_hdr = ("Microsoft JhengHei UI", 10, "bold")
        font_cell = ("Microsoft JhengHei UI", 10)

        gridf = tk.Frame(host, bg="#B0BEC5", padx=1, pady=1)
        gridf.pack(anchor=tk.NW)

        def add_lbl(r: int, c: int, text: str, *, hdr: bool = False, sum_cell: bool = False) -> None:
            bg = hdr_bg if hdr else (sum_bg if sum_cell else cell_bg)
            bold = hdr or sum_cell
            tk.Label(
                gridf,
                text=text,
                font=font_hdr if bold else font_cell,
                bg=bg,
                fg="#0D47A1" if bold else "#212121",
                padx=12,
                pady=8,
                relief=tk.FLAT,
                borderwidth=1,
                highlightthickness=1,
                highlightbackground="#CFD8DC",
            ).grid(row=r, column=c, sticky=tk.NSEW, padx=1, pady=1)

        add_lbl(0, 0, "人數 ＼ 標籤", hdr=True)
        for j, tname in enumerate(col_tags):
            add_lbl(0, j + 1, tname, hdr=True)
        add_lbl(0, n_cols + 1, "列合計", hdr=True)

        for i, rl in enumerate(row_labels):
            add_lbl(i + 1, 0, rl, hdr=True)
            for j in range(n_cols):
                add_lbl(i + 1, j + 1, str(mat[i][j]))
            add_lbl(i + 1, n_cols + 1, str(row_sums[i]), sum_cell=True)

        add_lbl(n_rows + 1, 0, "欄合計", hdr=True)
        for j in range(n_cols):
            add_lbl(n_rows + 1, j + 1, str(col_sums[j]), sum_cell=True)
        add_lbl(n_rows + 1, n_cols + 1, str(grand), sum_cell=True)

        for c in range(n_cols + 2):
            gridf.grid_columnconfigure(c, weight=1)
        self._status.set("交叉表：已更新。")

    def _build_tab_crosstab(self) -> None:
        tab = ttk.Frame(self._nb, padding=8)
        self._tab_crosstab = tab
        self._nb.add(tab, text="交叉表")

        ttk.Label(
            tab,
            text=(
                "以「人數區間」為列（2～3人→小、3～4人→大；無法判斷→未標示）、以「# 標籤庫」詞彙為欄，"
                "統計每格筆數並加總。「選擇欄標籤…」勾選後會寫入設定檔，下次開啟自動還原。"
            ),
            wraplength=820,
        ).pack(anchor=tk.W, pady=(0, 6))

        bar = ttk.Frame(tab)
        bar.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(bar, text="選擇欄標籤…", command=self._open_crosstab_column_picker).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Button(bar, text="更新表格", command=self._refresh_crosstab_table).pack(side=tk.LEFT)
        ttk.Button(bar, text="匯出試算表…", command=self._export_crosstab_spreadsheet).pack(
            side=tk.LEFT, padx=(12, 0)
        )
        self._crosstab_cols_summary = tk.StringVar(value="目前欄標籤：（尚未選擇）")
        ttk.Label(bar, textvariable=self._crosstab_cols_summary).pack(side=tk.LEFT, padx=(16, 0))

        self._crosstab_grid_host = ttk.Frame(tab)
        self._crosstab_grid_host.pack(fill=tk.BOTH, expand=True, anchor=tk.NW, pady=(8, 0))

        ttk.Label(
            tab,
            text=(
                "說明：每格為「該筆資料含此欄標籤」且「人數區間符合該列」的筆數，與「主標籤篩選」各區塊的 (N 筆) 口徑相同。"
                "同一筆若同時含多個欄標籤，會各欄各計一次；列／欄合計為各格數字相加。"
                "欄合計為 0 的標籤、列合計為 0 的人數列不會顯示；若所選欄標籤全部為 0 筆則僅顯示提示不畫表。"
                "「匯出試算表…」輸出與目前表相同的 CSV（UTF-8 BOM），可用 Excel 開啟。"
            ),
            wraplength=820,
        ).pack(anchor=tk.W, pady=(12, 0))

        self._update_crosstab_cols_summary()

    # --- 分頁：匯出 ---
    def _build_tab_export(self) -> None:
        tab = ttk.Frame(self._nb, padding=8)
        self._tab_export = tab
        self._nb.add(tab, text="匯出")

        self._export_preview_docx_path = None
        self._export_last_preview_tags = None
        self._export_last_preview_layout = None
        self._export_last_preview_inc_stats = None
        self._export_last_preview_inc_foot = None

        ttk.Label(
            tab,
            text="「主標籤篩選」可勾「輸出」；「匯出…」或下方按鈕會直接產生 Word 暫存檔作為預覽（不經本頁文字編輯）。"
            "請用「用預設程式開啟」在 Word 中檢視。「與篩選區塊相同」為表格排版。需 pip install python-docx。下方可匯出 JSON／CSV。",
            wraplength=820,
        ).pack(anchor=tk.W, pady=(0, 8))

        opt = ttk.Frame(tab)
        opt.pack(fill=tk.X, pady=(0, 6))
        self._var_export_block_stats = tk.IntVar(value=1)
        self._var_export_footer = tk.IntVar(value=1)
        ttk.Checkbutton(
            opt,
            text="預覽含各區統計列",
            variable=self._var_export_block_stats,
            command=self._on_export_tab_option_changed,
        ).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Checkbutton(
            opt,
            text="預覽含合計表（與主標籤篩選右上表相同；儲格格式如 35+3(拋)）",
            variable=self._var_export_footer,
            command=self._on_export_tab_option_changed,
        ).pack(side=tk.LEFT)

        opt_layout = ttk.Frame(tab)
        opt_layout.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(opt_layout, text="排列方式：").pack(side=tk.LEFT, padx=(0, 6))
        self._export_layout_label_to_key = {b: a for a, b in _EXPORT_LAYOUT_OPTIONS}
        _layout_labels = [b for _, b in _EXPORT_LAYOUT_OPTIONS]
        self._var_export_layout_ui = tk.StringVar(value=_layout_labels[0])
        self._cbb_export_layout = ttk.Combobox(
            opt_layout,
            textvariable=self._var_export_layout_ui,
            values=_layout_labels,
            state="readonly",
            width=44,
        )
        self._cbb_export_layout.pack(side=tk.LEFT)
        self._cbb_export_layout.bind(
            "<<ComboboxSelected>>",
            lambda _e: self._refresh_export_tab_preview(None),
        )
        ttk.Label(
            opt_layout,
            text="（建議列印用「Tab 分欄」或「直印對齊」；橫印用三／四欄）",
            foreground="#555555",
        ).pack(side=tk.LEFT, padx=(12, 0))

        custom_row = ttk.Frame(tab)
        custom_row.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(
            custom_row,
            text="編輯自訂格式…",
            command=self._open_export_custom_format_dialog,
        ).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(
            custom_row,
            text="排列選「自訂」時使用；另開視窗編輯，不佔本頁空間。",
            foreground="#555555",
        ).pack(side=tk.LEFT)

        btn_row = ttk.Frame(tab)
        btn_row.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(
            btn_row,
            text="產生／更新 Word 預覽",
            command=lambda: self._refresh_export_tab_preview(None),
        ).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(
            btn_row,
            text="用預設程式開啟預覽",
            command=self._export_open_docx_preview,
        ).pack(side=tk.LEFT)

        docx_frame = ttk.LabelFrame(tab, text="Word 預覽（暫存 .docx）", padding=8)
        docx_frame.pack(fill=tk.BOTH, expand=True, pady=(4, 12))

        self._lbl_export_preview_status = tk.Label(
            docx_frame,
            text="",
            justify=tk.LEFT,
            anchor=tk.W,
            wraplength=780,
        )
        self._lbl_export_preview_status.pack(fill=tk.X, pady=(0, 4))
        self._lbl_export_preview_path = tk.Label(
            docx_frame,
            text="",
            justify=tk.LEFT,
            anchor=tk.W,
            wraplength=780,
            font=("Consolas", 9),
            fg="#555555",
        )
        self._lbl_export_preview_path.pack(fill=tk.X, pady=(0, 8))

        docx_btns = ttk.Frame(docx_frame)
        docx_btns.pack(fill=tk.X)
        b_open = ttk.Button(docx_btns, text="用預設程式開啟", command=self._export_open_docx_preview)
        b_open.pack(side=tk.LEFT, padx=(0, 8))
        b_rev = ttk.Button(docx_btns, text="在檔案總管顯示", command=self._export_reveal_docx_preview)
        b_rev.pack(side=tk.LEFT, padx=(0, 8))
        b_sdocx = ttk.Button(docx_btns, text="另存 Word…", command=self._export_save_docx_copy_as)
        b_sdocx.pack(side=tk.LEFT, padx=(0, 8))
        b_stxt = ttk.Button(docx_btns, text="另存純文字…", command=self._export_save_txt_from_last_preview)
        b_stxt.pack(side=tk.LEFT)
        self._export_preview_action_buttons = [b_open, b_rev, b_sdocx, b_stxt]
        for b in self._export_preview_action_buttons:
            b.configure(state="disabled")

        self._export_preview_set_idle("尚未產生預覽。請按「產生／更新 Word 預覽」。")

        ttk.Separator(tab, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=(4, 10))
        ttk.Label(
            tab,
            text="完整解析結果（與主標籤篩選預覽無關）：",
        ).pack(anchor=tk.W, pady=(0, 6))
        bf = ttk.Frame(tab)
        bf.pack(anchor=tk.W)
        ttk.Button(bf, text="匯出 JSON…", command=self._export_json).pack(anchor=tk.W, pady=(0, 8))
        ttk.Button(bf, text="匯出 CSV…", command=self._export_csv).pack(anchor=tk.W)

    # --- 分頁：# 標籤庫 ---
    def _build_tab_hashtag_db(self) -> None:
        tab = ttk.Frame(self._nb, padding=8)
        self._nb.add(tab, text="# 標籤庫")

        self._lbl_db_path = tk.StringVar(value=str(database_path()))
        ttk.Label(tab, textvariable=self._lbl_db_path, wraplength=820).pack(anchor=tk.W, pady=(0, 4))
        ttk.Label(
            tab,
            text="清單順序＝「主標籤篩選」各區塊由上而下的順序。請選取一列後用右側按鈕調整；分析時若原文含 #單詞，新標籤會接在清單末尾。",
            wraplength=820,
        ).pack(anchor=tk.W, pady=(0, 8))

        row = ttk.Frame(tab)
        row.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(row, text="儲存至檔案", command=self._save_hashtag_db).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(row, text="從檔案重新載入", command=self._reload_hashtag_db_confirm).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(row, text="用文字編輯…", command=self._open_hashtag_bulk_text).pack(side=tk.LEFT)

        mid = ttk.Frame(tab)
        mid.pack(fill=tk.BOTH, expand=True)

        btn_col = ttk.Frame(mid)
        btn_col.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 8))
        for txt, cmd in (
            ("上移", self._hashtag_move_up),
            ("下移", self._hashtag_move_down),
            ("置頂", self._hashtag_move_top),
            ("置底", self._hashtag_move_bottom),
            ("刪除所選", self._hashtag_delete_selected),
        ):
            ttk.Button(btn_col, text=txt, command=cmd, width=10).pack(fill=tk.X, pady=(0, 4))

        lb_frame = ttk.Frame(mid)
        lb_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._list_hashtag_db = tk.Listbox(
            lb_frame,
            height=16,
            font=("Microsoft JhengHei UI", 10),
            exportselection=False,
        )
        sy_lb = ttk.Scrollbar(lb_frame, orient=tk.VERTICAL, command=self._list_hashtag_db.yview)
        self._list_hashtag_db.configure(yscrollcommand=sy_lb.set)
        self._list_hashtag_db.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sy_lb.pack(side=tk.RIGHT, fill=tk.Y)

        add_row = ttk.Frame(tab)
        add_row.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(add_row, text="新增：").pack(side=tk.LEFT, padx=(0, 4))
        self._var_new_hashtag = tk.StringVar()
        ttk.Entry(add_row, textvariable=self._var_new_hashtag, width=36).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(add_row, text="加入清單", command=self._hashtag_add_from_entry).pack(side=tk.LEFT)

        self._refresh_hashtag_db_view()

    def _hashtag_lb_select(self, index: int) -> None:
        n = self._list_hashtag_db.size()
        if n <= 0:
            return
        index = max(0, min(index, n - 1))
        self._list_hashtag_db.selection_clear(0, tk.END)
        self._list_hashtag_db.selection_set(index)
        self._list_hashtag_db.activate(index)
        self._list_hashtag_db.see(index)

    def _hashtag_lb_get_items(self) -> list[str]:
        return list(self._list_hashtag_db.get(0, tk.END))

    def _hashtag_lb_set_items(self, items: list[str]) -> None:
        self._list_hashtag_db.delete(0, tk.END)
        for t in items:
            self._list_hashtag_db.insert(tk.END, t)

    def _hashtag_persist_lb(self) -> int | None:
        tags = self._hashtag_lb_get_items()
        try:
            n = save_hashtag_list(tags)
        except OSError as e:
            messagebox.showerror("# 標籤庫", f"無法寫入：{e}", parent=self.root)
            return None
        self._after_hashtag_order_changed(n)
        return n

    def _after_hashtag_order_changed(self, n: int) -> None:
        self._lbl_db_path.set(str(database_path()))
        if self._rows and self._filter_selected_tags:
            lib = set(list_hashtags())
            vis = [t for t in self._filter_selected_tags if t in lib]
            if vis:
                self._rebuild_primary_filter_results()
                self._status.set(f"標籤庫已更新並依新順序刷新「主標籤篩選」（共 {n} 個標籤）。")
                return
        self._status.set(f"標籤庫已寫入（{n} 個）：{database_path()}")

    def _hashtag_move_up(self) -> None:
        sel = self._list_hashtag_db.curselection()
        if not sel or sel[0] == 0:
            return
        i = int(sel[0])
        items = self._hashtag_lb_get_items()
        items[i - 1], items[i] = items[i], items[i - 1]
        self._hashtag_lb_set_items(items)
        self._hashtag_lb_select(i - 1)
        self._hashtag_persist_lb()

    def _hashtag_move_down(self) -> None:
        sel = self._list_hashtag_db.curselection()
        items = self._hashtag_lb_get_items()
        if not sel or int(sel[0]) >= len(items) - 1:
            return
        i = int(sel[0])
        items[i], items[i + 1] = items[i + 1], items[i]
        self._hashtag_lb_set_items(items)
        self._hashtag_lb_select(i + 1)
        self._hashtag_persist_lb()

    def _hashtag_move_top(self) -> None:
        sel = self._list_hashtag_db.curselection()
        if not sel or sel[0] == 0:
            return
        i = int(sel[0])
        items = self._hashtag_lb_get_items()
        item = items.pop(i)
        items.insert(0, item)
        self._hashtag_lb_set_items(items)
        self._hashtag_lb_select(0)
        self._hashtag_persist_lb()

    def _hashtag_move_bottom(self) -> None:
        sel = self._list_hashtag_db.curselection()
        items = self._hashtag_lb_get_items()
        if not sel or int(sel[0]) >= len(items) - 1:
            return
        i = int(sel[0])
        item = items.pop(i)
        items.append(item)
        self._hashtag_lb_set_items(items)
        self._hashtag_lb_select(len(items) - 1)
        self._hashtag_persist_lb()

    def _hashtag_delete_selected(self) -> None:
        sel = self._list_hashtag_db.curselection()
        if not sel:
            messagebox.showinfo("# 標籤庫", "請先在清單中選取一列。", parent=self.root)
            return
        i = int(sel[0])
        items = self._hashtag_lb_get_items()
        name = items[i]
        if not messagebox.askyesno(
            "# 標籤庫",
            f"從清單移除「{name}」？",
            parent=self.root,
        ):
            return
        del items[i]
        self._hashtag_lb_set_items(items)
        if items:
            self._hashtag_lb_select(min(i, len(items) - 1))
        self._hashtag_persist_lb()

    def _hashtag_add_from_entry(self) -> None:
        raw = (self._var_new_hashtag.get() or "").strip()
        if not raw:
            messagebox.showinfo("# 標籤庫", "請輸入要加入的標籤文字。", parent=self.root)
            return
        items = self._hashtag_lb_get_items()
        if raw in items:
            messagebox.showwarning("# 標籤庫", "清單中已有相同標籤。", parent=self.root)
            return
        items.append(raw)
        self._hashtag_lb_set_items(items)
        self._var_new_hashtag.set("")
        self._hashtag_lb_select(len(items) - 1)
        self._hashtag_persist_lb()

    def _open_hashtag_bulk_text(self) -> None:
        dlg = tk.Toplevel(self.root)
        dlg.title("用文字編輯標籤庫")
        dlg.transient(self.root)
        dlg.geometry("520x420")
        dlg.minsize(400, 280)

        # 先固定底部按鈕列，再讓文字區 fill 剩餘空間；否則 Text 先 expand 會佔滿寬度，
        # 按鈕列被擠到右側狹縫，看起來像「確定無效／無法儲存」。
        bf = ttk.Frame(dlg, padding=(10, 0, 10, 10))
        bf.pack(side=tk.BOTTOM, fill=tk.X)

        ttk.Label(
            dlg,
            text="每行一個標籤；確定後會取代目前清單順序（與主標籤篩選區塊順序一致）。",
            wraplength=480,
        ).pack(anchor=tk.W, padx=10, pady=(10, 6), fill=tk.X)

        mid = ttk.Frame(dlg)
        mid.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        txt = tk.Text(mid, height=16, wrap=tk.WORD, font=("Consolas", 10))
        sy = ttk.Scrollbar(mid, orient=tk.VERTICAL, command=txt.yview)
        txt.configure(yscrollcommand=sy.set)
        txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sy.pack(side=tk.RIGHT, fill=tk.Y)

        txt.insert("1.0", "\n".join(self._hashtag_lb_get_items()))

        def on_ok() -> None:
            raw = txt.get("1.0", tk.END).rstrip("\r\n")
            try:
                n = replace_hashtags_from_text(raw)
            except OSError as e:
                messagebox.showerror("# 標籤庫", f"無法寫入：{e}", parent=dlg)
                return
            self._refresh_hashtag_db_view()
            self._after_hashtag_order_changed(n)
            dlg.destroy()

        def on_cancel() -> None:
            dlg.destroy()

        ttk.Button(bf, text="確定", command=on_ok).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(bf, text="取消", command=on_cancel).pack(side=tk.RIGHT)
        dlg.bind("<Escape>", lambda _e: on_cancel())
        dlg.grab_set()

    def _refresh_hashtag_db_view(self) -> None:
        self._lbl_db_path.set(str(database_path()))
        tags = list_hashtags()
        self._hashtag_lb_set_items(tags)

    def _save_hashtag_db(self) -> None:
        n = self._hashtag_persist_lb()
        if n is not None:
            self._status.set(f"已儲存標籤庫，共 {n} 個標籤：{database_path()}")

    def _reload_hashtag_db_confirm(self) -> None:
        if not messagebox.askyesno(
            "# 標籤庫",
            "從檔案重新載入？清單上未寫入的變更將遺失。",
            parent=self.root,
        ):
            return
        self._refresh_hashtag_db_view()
        n = len(list_hashtags())
        self._after_hashtag_order_changed(n)
        self._status.set(f"已從檔案載入標籤庫（{n} 個）：{database_path()}")

    def _clear_input(self) -> None:
        self.txt_in.delete("1.0", tk.END)
        p = self._page_by_id(self._pages_state.get("current_page_id") or "")
        if p is not None:
            p["text"] = ""
        try:
            save_input_pages_state(self._pages_state)
        except OSError:
            pass

    def _analyze(self) -> None:
        self._sync_txt_in_to_current_page()
        self._pages_state["roster_view"] = self._roster_view_var.get()
        try:
            save_input_pages_state(self._pages_state)
        except OSError as e:
            messagebox.showwarning("儲存", f"無法寫入輸入頁設定：{e}", parent=self.root)

        all_rows: list = []
        pages = self._pages_state.get("pages") or []
        page_names = [str(p.get("name") or "未命名").strip() or "未命名" for p in pages]
        try:
            register_hashtags(page_names)
        except OSError as e:
            messagebox.showerror("標籤資料庫", f"無法寫入標籤庫（頁名）：{e}", parent=self.root)

        for p in pages:
            raw = (p.get("text") or "").strip()
            if not raw:
                continue
            pname = str(p.get("name") or "未命名").strip() or "未命名"
            chunk = parse_bulk(p.get("text") or "")
            for r in chunk:
                r.source_page = pname
                r.tags.append({"category": "manual", "value": pname})
            all_rows.extend(chunk)

        self._rows = all_rows
        self._clear_primary_filter_panels()
        self._sync_roster_page_choices(set_combo=True)
        self._refresh_roster_tree()

        hashtag_values = [
            t["value"] for r in self._rows for t in r.tags if t.get("category") == "hashtag"
        ]
        try:
            n_new, n_total = register_hashtags(hashtag_values)
        except OSError as e:
            messagebox.showerror("標籤資料庫", f"無法寫入標籤庫：{e}", parent=self.root)
            n_new, n_total = 0, 0

        db_hint = f"# 標籤庫新增 {n_new} 個，庫內共 {n_total} 個。" if (hashtag_values or n_total) else ""
        self._refresh_hashtag_db_view()

        if self._rows and self._filter_selected_tags:
            self._rebuild_primary_filter_results()
            self._nb.select(self._tab_primary_filter)
            extra = "已依儲存的主標籤勾選顯示篩選。"
            self._status.set(f"已解析 {len(self._rows)} 筆。{extra}{db_hint}")
        else:
            if not self._rows:
                tk.Label(
                    self._filter_inner,
                    text="請先在上方貼上資料並分析；若有儲存主標籤勾選，分析成功後會自動顯示篩選結果。",
                    bg=FILTER_INNER_BG,
                    fg="#333333",
                    justify=tk.LEFT,
                    wraplength=780,
                ).pack(anchor=tk.W, pady=8)
            else:
                tk.Label(
                    self._filter_inner,
                    text="分析完成。請至本分頁按「選擇主標籤…」勾選要檢視的標籤（勾選後會自動儲存，下次不必重選）。",
                    bg=FILTER_INNER_BG,
                    fg="#333333",
                    justify=tk.LEFT,
                    wraplength=780,
                ).pack(anchor=tk.W, pady=8)
            self._filter_apply_scroll()
            self._nb.select(self._tab_roster)
            self._status.set(f"已解析 {len(self._rows)} 筆。{db_hint}")

        if self._rows and self._crosstab_col_tags:
            self._refresh_crosstab_table()

    def _selected_row_index(self) -> int | None:
        sel = self.tree.selection()
        if not sel:
            return None
        iid = sel[0]
        idx = self._row_by_iid.get(iid)
        if idx is None or idx >= len(self._rows):
            return None
        return idx

    def _add_manual_tag(self) -> None:
        if not self._rows:
            messagebox.showinfo("提示", "請先在「輸入與分析」分頁執行分析。")
            return
        idx = self._selected_row_index()
        if idx is None:
            messagebox.showwarning("提示", "請在下方表格選取一列。")
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("新增標籤")
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text="分類（category）：").grid(row=0, column=0, sticky=tk.W, pady=(0, 4))
        var_cat = tk.StringVar(value="manual")
        ent_cat = ttk.Entry(frm, textvariable=var_cat, width=28)
        ent_cat.grid(row=1, column=0, columnspan=2, sticky=tk.EW, pady=(0, 8))

        ttk.Label(frm, text="內容（value）：").grid(row=2, column=0, sticky=tk.W, pady=(0, 4))
        var_val = tk.StringVar()
        ent_val = ttk.Entry(frm, textvariable=var_val, width=28)
        ent_val.grid(row=3, column=0, columnspan=2, sticky=tk.EW, pady=(0, 12))
        frm.columnconfigure(0, weight=1)

        def on_ok() -> None:
            cat = var_cat.get().strip() or "manual"
            val = var_val.get().strip()
            if not val:
                messagebox.showwarning("提示", "請輸入標籤內容。", parent=dlg)
                return
            val_out = normalize_leading_no(val)
            r = self._rows[idx]
            r.tags.append({"category": cat, "value": val_out})
            sel_iids = self.tree.selection()
            tree_iid = sel_iids[0] if sel_iids else None
            if tree_iid and self.tree.exists(tree_iid):
                self.tree.set(tree_iid, "tag_count", len(r.tags))
            self._on_select_row()
            self._status.set(f"已於第 {r.line_no} 列新增標籤：{cat} / {val_out}")
            dlg.destroy()

        def on_cancel() -> None:
            dlg.destroy()

        btn_frm = ttk.Frame(frm)
        btn_frm.grid(row=4, column=0, columnspan=2, sticky=tk.E)
        ttk.Button(btn_frm, text="確定", command=on_ok).pack(side=tk.RIGHT, padx=(4, 0))
        ttk.Button(btn_frm, text="取消", command=on_cancel).pack(side=tk.RIGHT)

        ent_val.focus_set()
        dlg.bind("<Return>", lambda _e: on_ok())
        dlg.bind("<Escape>", lambda _e: on_cancel())

    def _on_select_row(self, _evt=None) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        idx = self._row_by_iid.get(iid)
        if idx is None or idx >= len(self._rows):
            return
        r = self._rows[idx]
        self.txt_detail.delete("1.0", tk.END)
        payload = {
            "line_no": r.line_no,
            "serial": r.serial,
            "customer_name": r.customer_name,
            "headcount": r.headcount,
            "source_page": getattr(r, "source_page", None),
            "paren_tags": r.paren_tags,
            "tags": r.tags,
            "errors": r.errors,
            "raw_line": r.raw_line,
        }
        self.txt_detail.insert(tk.END, json.dumps(payload, ensure_ascii=False, indent=2))

    def _export_json(self) -> None:
        if not self._rows:
            messagebox.showinfo("提示", "請先在「輸入與分析」分頁執行分析。")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON", "*.json"), ("全部", "*.*")],
        )
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            f.write(rows_to_json(self._rows))
        self._status.set(f"已匯出 JSON：{path}")

    def _export_csv(self) -> None:
        if not self._rows:
            messagebox.showinfo("提示", "請先在「輸入與分析」分頁執行分析。")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv"), ("全部", "*.*")],
        )
        if not path:
            return
        text = rows_to_csv_text(self._rows)
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            f.write(text)
        self._status.set(f"已匯出 CSV：{path}")

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    OrderNoteApp().run()


if __name__ == "__main__":
    main()
