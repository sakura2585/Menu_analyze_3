# -*- coding: utf-8 -*-
"""
訂餐備註分析小工具：分頁式桌面介面。
依賴：標準庫（tkinter）；匯出 Word 時需安裝 python-docx（pip install python-docx）。
執行：python app.py
"""

from __future__ import annotations

from collections import Counter
import csv
from datetime import date, timedelta
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import tkinter as tk
import urllib.error
import urllib.request
import webbrowser
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
from web_fetch_flow import WebFetchFlow, WebFetchRequest
from web_fetch_profiles import little_champion_profile
from web_fetch_settings_store import WebFetchSettings, load_web_fetch_settings, save_web_fetch_settings
from app_paths import project_data_dir

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
# 自備餐具：須完整詞「自備餐具」（與解析器 UTENSIL_SNIPPETS 一致）
_UTENSIL_MARKER = "自備餐具"
# 主篩選畫布：自備餐具姓名外框色（拋棄式為 #0D47A1）
FILTER_UTENSIL_OUTLINE = "#2E7D32"
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
# 交叉表：分量列（互斥）。小／大來自人數區間（2～3→小、3～4→大）；拋／自依 _fenji_stat_bucket。
_CROSSTAB_PARTITION_ROWS: tuple[str, ...] = (
    "小",
    "小拋",
    "小自",
    "大",
    "大拋",
    "大自",
    "未標",
    "未標拋",
    "未標自",
)

_EXPORT_LAYOUT_OPTIONS: tuple[tuple[str, str], ...] = (
    ("screen", "與篩選區塊相同（標題・流式欄距・統計）"),
    ("tsv", "Tab 分欄（表頭，試算表／直印）"),
    ("print_cols", "直印對齊（等寬欄・空格）"),
    ("flow3", "橫向三欄（省紙，│ 分隔）"),
    ("flow4", "橫向四欄（省紙，│ 分隔）"),
    ("names", "僅姓名（每筆一行）"),
    ("custom", "自訂（下方格式字串）"),
)

_PF_NAME_SORT_OPTIONS: tuple[tuple[str, str], ...] = (
    ("source", "原始順序"),
    ("serial", "序號"),
    ("name", "姓名"),
    ("headcount", "人數"),
)

_APP_VERSION = "v1.0.6"
_UPDATE_REPO = "sakura2585/Menu_analyze_3"

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
            self._primary_tag_order,
            self._crosstab_tag_order,
            self._export_tag_order,
        ) = load_filter_prefs()
        self._filter_export_vars: dict[str, tk.IntVar] = {}
        # 交叉表欄標籤由 filter_prefs（primary_filter_selection.json）載入／儲存
        self._pages_state: dict = load_input_pages_state()
        self._update_result_var = tk.StringVar(value="尚未檢查更新。")
        self._pf_name_sort_key_var = tk.StringVar(
            value=str(self._pages_state.get("pf_name_sort_key") or "source")
        )
        self._pf_name_sort_dir_var = tk.StringVar(
            value=str(self._pages_state.get("pf_name_sort_dir") or "asc")
        )
        if self._pf_name_sort_key_var.get() not in {k for k, _ in _PF_NAME_SORT_OPTIONS}:
            self._pf_name_sort_key_var.set("source")
        if self._pf_name_sort_dir_var.get() not in {"asc", "desc"}:
            self._pf_name_sort_dir_var.set("asc")
        mw = int(self._pages_state.get("main_ui_width") or 1200)
        mh = int(self._pages_state.get("main_ui_height") or 760)
        self.root.geometry(f"{max(900, mw)}x{max(640, mh)}")
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
        self._build_tab_help()

        self._status = tk.StringVar(value="就緒")
        ttk.Label(outer, textvariable=self._status).pack(anchor=tk.W, pady=(8, 0))

        self._refresh_page_listbox()
        self._apply_current_page_to_txt_in()
        self._sync_roster_page_choices(set_combo=True)
        if self._rows and self._crosstab_col_tags:
            self._refresh_crosstab_table()
        self.root.protocol("WM_DELETE_WINDOW", self._on_app_close)
        # 啟動後延遲檢查更新，避免干擾首屏操作。
        self.root.after(1600, lambda: self._check_updates_from_github(silent=True, show_latest_dialog=False))

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
            self.root.update_idletasks()
            self._pages_state["main_ui_width"] = int(self.root.winfo_width())
            self._pages_state["main_ui_height"] = int(self.root.winfo_height())
        except Exception:
            pass
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
                primary_tag_order=self._primary_tag_order,
                crosstab_tag_order=self._crosstab_tag_order,
                export_tag_order=self._export_tag_order,
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
        ttk.Button(btn_row, text="抓取網頁…", command=self._open_web_fetch_dialog).pack(side=tk.LEFT, padx=(8, 0))

    def _open_web_fetch_dialog(self) -> None:
        profile = little_champion_profile()
        saved = load_web_fetch_settings()
        cur_page = self._page_by_id(self._pages_state.get("current_page_id") or "")
        page_name = (cur_page or {}).get("name") or "目前資料頁"
        page_url = str((cur_page or {}).get("web_fetch_url") or "").strip()

        dlg = tk.Toplevel(self.root)
        dlg.title(f"抓取網頁｜{page_name}")
        dlg.transient(self.root)
        dlg.grab_set()
        ui_w = max(760, int(getattr(saved, "ui_width", 760) or 760))
        ui_h = max(560, int(getattr(saved, "ui_height", 560) or 560))
        dlg.geometry(f"{ui_w}x{ui_h}")
        dlg.minsize(700, 520)

        body = ttk.Frame(dlg, padding=8)
        body.pack(fill=tk.BOTH, expand=True)

        d1 = date.today() + timedelta(days=1)
        page_manual_date = str((cur_page or {}).get("web_fetch_manual_date") or "").strip()
        default_manual_date = page_manual_date or f"{d1.year}年{d1.month}月{d1.day}日"
        init_url = page_url or saved.base_url or profile.base_url
        vars_s: dict[str, tk.StringVar] = {
            "base_url": tk.StringVar(value=init_url),
            "login_account": tk.StringVar(value=saved.login_account or "a0824"),
            "login_password": tk.StringVar(value=saved.login_password or ""),
            "manual_date": tk.StringVar(value=default_manual_date),
            "source_xpath": tk.StringVar(value=saved.source_xpath or ""),
            "date_xpath": tk.StringVar(value=saved.date_xpath or ""),
            "pre_click_xpath": tk.StringVar(value=saved.pre_click_xpath or ""),
            "date_prev_xpath": tk.StringVar(value=saved.date_prev_xpath or ""),
            "date_next_xpath": tk.StringVar(value=saved.date_next_xpath or ""),
        }
        omit_var = tk.IntVar(value=1 if saved.omit_notes_column else 0)
        basic = ttk.LabelFrame(body, text="基本設定", padding=8)
        basic.pack(fill=tk.X)
        adv = ttk.LabelFrame(body, text="進階 XPath（留空=採用內建預設）", padding=8)
        adv.pack(fill=tk.X, pady=(8, 0))

        def _add_row(parent: ttk.LabelFrame, row: int, label: str, key: str, *, show: str | None = None) -> None:
            ttk.Label(parent, text=label).grid(row=row, column=0, sticky=tk.W, pady=3)
            ent = ttk.Entry(parent, textvariable=vars_s[key], show=show)
            ent.grid(row=row, column=1, sticky="ew", padx=(8, 0), pady=3)

        _add_row(basic, 0, "網址", "base_url")
        _add_row(basic, 1, "登入帳號", "login_account")
        _add_row(basic, 2, "登入密碼（可空）", "login_password", show="*")
        _add_row(basic, 3, "指定日期（可空，例：2026年4月7日）", "manual_date")
        ttk.Checkbutton(
            basic,
            text="四欄以上只取前三欄（忽略網站備註欄）",
            variable=omit_var,
        ).grid(row=4, column=1, sticky=tk.W, pady=(4, 0))
        basic.columnconfigure(1, weight=1)

        _add_row(adv, 0, "資料 XPath", "source_xpath")
        _add_row(adv, 1, "日期 XPath", "date_xpath")
        _add_row(adv, 2, "前置按鈕 XPath", "pre_click_xpath")
        _add_row(adv, 3, "日期上一天 XPath", "date_prev_xpath")
        _add_row(adv, 4, "日期下一天 XPath", "date_next_xpath")
        adv.columnconfigure(1, weight=1)

        ttk.Label(
            body,
            text=f"抓取完成後會直接覆蓋「{page_name}」頁文字；網址可各頁分開記憶。",
            foreground="#555555",
        ).pack(anchor=tk.W, pady=(8, 0))

        # 日期快捷列（放在視窗下方上方區塊）：左右箭頭快速切換指定日期
        date_bar = tk.Frame(
            body,
            bg="#E3F2FD",
            highlightbackground="#90CAF9",
            highlightthickness=1,
            padx=8,
            pady=8,
        )
        date_bar.pack(fill=tk.X, pady=(8, 0))
        tk.Label(
            date_bar,
            text="指定日期快捷切換",
            bg="#E3F2FD",
            fg="#0D47A1",
            font=("Microsoft JhengHei UI", 10, "bold"),
        ).pack(side=tk.LEFT, padx=(2, 12))

        def _parse_manual_date(value: str) -> date | None:
            s = (value or "").strip()
            if not s:
                return None
            import re

            m = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日", s)
            if not m:
                return None
            try:
                return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
            except ValueError:
                return None

        def _fmt_zh_date(d: date) -> str:
            return f"{d.year}年{d.month}月{d.day}日"

        def _shift_manual_date(days: int) -> None:
            base = _parse_manual_date(vars_s["manual_date"].get()) or date.today()
            vars_s["manual_date"].set(_fmt_zh_date(base + timedelta(days=days)))

        tk.Button(
            date_bar,
            text="◀ 前一天",
            command=lambda: _shift_manual_date(-1),
            bg="#1976D2",
            fg="#FFFFFF",
            activebackground="#1565C0",
            activeforeground="#FFFFFF",
            font=("Microsoft JhengHei UI", 11, "bold"),
            padx=14,
            pady=6,
            relief=tk.FLAT,
        ).pack(side=tk.LEFT, padx=(0, 10))

        tk.Label(
            date_bar,
            textvariable=vars_s["manual_date"],
            bg="#E3F2FD",
            fg="#0D47A1",
            font=("Microsoft JhengHei UI", 12, "bold"),
            width=16,
            anchor="center",
        ).pack(side=tk.LEFT, padx=(0, 10))

        tk.Button(
            date_bar,
            text="後一天 ▶",
            command=lambda: _shift_manual_date(1),
            bg="#1976D2",
            fg="#FFFFFF",
            activebackground="#1565C0",
            activeforeground="#FFFFFF",
            font=("Microsoft JhengHei UI", 11, "bold"),
            padx=14,
            pady=6,
            relief=tk.FLAT,
        ).pack(side=tk.LEFT)

        btnf = ttk.Frame(dlg, padding=(10, 0, 10, 10))
        btnf.pack(fill=tk.X)

        size_state = {"w": ui_w, "h": ui_h}

        def _on_dlg_configure(event: tk.Event) -> None:
            if event.widget is dlg and event.width > 100 and event.height > 100:
                size_state["w"] = event.width
                size_state["h"] = event.height

        dlg.bind("<Configure>", _on_dlg_configure)

        def _persist_page_url(url: str) -> None:
            p = self._page_by_id(self._pages_state.get("current_page_id") or "")
            if p is not None:
                p["web_fetch_url"] = url
                try:
                    save_input_pages_state(self._pages_state)
                except OSError:
                    pass

        def _persist_page_manual_date(s: str) -> None:
            p = self._page_by_id(self._pages_state.get("current_page_id") or "")
            if p is not None:
                p["web_fetch_manual_date"] = s
                try:
                    save_input_pages_state(self._pages_state)
                except OSError:
                    pass

        def _collect_settings() -> WebFetchSettings:
            base_url = vars_s["base_url"].get().strip()
            _persist_page_url(base_url)
            _persist_page_manual_date(vars_s["manual_date"].get().strip())
            return WebFetchSettings(
                profile_id=profile.profile_id,
                base_url=base_url,
                login_account=vars_s["login_account"].get().strip(),
                login_password=vars_s["login_password"].get(),
                source_xpath=vars_s["source_xpath"].get().strip(),
                date_xpath=vars_s["date_xpath"].get().strip(),
                date_prev_xpath=vars_s["date_prev_xpath"].get().strip(),
                date_next_xpath=vars_s["date_next_xpath"].get().strip(),
                pre_click_xpath=vars_s["pre_click_xpath"].get().strip(),
                omit_notes_column=bool(omit_var.get()),
                ui_width=int(size_state["w"]),
                ui_height=int(size_state["h"]),
            )

        def _close_dialog() -> None:
            try:
                save_web_fetch_settings(_collect_settings())
            except OSError:
                pass
            dlg.destroy()

        def _do_fetch() -> None:
            settings = _collect_settings()
            if not settings.base_url:
                messagebox.showwarning("抓取網頁", "請先填入網址。", parent=dlg)
                return
            try:
                save_web_fetch_settings(settings)
            except OSError:
                pass

            req = WebFetchRequest(
                url=settings.base_url,
                source_xpath=settings.source_xpath,
                date_xpath=settings.date_xpath,
                pre_click_xpath=settings.pre_click_xpath,
                manual_date=vars_s["manual_date"].get().strip(),
                login_account=settings.login_account,
                login_password=settings.login_password,
                profile=profile,
                date_prev_xpath=settings.date_prev_xpath,
                date_next_xpath=settings.date_next_xpath,
                omit_notes_column=settings.omit_notes_column,
            )
            btn_fetch.configure(state=tk.DISABLED)
            btn_cancel.configure(state=tk.DISABLED)
            self.root.config(cursor="watch")
            self.root.update_idletasks()
            try:
                flow = WebFetchFlow(req, status_cb=lambda msg: self._status.set(msg))
                res = flow.run()
            finally:
                self.root.config(cursor="")
                btn_fetch.configure(state=tk.NORMAL)
                btn_cancel.configure(state=tk.NORMAL)
                self.root.update_idletasks()

            if not res.ok:
                messagebox.showerror("抓取網頁", f"抓取失敗：{res.error}", parent=dlg)
                self._status.set("網路抓取：失敗")
                return

            self.txt_in.delete("1.0", tk.END)
            self.txt_in.insert("1.0", res.text)
            self._sync_txt_in_to_current_page()
            self._status.set(f"網路抓取完成：{res.row_count} 行")
            messagebox.showinfo("抓取網頁", f"抓取完成，已填入目前資料頁。\n行數：{res.row_count}", parent=dlg)
            dlg.destroy()

        btn_cancel = ttk.Button(btnf, text="取消", command=_close_dialog)
        btn_cancel.pack(side=tk.RIGHT)
        btn_fetch = ttk.Button(btnf, text="開始抓取", command=_do_fetch)
        btn_fetch.pack(side=tk.RIGHT, padx=(0, 8))
        dlg.protocol("WM_DELETE_WINDOW", _close_dialog)

    def _pdf_default_date_stamp(self) -> str:
        p = self._page_by_id(self._pages_state.get("current_page_id") or "")
        raw = str((p or {}).get("web_fetch_manual_date") or "").strip()
        if not raw:
            d = date.today() + timedelta(days=1)
            return f"{d.year:04d}{d.month:02d}{d.day:02d}"
        m = re.search(r"(\\d{4})\\D*(\\d{1,2})\\D*(\\d{1,2})", raw)
        if not m:
            digits = re.sub(r"\\D+", "", raw)
            if len(digits) >= 8:
                return digits[:8]
            d = date.today() + timedelta(days=1)
            return f"{d.year:04d}{d.month:02d}{d.day:02d}"
        y, mm, dd = int(m.group(1)), int(m.group(2)), int(m.group(3))
        return f"{y:04d}{mm:02d}{dd:02d}"

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

        left_panel = ttk.Frame(top_bar)
        left_panel.pack(side=tk.LEFT, anchor=tk.NW, fill=tk.X, expand=True)

        action_row = ttk.Frame(left_panel)
        action_row.pack(side=tk.TOP, anchor=tk.W, fill=tk.X)
        ttk.Button(action_row, text="選擇主標籤…", command=self._open_primary_tag_picker).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Button(action_row, text="排序主標籤…", command=self._open_primary_tag_order_picker).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Button(action_row, text="依目前選取更新名單", command=self._refresh_primary_filter_results).pack(
            side=tk.LEFT
        )

        sort_row = ttk.Frame(left_panel)
        sort_row.pack(side=tk.TOP, anchor=tk.W, fill=tk.X, pady=(6, 0))
        ttk.Label(sort_row, text="名單排序：").pack(side=tk.LEFT, padx=(0, 4))
        self._pf_name_sort_key_cb = ttk.Combobox(
            sort_row,
            state="readonly",
            width=10,
            textvariable=self._pf_name_sort_key_var,
            values=[label for _, label in _PF_NAME_SORT_OPTIONS],
        )
        key_to_label = {k: lab for k, lab in _PF_NAME_SORT_OPTIONS}
        self._pf_name_sort_key_cb.set(
            key_to_label.get(self._pf_name_sort_key_var.get(), key_to_label["source"])
        )
        self._pf_name_sort_key_cb.bind("<<ComboboxSelected>>", self._on_pf_name_sort_changed)
        self._pf_name_sort_key_cb.pack(side=tk.LEFT, padx=(0, 6))
        self._pf_name_sort_dir_cb = ttk.Combobox(
            sort_row,
            state="readonly",
            width=6,
            textvariable=self._pf_name_sort_dir_var,
            values=["升冪", "降冪"],
        )
        self._pf_name_sort_dir_cb.set("降冪" if self._pf_name_sort_dir_var.get() == "desc" else "升冪")
        self._pf_name_sort_dir_cb.bind("<<ComboboxSelected>>", self._on_pf_name_sort_changed)
        self._pf_name_sort_dir_cb.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(sort_row, text="匯出…", command=self._open_export_from_primary_filter).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Button(sort_row, text="A4 PDF列印…", command=self._open_primary_filter_pdf_dialog).pack(
            side=tk.LEFT
        )

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
    def _apply_tag_order(items: list[str], saved_order: list[str]) -> list[str]:
        keep = [str(x).strip() for x in items if str(x).strip()]
        if not keep:
            return []
        pos = {v: i for i, v in enumerate(keep)}
        out: list[str] = []
        used: set[str] = set()
        for t in saved_order or []:
            if t in pos and t not in used:
                used.add(t)
                out.append(t)
        for t in keep:
            if t not in used:
                out.append(t)
        return out

    def _pf_name_sort_key(self) -> str:
        cb = getattr(self, "_pf_name_sort_key_cb", None)
        label = (cb.get() if cb is not None else "") or ""
        label = label.strip()
        for key, lab in _PF_NAME_SORT_OPTIONS:
            if lab == label:
                return key
        return "source"

    def _pf_name_sort_desc(self) -> bool:
        cb = getattr(self, "_pf_name_sort_dir_cb", None)
        v = (cb.get() if cb is not None else "") or ""
        return v.strip() == "降冪"

    def _sort_primary_filter_matches(self, matches: list) -> list:
        mode = self._pf_name_sort_key()
        if mode == "source":
            return list(matches)

        def serial_num(r) -> int:
            s = str(getattr(r, "serial", "") or "").strip()
            m = re.search(r"\d+", s)
            return int(m.group(0)) if m else 0

        def name_key(r) -> str:
            return (getattr(r, "customer_name", "") or "").strip()

        def headcount_num(r) -> int:
            s = self._row_headcount_str(r) or ""
            m = re.search(r"\d+", s)
            return int(m.group(0)) if m else -1

        if mode == "serial":
            key_fn = lambda r: (serial_num(r), name_key(r))
        elif mode == "name":
            key_fn = lambda r: (name_key(r), serial_num(r))
        elif mode == "headcount":
            key_fn = lambda r: (headcount_num(r), serial_num(r), name_key(r))
        else:
            return list(matches)
        return sorted(matches, key=key_fn, reverse=self._pf_name_sort_desc())

    def _on_pf_name_sort_changed(self, _evt: tk.Event | None = None) -> None:
        self._pages_state["pf_name_sort_key"] = self._pf_name_sort_key()
        self._pages_state["pf_name_sort_dir"] = "desc" if self._pf_name_sort_desc() else "asc"
        try:
            save_input_pages_state(self._pages_state)
        except OSError:
            pass
        if self._filter_selected_tags and self._rows:
            self._rebuild_primary_filter_results()

    @staticmethod
    def _merge_selected_order(prev_order: list[str], selected: list[str], base_order: list[str]) -> list[str]:
        sel = [x for x in selected if x]
        seen = set(sel)
        out = [x for x in (prev_order or []) if x in seen]
        used = set(out)
        for x in base_order:
            if x in seen and x not in used:
                out.append(x)
                used.add(x)
        for x in sel:
            if x not in used:
                out.append(x)
                used.add(x)
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

    def _filter_footer_disposable_applies(
        self,
        r,
        tags: list[str],
        keys_by_tag: dict[str, set],
        person_key,
    ) -> bool:
        if not self._row_has_disposable_in_data(r):
            return False
        k = person_key(r)
        for t in tags:
            if not self._get_display_rule(t).get("disposable"):
                continue
            if k in keys_by_tag.get(t, set()):
                return True
        return False

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

    @staticmethod
    def _row_has_utensil_in_data(r) -> bool:
        """原始資料或姓名／備註／標籤等是否含「自備餐具」（須完整詞）。"""
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
        return _UTENSIL_MARKER in hay

    def _fenji_stat_bucket(self, r, rule: dict[str, bool]) -> str:
        """分量分計：(拋) 依勾選+資料；(自) 依資料「自備餐具」；拋優先於自。"""
        rule = normalize_display_rule(rule)
        if rule["disposable"] and self._row_has_disposable_in_data(r):
            return "disp"
        if self._row_has_utensil_in_data(r):
            return "ut"
        return "plain"

    def _crosstab_page_key(self, r) -> str:
        """交叉表資料頁欄鍵（與匯出頁尾一致）。"""
        return (getattr(r, "source_page", None) or "").strip() or "（無頁名）"

    def _name_roster_frame_kind(self, r, rule: dict[str, bool]) -> str | None:
        """姓名外框：拋／自須勾選對應選項；同列兩者皆有時拋優先。"""
        rule = normalize_display_rule(rule)
        if not rule["name"]:
            return None
        if rule["disposable"] and self._row_has_disposable_in_data(r):
            return "disp"
        if rule.get("utensil") and self._row_has_utensil_in_data(r):
            return "utens"
        return None

    def _roster_segments(self, r, rule: dict[str, bool]) -> list[tuple[str, str | None]]:
        """(片段文字, 外框類型)：None／disp（藍）／utens（綠），僅姓名可帶框。"""
        rule = normalize_display_rule(rule)
        sz = headcount_size_label(self._row_headcount_str(r))
        name_frame = self._name_roster_frame_kind(r, rule)
        segs: list[tuple[str, str | None]] = []
        if rule["serial"]:
            segs.append((format_order_serial(r.serial), None))
        if rule.get("page_tag"):
            pg = (getattr(r, "source_page", None) or "").strip()
            if pg:
                segs.append((pg, None))
        if rule["name"]:
            nm = (r.customer_name or "").strip() or "（無姓名）"
            segs.append((nm, name_frame))
        if rule["size_label"] and sz:
            segs.append((f"({sz})", None))
        return segs

    def _roster_plain_width(self, r, rule: dict[str, bool]) -> float:
        fnt = tkfont.Font(font=FILTER_ROSTER_FONT)
        segs = self._roster_segments(r, rule)
        if not segs:
            return float(fnt.measure(" "))
        sp = float(fnt.measure(" "))
        w = 0.0
        for i, (text, frame_kind) in enumerate(segs):
            if i > 0:
                w += sp
            w += float(fnt.measure(text))
            if frame_kind:
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
        rows_segs: list[list[tuple[str, str | None]]] = []
        for r in matches:
            sz = headcount_size_label(self._row_headcount_str(r))
            fk = self._fenji_stat_bucket(r, rule)
            if sz == "小":
                small_n += 1
                if fk == "disp":
                    small_disp += 1
            elif sz == "大":
                large_n += 1
                if fk == "disp":
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
                for i, (text, frame_kind) in enumerate(segs):
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
                    if bb and frame_kind:
                        p = 2
                        outline = (
                            "#0D47A1"
                            if frame_kind == "disp"
                            else FILTER_UTENSIL_OUTLINE
                        )
                        rid = cv.create_rectangle(
                            bb[0] - p,
                            bb[1] - p,
                            bb[2] + p,
                            bb[3] + p,
                            outline=outline,
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
        fenji_text = self._format_block_fenji_one_line(matches, rule)
        fenji_widget = tk.Text(
            stat_frame,
            height=1,
            wrap=tk.NONE,
            bg=bg,
            fg="#0D47A1",
            font=FILTER_STAT_FONT,
            relief=tk.FLAT,
            bd=0,
            highlightthickness=0,
            takefocus=0,
            cursor="arrow",
        )
        fenji_widget.tag_configure("num", foreground="#C62828")
        i = 0
        for m in re.finditer(r"\d+", fenji_text):
            s, e = m.span()
            if s > i:
                fenji_widget.insert(tk.END, fenji_text[i:s])
            fenji_widget.insert(tk.END, fenji_text[s:e], ("num",))
            i = e
        if i < len(fenji_text):
            fenji_widget.insert(tk.END, fenji_text[i:])
        fenji_widget.configure(state=tk.DISABLED)
        fenji_widget.pack(anchor=tk.W, fill=tk.X)

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
    ) -> None:
        """置頂合計：資料頁為欄、小／大為列（以所有資料頁資料計算，不依下方區塊範圍）。"""
        host = getattr(self, "_filter_summary_host", None)
        if host is None or not uniq:
            return

        page_keys = sorted(
            {
                (getattr(r, "source_page", None) or "").strip() or "（無頁名）"
                for r in uniq.values()
            }
        )
        spg: dict[str, int] = {p: 0 for p in page_keys}
        sdg: dict[str, int] = {p: 0 for p in page_keys}
        sug: dict[str, int] = {p: 0 for p in page_keys}
        lpg: dict[str, int] = {p: 0 for p in page_keys}
        ldg: dict[str, int] = {p: 0 for p in page_keys}
        lug: dict[str, int] = {p: 0 for p in page_keys}
        og: dict[str, int] = {p: 0 for p in page_keys}

        for r in uniq.values():
            pg = (getattr(r, "source_page", None) or "").strip() or "（無頁名）"
            sz = headcount_size_label(self._row_headcount_str(r))
            # 置頂總表固定以資料本身判定（不受下方已選標籤規則影響）
            disp = self._row_has_disposable_in_data(r)
            ut = (not disp) and self._row_has_utensil_in_data(r)
            if sz == "小":
                if disp:
                    sdg[pg] = sdg.get(pg, 0) + 1
                elif ut:
                    sug[pg] = sug.get(pg, 0) + 1
                else:
                    spg[pg] = spg.get(pg, 0) + 1
            elif sz == "大":
                if disp:
                    ldg[pg] = ldg.get(pg, 0) + 1
                elif ut:
                    lug[pg] = lug.get(pg, 0) + 1
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
            if hdr:
                tk.Label(
                    parent,
                    text=text,
                    font=font_hdr if bold else font_cell,
                    bg=bg,
                    fg="#0D47A1",
                    padx=10,
                    pady=6,
                    relief=tk.FLAT,
                    borderwidth=1,
                    highlightthickness=1,
                    highlightbackground="#CFD8DC",
                ).grid(row=r, column=c, sticky=tk.NSEW, padx=1, pady=1)
                return

            w = tk.Text(
                parent,
                width=max(8, len(text) + 2),
                height=1,
                font=font_hdr if bold else font_cell,
                bg=bg,
                relief=tk.FLAT,
                borderwidth=1,
                highlightthickness=1,
                highlightbackground="#CFD8DC",
                padx=10,
                pady=6,
                wrap=tk.NONE,
                takefocus=0,
                cursor="arrow",
            )
            w.tag_configure("base", foreground="#212121")
            w.tag_configure("num", foreground="#C62828")
            i = 0
            for m in re.finditer(r"\d+", text):
                s, e = m.span()
                if s > i:
                    w.insert(tk.END, text[i:s], ("base",))
                w.insert(tk.END, text[s:e], ("num",))
                i = e
            if i < len(text):
                w.insert(tk.END, text[i:], ("base",))
            w.configure(state=tk.DISABLED)
            w.grid(row=r, column=c, sticky=tk.NSEW, padx=1, pady=1)

        gridf = tk.Frame(host, bg="#B0BEC5", padx=1, pady=1)
        gridf.pack(side=tk.RIGHT, anchor=tk.N, padx=2, pady=2)

        npg = len(page_keys)
        _cell(gridf, 0, 0, "", hdr=True)
        for j, pk in enumerate(page_keys):
            _cell(gridf, 0, j + 1, pk, hdr=True)
        _cell(gridf, 0, npg + 1, "合計", hdr=True)

        def _pn_cell(sp: int, sd: int, su: int) -> str:
            return f"{su}(自)+{sp}+{sd}(拋)"

        gs_plain = sum(spg.get(pk, 0) for pk in page_keys)
        gs_disp = sum(sdg.get(pk, 0) for pk in page_keys)
        gs_ut = sum(sug.get(pk, 0) for pk in page_keys)
        gl_plain = sum(lpg.get(pk, 0) for pk in page_keys)
        gl_disp = sum(ldg.get(pk, 0) for pk in page_keys)
        gl_ut = sum(lug.get(pk, 0) for pk in page_keys)

        _cell(gridf, 1, 0, "大", hdr=True)
        for j, pk in enumerate(page_keys):
            lp, ld, lu = lpg.get(pk, 0), ldg.get(pk, 0), lug.get(pk, 0)
            _cell(gridf, 1, j + 1, _pn_cell(lp, ld, lu))
        _cell(gridf, 1, npg + 1, _pn_cell(gl_plain, gl_disp, gl_ut), sumc=True)

        _cell(gridf, 2, 0, "小", hdr=True)
        for j, pk in enumerate(page_keys):
            sp, sd, su = spg.get(pk, 0), sdg.get(pk, 0), sug.get(pk, 0)
            _cell(gridf, 2, j + 1, _pn_cell(sp, sd, su))
        _cell(gridf, 2, npg + 1, _pn_cell(gs_plain, gs_disp, gs_ut), sumc=True)

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

    def _open_tag_order_dialog(
        self,
        *,
        title: str,
        tags: list[str],
        on_save,
    ) -> None:
        if not tags:
            messagebox.showinfo(title, "目前沒有可排序的標籤。", parent=self.root)
            return
        dlg = tk.Toplevel(self.root)
        dlg.title(title)
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.geometry("420x520")
        dlg.minsize(360, 360)

        body = ttk.Frame(dlg, padding=8)
        body.pack(fill=tk.BOTH, expand=True)
        ttk.Label(body, text="選取一項後可上下移動，按確定儲存順序。").pack(anchor=tk.W, pady=(0, 6))

        lb = tk.Listbox(body, exportselection=False)
        for t in tags:
            lb.insert(tk.END, t)
        lb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        if tags:
            lb.selection_set(0)
            lb.activate(0)

        sb = ttk.Scrollbar(body, orient=tk.VERTICAL, command=lb.yview)
        lb.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.LEFT, fill=tk.Y, padx=(6, 0))

        side = ttk.Frame(body)
        side.pack(side=tk.LEFT, fill=tk.Y, padx=(8, 0))

        def _move(delta: int) -> None:
            sel = lb.curselection()
            if not sel:
                return
            i = int(sel[0])
            j = i + delta
            if j < 0 or j >= lb.size():
                return
            val = lb.get(i)
            lb.delete(i)
            lb.insert(j, val)
            lb.selection_clear(0, tk.END)
            lb.selection_set(j)
            lb.activate(j)
            lb.see(j)

        ttk.Button(side, text="上移", command=lambda: _move(-1)).pack(fill=tk.X, pady=(0, 6))
        ttk.Button(side, text="下移", command=lambda: _move(1)).pack(fill=tk.X)

        btnf = ttk.Frame(dlg, padding=8)
        btnf.pack(fill=tk.X)

        def _ok() -> None:
            ordered = [str(x) for x in lb.get(0, tk.END)]
            on_save(ordered)
            dlg.destroy()

        ttk.Button(btnf, text="確定", command=_ok).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(btnf, text="取消", command=dlg.destroy).pack(side=tk.RIGHT)

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
        dlg.geometry("960x560")
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

        # 主勾選＋顯示規則；拋棄式／自備餐具勾選且資料含對應詞時對姓名加外框
        vars_by: dict[str, tk.IntVar] = {}
        serial_vars: dict[str, tk.IntVar] = {}
        page_tag_vars: dict[str, tk.IntVar] = {}
        name_vars: dict[str, tk.IntVar] = {}
        size_vars: dict[str, tk.IntVar] = {}
        disposable_vars: dict[str, tk.IntVar] = {}
        utensil_vars: dict[str, tk.IntVar] = {}
        # 序號／資料頁／姓名／人數／拋棄式／自備餐具
        inner.grid_columnconfigure(0, weight=0)
        inner.grid_columnconfigure(1, weight=0, minsize=52)
        for col in range(2, 8):
            inner.grid_columnconfigure(col, uniform="picker_disp", minsize=72, weight=0)

        for ri, v in enumerate(values):
            vars_by[v] = tk.IntVar(value=1 if v in prev else 0)
            r0 = self._get_display_rule(v)
            serial_vars[v] = tk.IntVar(value=1 if r0["serial"] else 0)
            page_tag_vars[v] = tk.IntVar(value=1 if r0.get("page_tag") else 0)
            name_vars[v] = tk.IntVar(value=1 if r0["name"] else 0)
            size_vars[v] = tk.IntVar(value=1 if r0["size_label"] else 0)
            disposable_vars[v] = tk.IntVar(value=1 if r0.get("disposable") else 0)
            utensil_vars[v] = tk.IntVar(value=1 if r0.get("utensil") else 0)
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
                    ("自備餐具", utensil_vars),
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
            chosen = self._merge_selected_order(self._primary_tag_order, chosen, values)
            new_rules: dict[str, dict[str, bool]] = dict(self._filter_display_rules)
            for v in values:
                new_rules[v] = normalize_display_rule(
                    {
                        "serial": int(serial_vars[v].get() or 0) == 1,
                        "page_tag": int(page_tag_vars[v].get() or 0) == 1,
                        "name": int(name_vars[v].get() or 0) == 1,
                        "size_label": int(size_vars[v].get() or 0) == 1,
                        "disposable": int(disposable_vars[v].get() or 0) == 1,
                        "utensil": int(utensil_vars[v].get() or 0) == 1,
                    }
                )
            self._filter_selected_tags = chosen
            self._primary_tag_order = list(chosen)
            self._filter_display_rules = new_rules
            try:
                save_filter_prefs(
                    chosen,
                    new_rules,
                    self._filter_export_blocks,
                    self._get_export_templates_live(),
                    crosstab_col_tags=self._crosstab_col_tags,
                    primary_tag_order=self._primary_tag_order,
                    crosstab_tag_order=self._crosstab_tag_order,
                    export_tag_order=self._export_tag_order,
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

    def _open_primary_tag_order_picker(self) -> None:
        lib = set(list_hashtags())
        tags = [t for t in self._filter_selected_tags if t in lib]
        tags = self._apply_tag_order(tags, self._primary_tag_order)

        def _save(ordered: list[str]) -> None:
            self._primary_tag_order = list(ordered)
            self._filter_selected_tags = [t for t in ordered if t in set(self._filter_selected_tags)]
            try:
                save_filter_prefs(
                    self._filter_selected_tags,
                    self._filter_display_rules,
                    self._filter_export_blocks,
                    self._get_export_templates_live(),
                    crosstab_col_tags=self._crosstab_col_tags,
                    primary_tag_order=self._primary_tag_order,
                    crosstab_tag_order=self._crosstab_tag_order,
                    export_tag_order=self._export_tag_order,
                )
            except OSError as e:
                messagebox.showwarning("主標籤篩選", f"無法儲存排序：{e}", parent=self.root)
                return
            self._rebuild_primary_filter_results()
            self._status.set("主標籤篩選：已儲存排序。")

        self._open_tag_order_dialog(title="主標籤排序", tags=tags, on_save=_save)

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

        lib = set(list_hashtags())
        tags = [t for t in self._filter_selected_tags if t in lib]
        tags = self._apply_tag_order(tags, self._primary_tag_order)
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

        # 右上角總表：以所有資料頁（self._rows）計算，不依下方目前顯示的標籤區塊範圍。
        # 以「資料頁 + 序號 + 姓名」去重，避免同頁重複列造成重覆計數。
        uniq_all: dict[tuple[str, str, str], object] = {}
        for r in self._rows:
            k = (
                (getattr(r, "source_page", None) or "").strip(),
                str(getattr(r, "serial", "")).strip(),
                (getattr(r, "customer_name", "") or "").strip(),
            )
            uniq_all[k] = r
        if uniq_all:
            self._render_primary_filter_top_summary(uniq_all)

        bi = 0
        for tag in tags:
            matches = self._sort_primary_filter_matches(self._rows_matching_tag_value(tag))
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
        lib = set(list_hashtags())
        tags = [t for t in self._filter_selected_tags if t in lib]
        tags = self._apply_tag_order(tags, self._primary_tag_order)
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
        vis = self._apply_tag_order(self._visible_primary_filter_tags(), self._export_tag_order)
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
    ) -> tuple[int, int, int, int, int, int, int]:
        """回傳 (小, 大, 小含拋, 大含拋, 未標份量, 小含自, 大含自)；拋優先於自。"""
        rule = normalize_display_rule(rule)
        small_n = large_n = small_disp = large_disp = other_n = small_ut = large_ut = 0
        for r in matches:
            sz = headcount_size_label(self._row_headcount_str(r))
            fk = self._fenji_stat_bucket(r, rule)
            if sz == "小":
                small_n += 1
                if fk == "disp":
                    small_disp += 1
                elif fk == "ut":
                    small_ut += 1
            elif sz == "大":
                large_n += 1
                if fk == "disp":
                    large_disp += 1
                elif fk == "ut":
                    large_ut += 1
            else:
                other_n += 1
        return small_n, large_n, small_disp, large_disp, other_n, small_ut, large_ut

    def _format_block_fenji_one_line(self, matches: list, rule: dict[str, bool]) -> str:
        """區塊內分量單行：小／大（自+一般+拋）與混合計（大自+小自、大拋+小拋）。"""
        small_n, large_n, small_disp, large_disp, other_n, small_ut, large_ut = (
            self._count_size_breakdown(matches, rule)
        )
        sp = small_n - small_disp - small_ut
        lp = large_n - large_disp - large_ut
        sd, ld = small_disp, large_disp
        su, lu = small_ut, large_ut
        s = (
            f"分量分計｜ 大：{lu}(自)+{lp}+{ld}(拋)   ｜   小：{su}(自)+{sp}+{sd}(拋)   ｜   "
            f"混合計：大(自){lu}+小(自){su}   大(拋){ld}+小(拋){sd}"
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

        page_keys = sorted(
            {
                (getattr(r, "source_page", None) or "").strip() or "（無頁名）"
                for r in uniq.values()
            }
        )
        spg: dict[str, int] = {p: 0 for p in page_keys}
        sdg: dict[str, int] = {p: 0 for p in page_keys}
        sug: dict[str, int] = {p: 0 for p in page_keys}
        lpg: dict[str, int] = {p: 0 for p in page_keys}
        ldg: dict[str, int] = {p: 0 for p in page_keys}
        lug: dict[str, int] = {p: 0 for p in page_keys}
        og: dict[str, int] = {p: 0 for p in page_keys}

        for r in uniq.values():
            pg = (getattr(r, "source_page", None) or "").strip() or "（無頁名）"
            sz = headcount_size_label(self._row_headcount_str(r))
            disp = self._filter_footer_disposable_applies(r, tags, keys_by_tag, _person_key)
            ut = (not disp) and self._row_has_utensil_in_data(r)
            if sz == "小":
                if disp:
                    sdg[pg] += 1
                elif ut:
                    sug[pg] += 1
                else:
                    spg[pg] += 1
            elif sz == "大":
                if disp:
                    ldg[pg] += 1
                elif ut:
                    lug[pg] += 1
                else:
                    lpg[pg] += 1
            else:
                og[pg] += 1

        gs_plain = sum(spg.get(pk, 0) for pk in page_keys)
        gs_disp = sum(sdg.get(pk, 0) for pk in page_keys)
        gs_ut = sum(sug.get(pk, 0) for pk in page_keys)
        gl_plain = sum(lpg.get(pk, 0) for pk in page_keys)
        gl_disp = sum(ldg.get(pk, 0) for pk in page_keys)
        gl_ut = sum(lug.get(pk, 0) for pk in page_keys)
        go_total = sum(og.get(pk, 0) for pk in page_keys)

        def _pn(sp: int, sd: int, su: int) -> str:
            return f"{su}(自)+{sp}+{sd}(拋)"

        out_lines = [
            "\t".join(["", *page_keys, "合計"]),
            "\t".join(
                ["小"]
                + [_pn(spg[pk], sdg[pk], sug[pk]) for pk in page_keys]
                + [_pn(gs_plain, gs_disp, gs_ut)]
            ),
            "\t".join(
                ["大"]
                + [_pn(lpg[pk], ldg[pk], lug[pk]) for pk in page_keys]
                + [_pn(gl_plain, gl_disp, gl_ut)]
            ),
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
            for r in self._sort_primary_filter_matches(self._rows_matching_tag_value(tag)):
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
        tags = self._apply_tag_order(
            [str(t).strip() for t in tags_subset if str(t).strip()],
            self._export_tag_order,
        )
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
            matches = self._sort_primary_filter_matches(self._rows_matching_tag_value(tag))
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

    def _open_export_tag_order_picker(self) -> None:
        tags = self._apply_tag_order(self._visible_primary_filter_tags(), self._export_tag_order)

        def _save(ordered: list[str]) -> None:
            self._export_tag_order = list(ordered)
            try:
                save_filter_prefs(
                    self._filter_selected_tags,
                    self._filter_display_rules,
                    self._filter_export_blocks,
                    self._get_export_templates_live(),
                    crosstab_col_tags=self._crosstab_col_tags,
                    primary_tag_order=self._primary_tag_order,
                    crosstab_tag_order=self._crosstab_tag_order,
                    export_tag_order=self._export_tag_order,
                )
            except OSError as e:
                messagebox.showwarning("匯出", f"無法儲存匯出排序：{e}", parent=self.root)
                return
            self._status.set("匯出：已儲存標籤排序。")
            if self._export_preview_docx_path is not None:
                self._refresh_export_tab_preview(None)

        self._open_tag_order_dialog(title="匯出標籤排序", tags=tags, on_save=_save)

    def _open_primary_filter_pdf_dialog(self) -> None:
        try:
            from export_print_pdf import save_primary_filter_pdf
        except ImportError:
            messagebox.showerror(
                "A4 PDF",
                "請先安裝 reportlab：\npip install reportlab",
                parent=self.root,
            )
            return
        if not self._rows:
            messagebox.showinfo("A4 PDF", "請先在「輸入與分析」完成分析。", parent=self.root)
            return
        vis = self._apply_tag_order(self._visible_primary_filter_tags(), self._export_tag_order)
        if not vis:
            messagebox.showinfo(
                "A4 PDF",
                "目前沒有可列印的主標籤（請確認已選主標籤且有命中資料）。",
                parent=self.root,
            )
            return
        pre = self._tags_checked_for_export(vis)
        if not pre:
            pre = list(vis)

        dlg = tk.Toplevel(self.root)
        dlg.title("A4 PDF：選擇要列印的主標籤")
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.geometry("520x480")
        dlg.minsize(400, 320)

        ttk.Label(
            dlg,
            text="勾選要輸出到 PDF 的標籤（順序依「匯出標籤排序」）；各區含筆數與分量統計、名單表。",
            wraplength=480,
        ).pack(anchor=tk.W, padx=10, pady=(10, 6))

        opt = ttk.Frame(dlg)
        opt.pack(fill=tk.X, padx=10, pady=(0, 6))
        cols_var = tk.IntVar(value=int(self._pages_state.get("pdf_primary_name_cols") or 7))
        font_var = tk.StringVar(value=f"{float(self._pages_state.get('pdf_primary_name_font_size') or 7.8):.1f}")
        ttk.Label(opt, text="每排人數：").pack(side=tk.LEFT)
        tk.Spinbox(
            opt,
            from_=3,
            to=10,
            textvariable=cols_var,
            width=4,
            justify="center",
        ).pack(side=tk.LEFT, padx=(4, 10))
        ttk.Label(opt, text="名字字級：").pack(side=tk.LEFT)
        tk.Spinbox(
            opt,
            from_=6.0,
            to=12.0,
            increment=0.2,
            textvariable=font_var,
            width=5,
            justify="center",
        ).pack(side=tk.LEFT, padx=(4, 0))

        mid = ttk.Frame(dlg)
        mid.pack(fill=tk.BOTH, expand=True, padx=10, pady=4)
        canvas = tk.Canvas(mid, highlightthickness=0)
        sb = ttk.Scrollbar(mid, orient=tk.VERTICAL, command=canvas.yview)
        inner = ttk.Frame(canvas, padding=4)
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

        vars_by: dict[str, tk.IntVar] = {}
        for i, tag in enumerate(vis):
            n = len(self._rows_matching_tag_value(tag))
            vars_by[tag] = tk.IntVar(value=1 if tag in pre else 0)
            tk.Checkbutton(
                inner,
                text=f"{tag}　（{n} 筆）",
                variable=vars_by[tag],
                anchor=tk.W,
            ).grid(row=i, column=0, sticky=tk.W, pady=2)

        btnf = ttk.Frame(dlg, padding=8)
        btnf.pack(fill=tk.X)

        def _all(v: int) -> None:
            for iv in vars_by.values():
                iv.set(v)

        def _ok() -> None:
            chosen = [t for t in vis if int(vars_by[t].get() or 0) == 1]
            if not chosen:
                messagebox.showwarning("A4 PDF", "請至少勾選一個標籤。", parent=dlg)
                return
            try:
                ncols = int(cols_var.get())
            except Exception:
                ncols = 7
            try:
                nfont = float(font_var.get())
            except Exception:
                nfont = 7.8
            ncols = min(10, max(3, ncols))
            nfont = min(12.0, max(6.0, nfont))
            dest = filedialog.asksaveasfilename(
                parent=self.root,
                defaultextension=".pdf",
                filetypes=[("PDF (*.pdf)", "*.pdf"), ("全部", "*.*")],
                initialfile=f"主篩選名單_{self._pdf_default_date_stamp()}.pdf",
                title="另存主篩選名單 PDF…",
            )
            if not dest:
                return
            try:
                save_primary_filter_pdf(self, dest, chosen, name_cols=ncols, name_font_size=nfont)
            except Exception as e:
                messagebox.showerror("A4 PDF", str(e), parent=self.root)
                return
            self._pages_state["pdf_primary_name_cols"] = ncols
            self._pages_state["pdf_primary_name_font_size"] = nfont
            try:
                save_input_pages_state(self._pages_state)
            except OSError:
                pass
            dlg.destroy()
            self._status.set(f"已輸出主篩選 PDF：{dest}")
            try:
                self._open_path_in_default_app(Path(dest))
            except OSError as e:
                messagebox.showwarning("A4 PDF", f"檔案已輸出，但無法自動開啟：{e}", parent=self.root)

        ttk.Button(btnf, text="全選", command=lambda: _all(1)).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btnf, text="全不選", command=lambda: _all(0)).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Button(btnf, text="確定", command=_ok).pack(side=tk.RIGHT, padx=(6, 0))
        ttk.Button(btnf, text="取消", command=dlg.destroy).pack(side=tk.RIGHT)
        dlg.bind("<Escape>", lambda _e: dlg.destroy())

    def _export_crosstab_pdf_dialog(self) -> None:
        try:
            from export_print_pdf import save_crosstab_pdf
        except ImportError:
            messagebox.showerror(
                "A4 PDF",
                "請先安裝 reportlab：\npip install reportlab",
                parent=self.root,
            )
            return
        dest = filedialog.asksaveasfilename(
            parent=self.root,
            defaultextension=".pdf",
            filetypes=[("PDF (*.pdf)", "*.pdf"), ("全部", "*.*")],
            initialfile=f"交叉表_{self._pdf_default_date_stamp()}.pdf",
            title="另存交叉表 PDF…",
        )
        if not dest:
            return
        try:
            save_crosstab_pdf(self, dest)
        except Exception as e:
            messagebox.showerror("A4 PDF", str(e), parent=self.root)
            return
        self._status.set(f"已輸出交叉表 PDF：{dest}")
        try:
            self._open_path_in_default_app(Path(dest))
        except OSError as e:
            messagebox.showwarning("A4 PDF", f"檔案已輸出，但無法自動開啟：{e}", parent=self.root)

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

    # --- 分頁：交叉表（互斥分量列 × 共用標籤欄 × 資料頁分塊）---
    def _crosstab_row_category(self, r) -> str:
        """人數區間：小／大／未標示（與主標籤篩選一致）。"""
        sz = headcount_size_label(self._row_headcount_str(r))
        return sz if sz in ("小", "大") else "未標示"

    def _crosstab_partition_label(self, r, rule: dict[str, bool]) -> str:
        """交叉表分量列鍵：與小/大/未標互斥叉拋/自，共九類（同主篩選拋優先於自）。"""
        rule = normalize_display_rule(rule)
        sz = headcount_size_label(self._row_headcount_str(r))
        fk = self._fenji_stat_bucket(r, rule)
        if sz == "小":
            if fk == "disp":
                return "小拋"
            if fk == "ut":
                return "小自"
            return "小"
        if sz == "大":
            if fk == "disp":
                return "大拋"
            if fk == "ut":
                return "大自"
            return "大"
        if fk == "disp":
            return "未標拋"
        if fk == "ut":
            return "未標自"
        return "未標"

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
            chosen = self._merge_selected_order(self._crosstab_tag_order, chosen, values)
            self._crosstab_col_tags = chosen
            self._crosstab_tag_order = list(chosen)
            try:
                save_filter_prefs(
                    self._filter_selected_tags,
                    self._filter_display_rules,
                    self._filter_export_blocks,
                    self._get_export_templates_live(),
                    crosstab_col_tags=chosen,
                    primary_tag_order=self._primary_tag_order,
                    crosstab_tag_order=self._crosstab_tag_order,
                    export_tag_order=self._export_tag_order,
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

    def _open_crosstab_tag_order_picker(self) -> None:
        tags = self._apply_tag_order(list(self._crosstab_col_tags), self._crosstab_tag_order)

        def _save(ordered: list[str]) -> None:
            self._crosstab_tag_order = list(ordered)
            self._crosstab_col_tags = [t for t in ordered if t in set(self._crosstab_col_tags)]
            try:
                save_filter_prefs(
                    self._filter_selected_tags,
                    self._filter_display_rules,
                    self._filter_export_blocks,
                    self._get_export_templates_live(),
                    crosstab_col_tags=self._crosstab_col_tags,
                    primary_tag_order=self._primary_tag_order,
                    crosstab_tag_order=self._crosstab_tag_order,
                    export_tag_order=self._export_tag_order,
                )
            except OSError as e:
                messagebox.showwarning("交叉表", f"無法儲存排序：{e}", parent=self.root)
                return
            self._update_crosstab_cols_summary()
            self._refresh_crosstab_table()
            self._status.set("交叉表：已儲存欄標籤排序。")

        self._open_tag_order_dialog(title="交叉表欄標籤排序", tags=tags, on_save=_save)

    def _compute_crosstab_matrix(
        self,
    ) -> tuple[
        tuple[
            tuple[str, ...],
            list[str],
            list[tuple[str, list[list[int]], list[int], list[int], int]],
            list[int],
            int,
        ]
        | None,
        str | None,
    ]:
        """
        標籤欄：全資料頁加總仍為 0 的欄位不顯示（其餘欄各頁對齊）。
        分量列互斥：小、小拋、小自、大、大拋、大自、未標、未標拋、未標自。
        小／大／未標來自人數區間；拋／自依 _fenji_stat_bucket。
        回傳 grand_col_totals：各可見標籤欄在所有資料頁儲存格加總（= 台南該欄 + 高雄該欄 + …）。
        grand = 全表儲存格加總（同 grand_col_totals 之總和；非去重人數）。
        """
        col_tags = self._apply_tag_order(list(self._crosstab_col_tags), self._crosstab_tag_order)
        row_defs = _CROSSTAB_PARTITION_ROWS
        rix = {k: i for i, k in enumerate(row_defs)}
        nkind = len(row_defs)

        if not self._rows:
            return None, "no_rows"
        if not col_tags:
            return None, "no_cols"

        page_keys = sorted({self._crosstab_page_key(r) for r in self._rows})
        nt = len(col_tags)
        per_page: dict[str, list[list[int]]] = {p: [[0] * nt for _ in range(nkind)] for p in page_keys}
        rows_by_page: dict[str, list] = {p: [] for p in page_keys}
        for r in self._rows:
            rows_by_page[self._crosstab_page_key(r)].append(r)

        # 先鎖定資料頁，再做標籤命中，避免任何跨頁混算。
        for page in page_keys:
            rows_here = rows_by_page.get(page, [])
            for j, tag in enumerate(col_tags):
                rule = self._get_display_rule(tag)
                for r in rows_here:
                    if not any(t.get("value") == tag for t in getattr(r, "tags", []) or []):
                        continue
                    lbl = self._crosstab_partition_label(r, rule)
                    per_page[page][rix[lbl]][j] += 1

        col_total_all = [
            sum(per_page[p][i][j] for p in page_keys for i in range(nkind))
            for j in range(nt)
        ]
        active_cols = [j for j in range(nt) if col_total_all[j] > 0]
        if not active_cols:
            return None, "all_zero_cols"

        grand_col_totals = [col_total_all[j] for j in active_cols]
        grand = sum(grand_col_totals)
        col_tags_f = [col_tags[j] for j in active_cols]
        nc = len(active_cols)

        active_rows = [
            i
            for i in range(nkind)
            if sum(per_page[p][i][j] for p in page_keys for j in active_cols) > 0
        ]
        if not active_rows:
            return None, "all_zero_cols"

        row_labels = tuple(row_defs[i] for i in active_rows)
        page_blocks: list[tuple[str, list[list[int]], list[int], list[int], int]] = []
        for page in page_keys:
            mat_f = [
                [per_page[page][i][j] for j in active_cols] for i in active_rows
            ]
            nr = len(active_rows)
            row_sums_f = [sum(mat_f[ri][j] for j in range(nc)) for ri in range(nr)]
            col_sums_f = [sum(mat_f[ri][j] for ri in range(nr)) for j in range(nc)]
            blk = sum(row_sums_f)
            page_blocks.append((page, mat_f, row_sums_f, col_sums_f, blk))

        return (row_labels, col_tags_f, page_blocks, grand_col_totals, grand), None

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

        row_kinds, col_tags, page_blocks, grand_col_totals, grand = data
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
                for page, mat, row_sums, col_sums, blk in page_blocks:
                    w.writerow([f"【{page}】"])
                    w.writerow(["分量", *col_tags, "列合計"])
                    for i, rk in enumerate(row_kinds):
                        w.writerow(
                            [rk, *[str(mat[i][j]) for j in range(len(col_tags))], str(row_sums[i])]
                        )
                    w.writerow(
                        ["欄合計", *[str(col_sums[j]) for j in range(len(col_tags))], str(blk)]
                    )
                    w.writerow([])
                w.writerow(
                    [
                        "標籤合計（右下為全表儲存格加總）",
                        *[str(x) for x in grand_col_totals],
                        str(grand),
                    ]
                )
        except OSError as e:
            messagebox.showerror("交叉表", f"無法寫入檔案：{e}", parent=self.root)
            return
        self._status.set(f"交叉表已匯出：{dest}")

    def _update_crosstab_cols_summary(self) -> None:
        if not getattr(self, "_crosstab_cols_summary", None):
            return
        tags = self._apply_tag_order(list(self._crosstab_col_tags), self._crosstab_tag_order)
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
        row_kinds, col_tags, page_blocks, grand_col_totals, grand = data
        n_kind_rows = len(row_kinds)
        nloc = len(col_tags)

        hdr_bg = "#E3F2FD"
        cell_bg = "#FAFAFA"
        sum_bg = "#FFF8E1"
        sect_bg = "#E8EAF6"
        total_row_bg = "#E1F5FE"
        # 分量資料列：小／大／未標 系列底色（含該列「列合計」格）
        cell_bg_small = "#E8F5E9"
        cell_bg_large = "#FFF3E0"
        cell_bg_unspec = "#ECEFF1"
        font_hdr = ("Microsoft JhengHei UI", 10, "bold")
        font_cell = ("Microsoft JhengHei UI", 10)

        def _partition_row_bg(rk: str) -> str:
            if rk.startswith("小"):
                return cell_bg_small
            if rk.startswith("大"):
                return cell_bg_large
            return cell_bg_unspec

        outer = tk.Frame(host)
        outer.pack(anchor=tk.NW, fill=tk.X)
        border = tk.Frame(outer, bg="#B0BEC5", padx=1, pady=1)
        border.pack(anchor=tk.NW, pady=(0, 8))
        gridf = tk.Frame(border, bg="#B0BEC5", padx=1, pady=1)
        gridf.pack(anchor=tk.NW)

        def make_add_lbl(parent: tk.Widget):
            def add_lbl(
                r: int,
                c: int,
                text: str,
                *,
                hdr: bool = False,
                sum_cell: bool = False,
                sect: bool = False,
                partition_bg: str | None = None,
                total_row: bool = False,
                wrap: int = 0,
                rowspan: int = 1,
                columnspan: int = 1,
            ) -> None:
                if sect:
                    bg = sect_bg
                    bold = True
                elif total_row:
                    bg = total_row_bg
                    bold = True
                elif partition_bg is not None:
                    bg = partition_bg
                    bold = hdr or sum_cell
                elif hdr:
                    bg = hdr_bg
                    bold = True
                elif sum_cell:
                    bg = sum_bg
                    bold = True
                else:
                    bg = cell_bg
                    bold = False
                kw: dict = dict(
                    font=font_hdr if bold else font_cell,
                    bg=bg,
                    fg="#0D47A1" if bold else "#212121",
                    padx=10,
                    pady=6,
                    relief=tk.FLAT,
                    borderwidth=1,
                    highlightthickness=1,
                    highlightbackground="#CFD8DC",
                )
                if wrap > 0:
                    kw["wraplength"] = wrap
                tk.Label(parent, text=text, **kw).grid(
                    row=r,
                    column=c,
                    rowspan=rowspan,
                    columnspan=columnspan,
                    sticky=tk.NSEW,
                    padx=1,
                    pady=1,
                )

            return add_lbl

        add_lbl = make_add_lbl(gridf)
        row = 0
        n_col_grid = nloc + 2

        for bi, (page, mat_f, row_sums_f, col_sums_f, blk_corner) in enumerate(page_blocks):
            if bi:
                row += 1
            add_lbl(row, 0, f"【{page}】", sect=True, columnspan=n_col_grid)
            row += 1
            add_lbl(row, 0, "分量", hdr=True)
            for j, tn in enumerate(col_tags):
                add_lbl(row, 1 + j, tn, hdr=True)
            add_lbl(row, 1 + nloc, "列合計", hdr=True)
            row += 1
            dr0 = row
            for i, rk in enumerate(row_kinds):
                pr = _partition_row_bg(rk)
                add_lbl(dr0 + i, 0, rk, hdr=True, partition_bg=pr)
                for j in range(nloc):
                    add_lbl(dr0 + i, 1 + j, str(mat_f[i][j]), partition_bg=pr)
                add_lbl(
                    dr0 + i,
                    1 + nloc,
                    str(row_sums_f[i]),
                    sum_cell=True,
                    partition_bg=pr,
                )
            sr = dr0 + n_kind_rows
            add_lbl(sr, 0, "欄合計", hdr=True)
            for j in range(nloc):
                add_lbl(sr, 1 + j, str(col_sums_f[j]), sum_cell=True)
            add_lbl(sr, 1 + nloc, str(blk_corner), sum_cell=True)
            row = sr + 1

        add_lbl(row, 0, "標籤合計", total_row=True)
        for j in range(nloc):
            add_lbl(row, 1 + j, str(grand_col_totals[j]), total_row=True)
        add_lbl(row, 1 + nloc, str(grand), total_row=True)

        gridf.grid_columnconfigure(0, weight=0, minsize=96)
        for c in range(1, n_col_grid):
            gridf.grid_columnconfigure(c, weight=1, uniform="ctab")

        tot_f = tk.Frame(outer, bg="#FFF8E1", padx=12, pady=10)
        tot_f.pack(anchor=tk.W, fill=tk.X, pady=(4, 0))
        tk.Label(
            tot_f,
            text=f"總計（與上表「標籤合計」右下相同）：{grand}",
            font=("Microsoft JhengHei UI", 11, "bold"),
            bg="#FFF8E1",
            fg="#0D47A1",
            anchor=tk.W,
            justify=tk.LEFT,
        ).pack(anchor=tk.W)

        note = tk.Label(
            outer,
            text=(
                "數量說明：列為互斥分量——小／大／未標依人數標籤（2～3→小、3～4→大）；"
                "小拋、小自、大拋、大自、未標拋、未標自 則再依拋棄式／自備餐具（與主篩選分量規則相同）。"
                "同一欄內 小+小拋+小自 為該標籤之小份量筆次合計；大、未標同理。"
                "表末「標籤合計」= 各資料頁同欄相加；"
                "所選標籤在全資料頁加總仍為 0 者自動不顯示。"
                "列合計橫向加總；全表總計為儲存格加總（同人多標籤會重複，非去重人數）。"
            ),
            font=("Microsoft JhengHei UI", 9),
            fg="#555555",
            anchor=tk.W,
            justify=tk.LEFT,
            wraplength=860,
        )
        note.pack(anchor=tk.W, pady=(10, 0))

        self._status.set("交叉表：已更新。")

    def _build_tab_crosstab(self) -> None:
        tab = ttk.Frame(self._nb, padding=8)
        self._tab_crosstab = tab
        self._nb.add(tab, text="交叉表")

        ttk.Label(
            tab,
            text=(
                "交叉表為單一表格分區：【資料頁】列為區塊標題，欄寬對齊；"
                "「標籤合計」列為每個欄標籤跨所有資料頁加總。"
                "所選欄標籤若在全資料頁加總仍為 0 則不顯示該欄。"
                "列為互斥分量（小／小拋／小自、大／大拋／大自、未標…）；規則與主篩選一致。"
                "「選擇欄標籤…」勾選會寫入設定檔，下次開啟還原。"
            ),
            wraplength=820,
        ).pack(anchor=tk.W, pady=(0, 6))

        bar = ttk.Frame(tab)
        bar.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(bar, text="選擇欄標籤…", command=self._open_crosstab_column_picker).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Button(bar, text="欄位排序…", command=self._open_crosstab_tag_order_picker).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Button(bar, text="更新表格", command=self._refresh_crosstab_table).pack(side=tk.LEFT)
        ttk.Button(bar, text="匯出試算表…", command=self._export_crosstab_spreadsheet).pack(
            side=tk.LEFT, padx=(12, 0)
        )
        ttk.Button(bar, text="A4 PDF列印…", command=self._export_crosstab_pdf_dialog).pack(
            side=tk.LEFT, padx=(12, 0)
        )
        self._crosstab_cols_summary = tk.StringVar(value="目前欄標籤：（尚未選擇）")
        ttk.Label(bar, textvariable=self._crosstab_cols_summary).pack(side=tk.LEFT, padx=(16, 0))

        self._crosstab_grid_host = ttk.Frame(tab)
        self._crosstab_grid_host.pack(fill=tk.BOTH, expand=True, anchor=tk.NW, pady=(8, 0))

        ttk.Label(
            tab,
            text=(
                "匯出 CSV：每資料頁一段＋最末「標籤合計」列（同畫面）。"
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
            "請用「用預設程式開啟」在 Word 中檢視。「與篩選區塊相同」為表格排版。需 pip install python-docx。"
            "A4 PDF 列印需 pip install reportlab（使用系統微軟正黑體）。下方可匯出 JSON／CSV。",
            wraplength=820,
        ).pack(anchor=tk.W, pady=(0, 8))

        pdf_row = ttk.Frame(tab)
        pdf_row.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(pdf_row, text="A4 PDF：").pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(pdf_row, text="主篩選名單…", command=self._open_primary_filter_pdf_dialog).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Button(pdf_row, text="交叉表…", command=self._export_crosstab_pdf_dialog).pack(side=tk.LEFT)

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
            text="匯出標籤排序…",
            command=self._open_export_tag_order_picker,
        ).pack(side=tk.LEFT, padx=(0, 10))
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

    def _build_tab_help(self) -> None:
        tab = ttk.Frame(self._nb, padding=8)
        self._tab_help = tab
        self._nb.add(tab, text="說明與教學")

        top = ttk.LabelFrame(tab, text="線上更新（GitHub Releases）", padding=8)
        top.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(top, text=f"目前版本：v{_APP_VERSION}").grid(row=0, column=0, sticky=tk.W)
        ttk.Button(top, text="檢查更新", command=self._check_updates_from_github).grid(
            row=1, column=0, sticky=tk.W, pady=(8, 0)
        )
        ttk.Label(top, textvariable=self._update_result_var, foreground="#0D47A1").grid(
            row=2, column=0, sticky=tk.W, pady=(8, 0)
        )

        guide = tk.Text(tab, height=18, wrap=tk.WORD, font=("Microsoft JhengHei UI", 10))
        guide.pack(fill=tk.BOTH, expand=True)
        guide.insert(
            "1.0",
            (
                "【本程式說明】\n"
                "1) 貼上資料後到「輸入與分析」執行分析。\n"
                "2) 到「主標籤篩選」勾選標籤並檢視名單。\n"
                "3) 依需求使用「匯出」或「A4 PDF列印」。\n\n"
                "【GitHub 更新教學（你自己發版）】\n"
                "A. 先在 GitHub 建立 Releases（例如 tag: v1.0.1）。\n"
                "B. 按「檢查更新」：若有新版本會提供開啟下載頁。\n"
                "D. 下載新版後手動替換執行檔（目前為安全模式，不自動覆蓋）。\n\n"
                "【建議的 Release 命名】\n"
                "- tag：v1.0.1\n"
                "- title：Menu Analyze v1.0.1\n"
                "- notes：列出修正與新增項目\n"
            ),
        )
        guide.configure(state=tk.DISABLED)

    @staticmethod
    def _version_key(v: str) -> tuple[int, ...]:
        s = (v or "").strip().lower().lstrip("v")
        nums = [int(x) for x in re.findall(r"\d+", s)]
        return tuple(nums) if nums else (0,)

    @staticmethod
    def _pick_release_download_url(data: dict, page_url: str) -> str:
        assets = data.get("assets")
        if isinstance(assets, list):
            # 優先給可執行或壓縮包，避免使用者不知要點哪個。
            for a in assets:
                if not isinstance(a, dict):
                    continue
                name = str(a.get("name") or "").lower()
                u = str(a.get("browser_download_url") or "").strip()
                if not u:
                    continue
                if name.endswith(".exe") or name.endswith(".zip"):
                    return u
            for a in assets:
                if not isinstance(a, dict):
                    continue
                u = str(a.get("browser_download_url") or "").strip()
                if u:
                    return u
        return page_url

    @staticmethod
    def _is_exe_url(url: str) -> bool:
        return (url or "").lower().split("?", 1)[0].endswith(".exe")

    def _download_update_exe(self, download_url: str, latest: str) -> Path | None:
        base = project_data_dir() / "updates"
        base.mkdir(parents=True, exist_ok=True)
        safe_ver = re.sub(r"[^0-9A-Za-z._-]+", "_", latest or "latest")
        out = base / f"menu_analyze_{safe_ver}.exe"
        req = urllib.request.Request(
            download_url,
            headers={"User-Agent": "menut-updater"},
        )
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                with open(out, "wb") as f:
                    shutil.copyfileobj(r, f)
        except Exception:
            return None
        return out if out.is_file() else None

    def _launch_swap_updater_and_exit(self, new_exe: Path) -> bool:
        if not getattr(sys, "frozen", False):
            return False
        target_exe = Path(sys.executable).resolve()
        pid = os.getpid()
        script = project_data_dir() / "updates" / "_apply_update.cmd"
        script.parent.mkdir(parents=True, exist_ok=True)
        lines = [
            "@echo off",
            "setlocal",
            f"set \"PID={pid}\"",
            f"set \"TARGET={target_exe}\"",
            f"set \"NEW={new_exe}\"",
            "",
            "for /l %%i in (1,1,90) do (",
            "  tasklist /FI \"PID eq %PID%\" | find \"%PID%\" >nul",
            "  if errorlevel 1 goto do_update",
            "  timeout /t 1 >nul",
            ")",
            "",
            ":do_update",
            "copy /y \"%NEW%\" \"%TARGET%\" >nul",
            "start \"\" \"%TARGET%\"",
            "del \"%NEW%\" >nul 2>nul",
            "del \"%~f0\" >nul 2>nul",
        ]
        try:
            script.write_text("\n".join(lines) + "\n", encoding="utf-8")
            subprocess.Popen(["cmd", "/c", str(script)], close_fds=True)
        except Exception:
            return False
        self.root.after(120, self.root.destroy)
        return True

    def _check_updates_from_github(self, *, silent: bool = False, show_latest_dialog: bool = True) -> None:
        repo = _UPDATE_REPO
        url = f"https://api.github.com/repos/{repo}/releases/latest"
        req = urllib.request.Request(
            url,
            headers={"Accept": "application/vnd.github+json", "User-Agent": "menut-update-checker"},
        )
        try:
            with urllib.request.urlopen(req, timeout=10) as r:
                raw = r.read().decode("utf-8", errors="replace")
                data = json.loads(raw)
        except urllib.error.HTTPError as e:
            if e.code == 404:
                self._update_result_var.set("找不到 Release（404）")
                if not silent:
                    messagebox.showerror(
                        "檢查更新",
                        "GitHub 回應 404。\n\n"
                        "常見原因：\n"
                        "1) 尚未建立任何 Release（/releases/latest 會 404）\n"
                        "2) 倉庫是 Private（未授權 API 也會 404）\n\n"
                        "請先確認 repo 可公開存取，並至少建立一個 release tag（例如 v1.0.1）。",
                        parent=self.root,
                    )
            else:
                self._update_result_var.set(f"檢查失敗：HTTP {e.code}")
                if not silent:
                    messagebox.showerror("檢查更新", f"GitHub 回應錯誤：HTTP {e.code}", parent=self.root)
            return
        except Exception as e:
            self._update_result_var.set("更新檢查失敗（網路或解析錯誤）")
            if not silent:
                messagebox.showerror("檢查更新", f"無法連線或解析更新資訊：{e}", parent=self.root)
            return

        latest = str(data.get("tag_name") or "").strip()
        latest_key = self._version_key(latest)
        cur_key = self._version_key(_APP_VERSION)
        page_url = str(data.get("html_url") or f"https://github.com/{repo}/releases").strip()
        download_url = self._pick_release_download_url(data, page_url)
        cur_major = cur_key[0] if cur_key else 0
        latest_major = latest_key[0] if latest_key else 0
        if latest and latest_key > cur_key:
            self._update_result_var.set(f"有新版本：{latest}（目前 v{_APP_VERSION}）")
            if latest_major > cur_major:
                ask = messagebox.askyesno(
                    "重大更新",
                    f"發現重大更新 {latest}（目前 v{_APP_VERSION}）。\n"
                    "此類版本通常不建議直接覆蓋，請重新下載新版。\n\n"
                    "要開啟下載頁嗎？",
                    parent=self.root,
                )
            else:
                ask = messagebox.askyesno(
                    "有新版本",
                    f"發現新版本 {latest}（目前 v{_APP_VERSION}）。\n要立即更新嗎？",
                    parent=self.root,
                )
            if ask:
                # 重大版本或非 exe 資產：引導人工下載。
                if latest_major > cur_major or not self._is_exe_url(download_url):
                    try:
                        webbrowser.open(download_url)
                    except Exception:
                        pass
                    return
                self._status.set("正在下載更新檔…")
                new_exe = self._download_update_exe(download_url, latest)
                if not new_exe:
                    self._status.set("更新下載失敗，已改為開啟下載頁。")
                    if not silent:
                        messagebox.showwarning("更新", "自動下載失敗，將改為開啟下載頁。", parent=self.root)
                    try:
                        webbrowser.open(download_url)
                    except Exception:
                        pass
                    return
                if not getattr(sys, "frozen", False):
                    self._status.set("目前為開發模式，無法自動覆蓋執行中程式。")
                    if not silent:
                        messagebox.showinfo(
                            "更新已下載",
                            f"已下載更新檔：\n{new_exe}\n\n開發模式不執行自動覆蓋，請手動使用此檔。",
                            parent=self.root,
                        )
                    return
                if messagebox.askyesno(
                    "準備套用更新",
                    f"已下載 {latest}。\n按「是」後將關閉程式並自動更新重啟。",
                    parent=self.root,
                ):
                    ok = self._launch_swap_updater_and_exit(new_exe)
                    if not ok and not silent:
                        messagebox.showerror("更新失敗", "無法啟動更新器，請改為手動更新。", parent=self.root)
        else:
            self._update_result_var.set(f"目前已是最新版本（v{_APP_VERSION}）")
            if show_latest_dialog and not silent:
                messagebox.showinfo("檢查更新", f"目前已是最新版本（v{_APP_VERSION}）。", parent=self.root)

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
        existing_before = set(list_hashtags())
        hashtag_clean = [str(v).strip() for v in hashtag_values if str(v).strip()]
        new_tags_this_run = [v for v in hashtag_clean if v not in existing_before]
        # 去重且保序
        seen_new: set[str] = set()
        new_tags_this_run = [v for v in new_tags_this_run if not (v in seen_new or seen_new.add(v))]
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

        if new_tags_this_run:
            preview = "、".join(new_tags_this_run[:30])
            if len(new_tags_this_run) > 30:
                preview += f"…（共 {len(new_tags_this_run)} 個）"
            messagebox.showinfo(
                "新增 #標籤",
                f"本次分析新增 {len(new_tags_this_run)} 個 #標籤：\n{preview}",
                parent=self.root,
            )

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
