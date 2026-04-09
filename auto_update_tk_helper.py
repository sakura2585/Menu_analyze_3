from __future__ import annotations

import sys
from pathlib import Path
from typing import Callable

import tkinter as tk
from tkinter import messagebox, ttk

from auto_update_template import GithubAutoUpdater, UpdateInfo


class TkAutoUpdateController:
    """
    Tkinter integration helper for GithubAutoUpdater.

    Features:
    - User preference checkbox: check updates on startup
    - Manual "Check updates now" button
    - Skip this version
    - Download + apply update (without auto restart)
    """

    def __init__(
        self,
        root: tk.Misc,
        updater: GithubAutoUpdater,
        *,
        app_version: str,
        on_status: Callable[[str], None] | None = None,
        on_after_apply_started: Callable[[], None] | None = None,
    ) -> None:
        self.root = root
        self.updater = updater
        self.app_version = app_version
        self.on_status = on_status
        self.on_after_apply_started = on_after_apply_started
        prefs = self.updater.load_prefs()
        self.var_check_on_startup = tk.BooleanVar(value=prefs.check_on_startup)

    def _status(self, text: str) -> None:
        if self.on_status:
            self.on_status(text)

    def build_controls(self, parent: tk.Misc) -> ttk.Frame:
        frame = ttk.Frame(parent)
        ttk.Checkbutton(
            frame,
            text="啟動時檢查更新",
            variable=self.var_check_on_startup,
            command=self._on_toggle_check_on_startup,
        ).pack(side=tk.LEFT)
        ttk.Button(frame, text="立即檢查更新", command=self.check_now).pack(side=tk.LEFT, padx=(8, 0))
        return frame

    def _on_toggle_check_on_startup(self) -> None:
        enabled = bool(self.var_check_on_startup.get())
        self.updater.set_check_on_startup(enabled)
        self._status("已啟用啟動檢查更新" if enabled else "已停用啟動檢查更新")

    def check_on_startup(self) -> None:
        self._check_and_prompt(force=False, silent=True)

    def check_now(self) -> None:
        self._check_and_prompt(force=True, silent=False)

    def _check_and_prompt(self, *, force: bool, silent: bool) -> None:
        try:
            info = self.updater.check_latest_if_enabled(force=force)
        except Exception as e:
            self._status("更新檢查失敗")
            if not silent:
                messagebox.showerror("檢查更新", f"無法檢查更新：{e}", parent=self.root)
            return

        if info is None:
            self._status("已略過啟動檢查更新")
            return
        if not info.has_update:
            self._status("目前已是最新版本")
            if force and not silent:
                messagebox.showinfo("檢查更新", f"目前已是最新版本（{self.app_version}）。", parent=self.root)
            return

        if not self.updater.should_offer_update(info):
            self._status(f"此版本已略過：{info.latest_version}")
            if force and not silent:
                if not messagebox.askyesno(
                    "已略過此版",
                    f"{info.latest_version} 已被標記略過。\n仍要查看並更新嗎？",
                    parent=self.root,
                ):
                    return
            else:
                return

        self._prompt_update(info, silent=silent)

    def _prompt_update(self, info: UpdateInfo, *, silent: bool) -> None:
        choice = messagebox.askyesnocancel(
            "有新版本",
            f"發現新版本 {info.latest_version}（目前 {self.app_version}）。\n\n"
            "按「是」：下載並準備套用\n"
            "按「否」：略過此版本\n"
            "按「取消」：這次先不處理",
            parent=self.root,
        )
        if choice is None:
            self._status("本次更新已取消")
            return
        if choice is False:
            self.updater.mark_skip_version(info.latest_version)
            self._status(f"已略過版本 {info.latest_version}")
            if not silent:
                messagebox.showinfo("更新", f"已略過 {info.latest_version}", parent=self.root)
            return

        self.updater.clear_skipped_version()
        self._status("正在下載更新檔…")
        try:
            downloaded = self.updater.download(info)
        except Exception as e:
            self._status("更新下載失敗")
            messagebox.showerror("更新", f"下載失敗：{e}", parent=self.root)
            return

        messagebox.showinfo(
            "下載完成",
            f"已下載更新檔：\n{downloaded}\n\n下一步可套用更新（舊版會改名為 .bak）。",
            parent=self.root,
        )
        if not messagebox.askyesno(
            "準備套用更新",
            f"已下載 {info.latest_version}（{downloaded.name}）。\n"
            "按「是」後將關閉程式並套用更新。\n"
            f"舊版本將更名為：{Path(sys.executable).name}.bak\n"
            "更新完成後請手動重新開啟程式。\n"
            "（zip 會自動解壓與清理暫存）",
            parent=self.root,
        ):
            self._status("更新檔已保留，尚未套用")
            return

        if not getattr(sys, "frozen", False):
            self._status("目前為開發模式，不自動覆蓋執行中程式")
            messagebox.showinfo(
                "更新已下載",
                f"已下載更新檔：\n{downloaded}\n\n開發模式不執行自動覆蓋，請手動使用此檔。",
                parent=self.root,
            )
            return

        ok = self.updater.apply_update_without_restart(downloaded)
        if not ok:
            self._status("無法啟動更新器")
            messagebox.showerror("更新失敗", "無法啟動更新器，請改為手動更新。", parent=self.root)
            return

        self._status("正在套用更新，程式即將關閉")
        if self.on_after_apply_started:
            self.on_after_apply_started()
            return
        try:
            self.root.after(120, self.root.destroy)
        except Exception:
            pass

