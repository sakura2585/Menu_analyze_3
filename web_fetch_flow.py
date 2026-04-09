# -*- coding: utf-8 -*-
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
import re
import sys
import time
from typing import Callable

from web_fetch_profiles import WebFetchProfile

_DEFAULT_LOGIN_ACCOUNT = "a0824"

# 給打包器明確的靜態依賴提示，避免 selenium 在部分環境被漏包。
try:  # pragma: no cover
    import selenium  # type: ignore  # noqa: F401
    from selenium.webdriver.common.by import By as _SeleniumBy  # type: ignore  # noqa: F401
except Exception:  # pragma: no cover
    _SeleniumBy = None


def _parse_zh_date(s: str) -> date | None:
    m = re.search(r"(\d{4})年(\d{1,2})月(\d{1,2})日", s or "")
    if not m:
        return None
    try:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    except ValueError:
        return None


@dataclass
class WebFetchRequest:
    url: str
    source_xpath: str
    date_xpath: str
    pre_click_xpath: str
    manual_date: str
    login_account: str
    login_password: str
    profile: WebFetchProfile
    # 若為空字串則使用 profile 內建（小狀元箭頭圖示等）
    date_prev_xpath: str = ""
    date_next_xpath: str = ""
    # True：每列 Tab 分欄時若至少有 4 欄，僅保留前三欄（略過網站備註欄）
    omit_notes_column: bool = True


@dataclass
class WebFetchResult:
    ok: bool
    text: str
    used_date: str
    row_count: int
    error: str = ""


class WebFetchFlow:
    def __init__(self, req: WebFetchRequest, status_cb: Callable[[str], None] | None = None) -> None:
        self.req = req
        self._status_cb = status_cb

    def _status(self, msg: str) -> None:
        if self._status_cb is not None:
            self._status_cb(msg)

    @staticmethod
    def _elem_text_now(elem) -> str:
        txt = (elem.text or "").strip()
        if txt:
            return txt
        for attr in ("innerText", "textContent"):
            try:
                txt = (elem.get_attribute(attr) or "").strip()
            except Exception:
                txt = ""
            if txt:
                return txt
        return ""

    def _create_selenium_driver(self):
        from selenium import webdriver

        last_exc: Exception | None = None
        for ctor, opt_ctor in (
            (webdriver.Edge, webdriver.EdgeOptions),
            (webdriver.Chrome, webdriver.ChromeOptions),
        ):
            try:
                opts = opt_ctor()
                opts.add_argument("--window-size=1280,900")
                return ctor(options=opts)
            except Exception as e:
                last_exc = e
        if last_exc is not None:
            raise last_exc
        raise RuntimeError("無法建立 Selenium 瀏覽器實例。")

    @staticmethod
    def _wait_document_ready(driver, timeout: float = 15.0) -> None:
        t0 = time.time()
        while time.time() - t0 < timeout:
            try:
                if driver.execute_script("return document.readyState") == "complete":
                    return
            except Exception:
                pass
            time.sleep(0.2)

    def _wait_non_empty_text(self, driver, by, locator: str, timeout: int = 120) -> str:
        from selenium.webdriver.support.ui import WebDriverWait

        def _cond(drv):
            try:
                e = drv.find_element(by, locator)
                t = self._elem_text_now(e)
                return t if t else False
            except Exception:
                return False

        return WebDriverWait(driver, timeout).until(_cond)

    @staticmethod
    def _normalize_table_xpath(xpath: str) -> str:
        xp = (xpath or "").strip()
        m_table_idx = re.match(r"^(.*?/table\[\d+\]).*$", xp)
        if m_table_idx:
            return m_table_idx.group(1)
        m_table_plain = re.match(r"^(.*?/table).*$", xp)
        if m_table_plain:
            return m_table_plain.group(1)
        xp = re.sub(r"/tbody/tr\[\d+\]\s*$", "", xp)
        xp = re.sub(r"/tbody/tr\s*$", "", xp)
        xp = re.sub(r"/tr\[\d+\]\s*$", "", xp)
        xp = re.sub(r"/tr\s*$", "", xp)
        return xp

    @staticmethod
    def _drop_last_column_if_four_plus(cells: list[str], do_drop: bool) -> list[str]:
        if do_drop and len(cells) >= 4:
            return cells[:3]
        return cells

    @staticmethod
    def _strip_last_tab_field(line: str, do_drop: bool) -> str:
        if not do_drop or "\t" not in line:
            return line
        parts = line.split("\t")
        if len(parts) < 4:
            return line
        return "\t".join(parts[:3])

    def _table_rows_text(self, elem) -> list[str]:
        drop = bool(getattr(self.req, "omit_notes_column", True))
        rows: list[str] = []
        try:
            tr_list = elem.find_elements("xpath", ".//tr")
        except Exception:
            tr_list = []
        for tr in tr_list:
            cells: list[str] = []
            try:
                cell_elems = tr.find_elements("xpath", "./th|./td")
            except Exception:
                cell_elems = []
            for c in cell_elems:
                t = self._elem_text_now(c)
                cells.append(" ".join(t.split()) if t else "")
            if not any(x.strip() for x in cells):
                continue
            cells = self._drop_last_column_if_four_plus(cells, drop)
            rows.append("\t".join(cells))
        return rows

    def _scroll_table_container_to_load_more(self, driver, table_elem) -> None:
        script = """
const table = arguments[0];
let node = table;
let target = null;
while (node) {
  const sh = node.scrollHeight || 0;
  const ch = node.clientHeight || 0;
  if (sh > ch + 20) { target = node; break; }
  node = node.parentElement;
}
if (!target) {
  window.scrollTo(0, document.body.scrollHeight);
  return 0;
}
const before = target.scrollTop;
target.scrollTop = target.scrollHeight;
return Math.abs(target.scrollTop - before);
"""
        for _ in range(20):
            moved = driver.execute_script(script, table_elem)
            time.sleep(0.18)
            if not moved:
                break

    @staticmethod
    def _extract_block_text_lines(driver, root_xpath: str) -> list[str]:
        script = """
const xp = arguments[0];
const res = document.evaluate(xp, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
const root = res.singleNodeValue;
if (!root) return [];
return (root.innerText || root.textContent || '').split(/\\r?\\n/).map(s => s.trim()).filter(Boolean);
"""
        try:
            lines = driver.execute_script(script, root_xpath)
            if isinstance(lines, list):
                return [str(x).strip() for x in lines if str(x).strip()]
        except Exception:
            pass
        return []

    def _extract_rich_table_text(self, driver, table_xpath: str) -> str:
        from selenium.webdriver.common.by import By

        drop = bool(getattr(self.req, "omit_notes_column", True))
        all_rows: list[str] = []
        seen: set[str] = set()
        base_xpath = self._normalize_table_xpath(table_xpath)
        candidates = [base_xpath]
        m = re.match(r"^(.*?/table)\[(\d+)\]$", base_xpath)
        if m:
            prefix = m.group(1)
            idx = int(m.group(2))
            for i in range(idx + 1, idx + 5):
                candidates.append(f"{prefix}[{i}]")
        for xp in candidates:
            try:
                table_elem = driver.find_element(By.XPATH, xp)
            except Exception:
                continue
            try:
                self._scroll_table_container_to_load_more(driver, table_elem)
            except Exception:
                pass
            rows = self._table_rows_text(table_elem)
            if not rows:
                txt = self._elem_text_now(table_elem)
                rows = [
                    self._strip_last_tab_field(ln.strip(), drop)
                    for ln in txt.splitlines()
                    if ln.strip()
                ]
            for ln in rows:
                ln = self._strip_last_tab_field(ln, drop)
                if ln not in seen:
                    seen.add(ln)
                    all_rows.append(ln)
        if len(all_rows) <= 2:
            for ln in self._extract_block_text_lines(driver, '//*[@id="meal-calc"]'):
                ln = self._strip_last_tab_field(ln.strip(), drop)
                if ln not in seen:
                    seen.add(ln)
                    all_rows.append(ln)
        if all_rows:
            return "\n".join(all_rows).strip()
        try:
            elem = driver.find_element(By.XPATH, table_xpath)
            return self._elem_text_now(elem).strip()
        except Exception:
            return ""

    def _click_xpath(self, driver, xpath: str, timeout: float = 15.0) -> bool:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.support.ui import WebDriverWait

        if not xpath:
            return False
        try:
            elem = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, xpath)))
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
            WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((By.XPATH, xpath)))
            try:
                elem.click()
            except Exception:
                driver.execute_script("arguments[0].click();", elem)
            return True
        except Exception:
            return False

    def _text_of_xpath(self, driver, xpath: str) -> str:
        from selenium.webdriver.common.by import By

        if not xpath:
            return ""
        try:
            e = driver.find_element(By.XPATH, xpath)
            return self._elem_text_now(e).strip()
        except Exception:
            return ""

    def _date_prev_click_xpath(self) -> str:
        s = (self.req.date_prev_xpath or "").strip()
        return s if s else self.req.profile.date_prev_xpath

    def _date_next_click_xpath(self) -> str:
        s = (self.req.date_next_xpath or "").strip()
        return s if s else self.req.profile.date_next_xpath

    def _adjust_date_by_arrows(self, driver) -> bool:
        target = _parse_zh_date(self.req.manual_date)
        if target is None:
            return False
        prev_xp = self._date_prev_click_xpath()
        next_xp = self._date_next_click_xpath()
        if not prev_xp or not next_xp:
            return False
        for _ in range(45):
            cur = _parse_zh_date(self._text_of_xpath(driver, self.req.date_xpath))
            if cur is None:
                time.sleep(0.2)
                continue
            if cur == target:
                return True
            old_text = self._text_of_xpath(driver, self.req.date_xpath)
            if cur > target:
                ok = self._click_xpath(driver, prev_xp, timeout=10)
            else:
                ok = self._click_xpath(driver, next_xp, timeout=10)
            if not ok:
                return False
            t0 = time.time()
            while time.time() - t0 < 4:
                if self._text_of_xpath(driver, self.req.date_xpath) != old_text:
                    break
                time.sleep(0.15)
            time.sleep(0.2)
        return False

    def _try_login(self, driver) -> None:
        from selenium.webdriver.common.by import By
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.support.ui import WebDriverWait

        acct = (self.req.login_account or "").strip() or _DEFAULT_LOGIN_ACCOUNT
        try:
            inp = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.XPATH, self.req.profile.login_input_xpath))
            )
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", inp)
            inp.clear()
            inp.send_keys(acct)
            if (self.req.profile.login_password_xpath or "").strip():
                pw = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, self.req.profile.login_password_xpath))
                )
                pw.clear()
                pw.send_keys(self.req.login_password or "")
            btn = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.XPATH, self.req.profile.login_confirm_xpath))
            )
            try:
                btn.click()
            except Exception:
                driver.execute_script("arguments[0].click();", btn)
        except Exception:
            return

    def run(self) -> WebFetchResult:
        try:
            from selenium.webdriver.common.by import By
        except ImportError:
            if getattr(sys, "frozen", False):
                msg = (
                    "抓取模組缺少 selenium。請使用含 selenium 的新版安裝包，"
                    "或改由原始碼環境執行並先安裝：pip install selenium"
                )
            else:
                msg = "缺少 selenium 套件，請先安裝：pip install selenium"
            return WebFetchResult(False, "", "", 0, msg)

        driver = None
        try:
            self._status("網路抓取：進網頁中…")
            driver = self._create_selenium_driver()
            driver.get(self.req.url)
            self._wait_document_ready(driver, timeout=20)
            self._status("網路抓取：嘗試登入…")
            self._try_login(driver)
            if self.req.pre_click_xpath:
                self._status("網路抓取：等待按鈕並點擊…")
                if not self._click_xpath(driver, self.req.pre_click_xpath, timeout=60):
                    raise ValueError("前置按鈕點擊失敗。")
                time.sleep(0.6)
            if self.req.manual_date:
                self._status("網路抓取：指定日期中…")
                if not self._adjust_date_by_arrows(driver):
                    raise ValueError("指定日期未成功套用（箭頭調整失敗或超過可調整次數）。")
                time.sleep(0.8)
            self._status("網路抓取：獲取資料中…")
            self._wait_non_empty_text(driver, By.XPATH, self.req.source_xpath, timeout=120)
            main_text = self._extract_rich_table_text(driver, self.req.source_xpath)
            if not main_text:
                raise ValueError("表格已定位，但未擷取到可用內容。")
            used_date = ""
            if self.req.date_xpath:
                try:
                    used_date = self._wait_non_empty_text(driver, By.XPATH, self.req.date_xpath, timeout=40)
                except Exception:
                    used_date = ""
            output_text = f"{used_date}\n{main_text}".strip() if used_date else main_text
            row_count = len([ln for ln in output_text.splitlines() if ln.strip()])
            return WebFetchResult(True, output_text, used_date, row_count, "")
        except Exception as e:
            return WebFetchResult(False, "", "", 0, str(e))
        finally:
            try:
                if driver is not None:
                    driver.quit()
            except Exception:
                pass

