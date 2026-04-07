# Web Fetch Profile 規格

本文件定義「一個資料頁抓取流程」所需的最小設定。

## 資料結構

對應 `web_fetch_profiles.py` 的 `WebFetchProfile`：

- `profile_id`: 唯一識別（例：`little_champion_home`）
- `base_url`: 預設入口網址
- `source_xpath`: 主抓取區塊（通常表格）
- `date_xpath`: 目前日期顯示
- `date_prev_xpath`: 日期前一天按鈕
- `date_next_xpath`: 日期後一天按鈕
- `pre_click_xpath`: 進入資料視窗前的必要按鈕
- `login_input_xpath`: 登入帳號輸入框
- `login_password_xpath`: 登入密碼輸入框（可空）
- `login_confirm_xpath`: 登入確認按鈕

## 流程契約

每個 profile 預設流程：

1. 進入 `base_url`
2. 依登入 XPath 嘗試登入
3. 點擊 `pre_click_xpath`
4. 若指定日期，透過 `date_prev_xpath/date_next_xpath` 調整到目標日期
5. 從 `source_xpath` 抽取資料

## 擴充原則

- 不要在 `app.py` 內硬寫新站台 XPath。
- 新站台應新增 profile（同格式）。
- 同站不同頁可建立多個 profile，避免條件分支過多。

## 後續規劃

- 後續可把 profile 從 Python 搬到 `profiles/*.json`。
- UI 可加入 profile 下拉選單，讓使用者切換資料頁規則。

