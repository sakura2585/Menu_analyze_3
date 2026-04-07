# 小狀元訂餐分析工具：重構架構藍圖

## 1. 目標

- 降低 `app.py` 複雜度，避免抓取/分析/匯出互相耦合。
- 讓不同資料頁（不同網址、不同欄位）可配置，不必反覆改核心程式。
- 讓失敗訊息可追蹤：明確知道卡在哪一步與原因。

## 2. 分層設計

### UI 層（`app.py`，後續拆到 `ui/`）

- 畫面、按鈕、欄位輸入、結果展示。
- 呼叫 Application 用例，不直接處理 Selenium 細節。

### Application 層（`application/`，待建立）

- 流程編排（抓取 -> 解析 -> 標籤 -> 匯出）。
- 聚合錯誤、回傳統一結果給 UI。

### Domain 層（`domain/`，待建立）

- 資料模型、規則、解析邏輯（純邏輯，無 UI 與外部依賴）。
- 例如：`ParsedRow`、規則引擎、欄位分析器。

### Infrastructure 層（`infra/`，待建立）

- Selenium 抓取、JSON 儲存、檔案 IO、外部網站互動。

## 3. 目前已落地（Phase 1）

- `web_fetch_profiles.py`
  - 管理站台流程設定與 XPath。
- `web_fetch_settings_store.py`
  - 儲存 UI 輸入的網站/帳號/密碼。
  - 輸出檔：`web_fetch_settings.json`。
- `app.py`
  - 已串接 profile + settings（不再完全寫死於程式常數）。

## 4. 欄位策略（重要）

- 抓取輸出應保留原欄位結構（依頁面 profile 定義）。
- 分析流程不可一套吃全部欄位：
  - 每個分析器只吃對應欄位。
  - 透過管線組合分析器。

## 5. 失敗策略

所有抓取錯誤需標示：

- `step`: 失敗步驟（登入、前置按鈕、日期調整、表格抓取、解析）
- `reason`: 具體原因
- `hint`: 建議修正方向（檢查哪個 XPath/設定）

## 6. 下一步（Phase 2）

1. [x] 建 `web_fetch_flow.py`，將 `app.py` 抓取邏輯搬出。
2. [x] 定義 `WebFetchRequest/WebFetchResult`（`text`, `row_count`, `used_date`, `error`）。
3. [x] `app.py` 改為只收參數 + 呼叫 flow + 顯示結果。
4. [x] 保持現行功能不退化（先重構，不加新功能）。

## 7. Phase 3 之後

- 多 profile（多網址/多頁型）切換。
- 分析器管線化（欄位對應分析器）。
- 測試補齊（parser unit test + fetch smoke test）。

## 8. 現況（2026-04-06）

- 抓取流程已抽離：`web_fetch_flow.py`
- 分析流程已抽離（第一版）：`analyze_flow.py`
  - 含步驟：`sync_page_names_to_hashtag_db` → `parse_all_pages` → `sync_hashtags_from_rows`
  - 列級擴充點：`apply_row_enrichers`（預設空；後續接欄位專用分析器）
- `app.py` 角色已縮小為 UI 協調與結果呈現
- 網路抓取可選「完成後立即分析」（同一事件迴圈內：寫入 → `_analyze`）
- 分析欄位可自訂：`analyze_field_prefs.py` + `analyze_fields.json`；同一套 `_build_tags` 僅掃勾選欄位（見 `docs/ANALYZE_FIELDS.md`）

