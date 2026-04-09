# 打包與搬機說明

本文件提供「本機開發 → 產出 EXE → 搬到新電腦」的最小步驟。

## 1) 開發環境需求

- Python 3.12+
- Windows（目前字型與批次檔流程以 Windows 為主）
- 依賴套件（見 `requirements.txt`）
  - `selenium`
  - `python-docx`
  - `reportlab`

安裝：

```bash
pip install -r requirements.txt
```

## 2) 打包 EXE（PyInstaller）

專案已提供 `order_note.spec`。

```bash
python -m pip install "pyinstaller>=6.0"
python -m PyInstaller --noconfirm order_note.spec
```

輸出：

- `dist/訂餐備註分析.exe`

### 發布建議：完整打包

若要給一般使用者下載，建議使用完整打包（含 `selenium` 與 `certifi` 憑證）：

```bash
python -m pip install -U selenium certifi
python -m PyInstaller --noconfirm order_note_full.spec
```

或直接執行：

- `build_exe_full.bat`（完整打包，發布用）
- `build_exe.bat`（一般打包，本機測試）

## 3) 設定檔位置與搬機重點

本專案採「可攜式資料目錄」：

- 開發模式：資料寫在專案目錄
- EXE 模式：資料寫在 `exe` 同層目錄

主要設定檔（建議一起搬）：

- `input_pages.json`
- `primary_filter_selection*.json`
- `tag_database*.json`
- `tag_profile_prefs.json`
- `web_fetch_settings.json`
- 其他你實際使用中的 `*.json`

## 4) 懶人包內容（給非技術使用者）

建議釋出資料夾至少包含：

- `訂餐備註分析.exe`
- `settings_json/`（上述 JSON 備份）
- `一鍵啟動.bat`
- `一鍵還原設定.bat`
- `請先看我_懶人包說明.txt`

## 5) GitHub 上傳建議

建議上傳：

- 程式碼（`.py`、`docs`、`*.spec`、`requirements.txt`）
- 文件（本檔、架構檔、規格檔）

建議不要上傳：

- `build/`、`dist/` 大檔（可用 Release 附件提供）
- 個人臨時測試輸出（大量 `pdf/png`）

可在 GitHub Release 附上：

- `訂餐備註分析.exe`
- 懶人包 zip（含設定範例/一鍵批次檔）

## 6) Release 內容說明（建議每版都寫）

請直接使用模板檔：

- `docs/RELEASE_NOTES_TEMPLATE.md`

最少可用「超短版 3 行」，建議使用「完整版」以利日後回溯版本差異。

