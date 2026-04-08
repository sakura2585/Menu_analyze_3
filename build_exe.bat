@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo [1/2] 安裝 PyInstaller、python-docx（匯出 Word 用）…
python -m pip install -q "pyinstaller>=6.0" "python-docx>=1.1.0"
if errorlevel 1 (
    echo 請確認已安裝 Python 並加入 PATH。
    pause
    exit /b 1
)

echo [2/2] 打包為單一 EXE…
python -m PyInstaller --noconfirm order_note.spec
if errorlevel 1 (
    echo 打包失敗。
    pause
    exit /b 1
)

echo.
echo 完成。執行檔： dist\MenuAnalyze.exe
echo 請將 .exe 複製到欲使用的資料夾；設定與標籤庫會寫在 .exe 同一層目錄。
pause
