@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo [1/3] 安裝 PyInstaller…
python -m pip install -q "pyinstaller>=6.0"
if errorlevel 1 (
    echo 請確認已安裝 Python 並加入 PATH。
    pause
    exit /b 1
)

echo [2/3] 安裝完整打包依賴（selenium, certifi）…
python -m pip install -q -U selenium certifi
if errorlevel 1 (
    echo 安裝 selenium/certifi 失敗。
    pause
    exit /b 1
)

echo [3/3] 完整打包為單一 EXE…
python -m PyInstaller --noconfirm order_note_full.spec
if errorlevel 1 (
    echo 打包失敗。
    pause
    exit /b 1
)

echo.
echo 完成。執行檔： dist\MenuAnalyze.exe
echo 這是「完整打包」版本，建議用於發版提供使用者下載。
pause
