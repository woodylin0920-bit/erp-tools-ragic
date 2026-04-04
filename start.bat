@echo off
chcp 65001 >nul
cd /d "%~dp0"

:: 第一次執行自動建立虛擬環境並安裝套件
if not exist "venv\" (
    echo 🔧 首次執行，正在安裝環境（約需 1 分鐘）...
    python -m venv venv
    call venv\Scripts\activate.bat
    pip install -r requirements.txt --quiet
    echo ✅ 環境安裝完成
) else (
    call venv\Scripts\activate.bat
)

python app\ragic_upload.py %*
echo.
pause
