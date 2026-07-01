@echo off
cd /d "%~dp0"
if not exist venv (
    echo Installing environment...
    python -m venv venv
)
call venv\Scripts\activate
python -c "import customtkinter, tkinterdnd2" 2>NUL || (
    echo Installing/updating dependencies...
    pip install -r requirements.txt --quiet
)
rem GUI 模式：輸出放桌面 潮玩波普ERP（與打包 exe 一致，行政好找）
set BOPTOYS_GUI=1
python app\gui.py
pause
