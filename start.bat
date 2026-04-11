@echo off
chcp 65001 >nul
cd /d "%~dp0"

if not exist "venv\" (
    echo [Setup] First run - installing environment, please wait...
    python -m venv venv
    if errorlevel 1 (
        echo [Error] Python not found. Please install Python 3.8+ first.
        pause
        exit /b 1
    )
    call venv\Scripts\activate.bat
    pip install -r requirements.txt --quiet
    echo [Done] Environment ready.
) else (
    call venv\Scripts\activate.bat
)

echo [Tip] For best experience, use Windows Terminal instead of CMD.
python app\ragic_upload.py %*
echo.
pause
