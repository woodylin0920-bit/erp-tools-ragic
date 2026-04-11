@echo off
chcp 65001 >nul
cd /d "%~dp0"

python --version >nul 2>&1
if errorlevel 1 (
    echo [Error] Python not found.
    echo Please install Python 3.8+ from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

if not exist "venv\" (
    echo [Setup] First run - installing environment, please wait...
    python -m venv venv
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
