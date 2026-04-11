@echo off
:: If not already running in Windows Terminal, relaunch in it
if not defined WT_SESSION (
    where wt >nul 2>&1
    if not errorlevel 1 (
        wt cmd /c "%~f0"
        exit /b
    )
)
chcp 65001 >nul
cd /d "%~dp0"

python --version >nul 2>&1
if errorlevel 1 (
    echo [Setup] Python not found. Attempting auto-install...
    winget --version >nul 2>&1
    if not errorlevel 1 (
        echo [Setup] Installing Python via winget...
        winget install Python.Python.3.11 --silent --accept-package-agreements --accept-source-agreements
    ) else (
        echo [Setup] winget not available. Downloading Python installer...
        powershell -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe' -OutFile '%TEMP%\python_installer.exe'"
        echo [Setup] Running Python installer (please follow the prompts, check "Add Python to PATH")...
        "%TEMP%\python_installer.exe" /passive InstallAllUsers=0 PrependPath=1
    )
    echo [Done] Python installed. Please restart this script.
    pause
    exit /b 0
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
