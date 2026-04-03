@echo off
chcp 65001 >nul
cd /d "%~dp0"
where python3 >nul 2>&1
if %errorlevel% == 0 (
    python3 app\ragic_upload.py %*
) else (
    python app\ragic_upload.py %*
)
echo.
pause
