@echo off
cd /d "%~dp0"
if not exist venv (
    echo Installing environment...
    python -m venv venv
    call venv\Scripts\activate
    pip install -r requirements.txt --quiet
) else (
    call venv\Scripts\activate
)
python app\gui.py
pause
