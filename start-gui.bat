@echo off
cd /d "%~dp0"
if not exist venv (
    echo Installing environment...
    python -m venv venv
)
call venv\Scripts\activate
python -c "import customtkinter" 2>NUL || (
    echo Installing/updating dependencies...
    pip install -r requirements.txt --quiet
)
python app\gui.py
pause
