#!/bin/bash
cd "$(dirname "$0")"
if [ ! -d "venv" ]; then
    echo "🔧 首次執行，正在安裝環境..."
    python3 -m venv venv
    source venv/bin/activate
    pip install -r requirements.txt --quiet
else
    source venv/bin/activate
fi
# 確認 tkinter（GUI 需要）
python3 -c "import tkinter" 2>/dev/null || {
    echo "⚠ 此 Python 缺 tkinter，GUI 無法啟動。"
    echo "  Mac 修法： brew install python-tk"
    read -p "按任意鍵關閉..." -n1; exit 1
}
python3 app/gui.py
