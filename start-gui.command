#!/bin/bash
cd "$(dirname "$0")"
if [ ! -d "venv" ]; then
    echo "🔧 首次執行，正在安裝環境..."
    python3 -m venv venv
fi
source venv/bin/activate
# 既有 venv 也檢查，確保新增套件（customtkinter / tkinterdnd2 等）有裝
python3 -c "import customtkinter, tkinterdnd2" 2>/dev/null || {
    echo "🔧 安裝/更新相依套件..."
    pip install -r requirements.txt --quiet
}
# 確認 tkinter（GUI 需要）
python3 -c "import tkinter" 2>/dev/null || {
    echo "⚠ 此 Python 缺 tkinter，GUI 無法啟動。"
    echo "  Mac 修法： brew install python-tk"
    read -p "按任意鍵關閉..." -n1; exit 1
}
# GUI 模式：輸出放桌面 ~/Desktop/潮玩波普ERP/（與打包 exe 一致，行政好找）
export BOPTOYS_GUI=1
python3 app/gui.py
