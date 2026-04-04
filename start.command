#!/bin/bash
cd "$(dirname "$0")"

# 第一次執行自動建立虛擬環境並安裝套件
if [ ! -d "venv" ]; then
    echo "🔧 首次執行，正在安裝環境（約需 1 分鐘）..."
    python3 -m venv venv
    source venv/bin/activate
    pip install -r requirements.txt --quiet
    echo "✅ 環境安裝完成"
else
    source venv/bin/activate
fi

python3 app/ragic_upload.py "$@"
echo ""
read -p "按任意鍵關閉視窗..." -n1
osascript -e 'tell application "Terminal" to close front window'
