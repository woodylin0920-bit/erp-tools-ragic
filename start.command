#!/bin/bash
cd "$(dirname "$0")"
python3 app/ragic_upload.py "$@"
echo ""
read -p "按任意鍵關閉視窗..." -n1
osascript -e 'tell application "Terminal" to close front window'
