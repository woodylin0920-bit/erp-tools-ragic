#!/bin/bash
# 打包桌面 GUI 成單一執行檔（Mac .app / Linux 執行檔）。Windows 用 GitHub Actions 建 .exe。
set -e
cd "$(dirname "$0")/.."
[ -d venv ] && source venv/bin/activate
pip install -r requirements.txt pyinstaller --quiet
rm -rf build dist
pyinstaller --noconfirm --windowed --name BoptoysERP \
  --collect-data customtkinter \
  --collect-all tkinterdnd2 \
  --add-data "templates:templates" \
  --paths app \
  --hidden-import sample_core --hidden-import outbound_core \
  --hidden-import export_core --hidden-import sales_core --hidden-import ragic_upload \
  --collect-submodules parsers --collect-submodules ecom \
  app/gui.py
echo "完成 → dist/BoptoysERP.app（Mac）"
