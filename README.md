# erp-tools-ragic

自動將客戶採購單 Excel 解析，並上傳至 [Ragic](https://www.ragic.com/) 建立銷貨單。

支援客戶格式：麗嬰（LE）、玩具反斗城（TRU）。設計上可快速擴充新客戶格式。

---

## 快速啟動

| 平台 | 方式 |
|------|------|
| Mac | 雙擊 `start.command` |
| Windows | 雙擊 `start.bat` |

第一次使用請先參閱 [完整說明文件](readme/README.md)。

---

## 目錄結構

```
erp-tools-ragic/
├── start.command       ← Mac 啟動
├── start.bat           ← Windows 啟動
├── client_order/       ← 採購單放這裡（依客戶分資料夾）
├── app/
│   ├── ragic_upload.py ← 主程式
│   └── parsers/        ← 各客戶格式解析器
├── .env.example        ← 設定範本
└── requirements.txt    ← Python 套件
```

---

## 環境需求

- Python 3.8+
- 套件：`pip install -r requirements.txt`
- Ragic API 金鑰（第一次執行時程式會自動提示輸入）

---

## 新增客戶格式

1. 在 `app/parsers/` 新增 `xxx_parser.py`（繼承 `BaseParser`，實作 `parse()`）
2. 在 `app/parsers/__init__.py` 的 `PARSERS` 字典登記
3. 建立 `client_order/XXX/` 資料夾

詳細步驟見 [說明文件 → 如何新增客戶格式](readme/README.md#如何新增客戶格式)。
