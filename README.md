# erp-tools-ragic

Ragic ERP 自動化工具，支援銷貨單建立、出貨單拋轉、出庫單拋轉。

---

## 功能

| 功能 | 說明 |
|------|------|
| 新建銷售單 | 解析客戶採購單 Excel，自動建立 Ragic 銷貨單 |
| 建立出貨單 | 從銷貨單批量拋轉建立出貨單（一鍵觸發 Ragic 按鈕） |
| 建立出庫單 | 從出貨單批量拋轉建立出庫單，並自動補填倉庫代碼與庫存編號 |

支援客戶格式：麗嬰（LE）、玩具反斗城（TRU）。設計上可快速擴充新客戶格式。

---

## 快速啟動

| 平台 | 方式 |
|------|------|
| Mac | 雙擊 `start.command` |
| Windows | 雙擊 `start.bat` |

第一次使用請先參閱 [完整說明文件](readme/README.md)。

---

## 使用流程

啟動後出現主選單，依需求選擇功能：

```
請選擇功能：
 » 新建銷售單
   建立出貨單（銷貨單拋轉）
   建立出庫單（出貨單拋轉）
   退出
```

### 新建銷售單

1. 將客戶採購單 Excel 放入 `client_order/<客戶代碼>/`
2. 選擇要處理的檔案（可多選）
3. 填寫訂單資訊（訂單單別、狀態、稅率、運費、備註）
4. 確認後自動上傳至 Ragic，Excel 移入 `done/`

### 建立出貨單

1. 選擇要拋轉的銷貨單（狀態：未出貨 / 預接單 / 已收款未出貨）
2. 確認後批量觸發 Ragic「建立出貨單」按鈕

### 建立出庫單

1. 選擇要拋轉的出貨單
2. 選擇倉庫（預設 TW01 台灣總部）
3. 確認每項商品的庫存編號（唯一選項自動帶入）
4. 確認後批量觸發「建立出庫單」並自動補填倉庫資料

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
