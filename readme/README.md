# erp-tools-ragic — Ragic 銷貨單自動上傳工具

自動將客戶採購單 Excel 解析並上傳至 Ragic 建立銷貨單，支援麗嬰（LE）、玩具反斗城（TRU）格式。

---

## 目錄結構

```
erp-tools-ragic/
├── start.command         ← Mac 雙擊啟動
├── start.bat             ← Windows 雙擊啟動
├── client_order/         ← 採購單放這裡
│   ├── LE/               ← 麗嬰採購單
│   │   └── done/         ← 上傳完成後自動移入（程式自動建立）
│   └── TRU/              ← 玩具反斗城採購單
│       └── done/
├── app/
│   ├── ragic_upload.py   ← 主程式
│   └── parsers/          ← 各客戶格式解析器
├── .env                  ← 設定檔（需自行建立，見下方）
├── .env.example          ← 設定範本
├── requirements.txt      ← Python 套件清單
└── readme/README.md      ← 本說明文件
```

---

## 環境設定（第一次使用，只需做一次）

### 1. 確認 Python 版本（需 3.8 以上）

```bash
# Mac / Linux
python3 --version

# Windows
python --version
```

如果沒有 Python，請至 https://www.python.org/downloads/ 下載安裝。

---

### 2. 建立虛擬環境（venv）

**Mac / Linux：**
```bash
cd ~/Desktop/erp-tools-ragic
python3 -m venv venv
source venv/bin/activate
```

**Windows（命令提示字元）：**
```cmd
cd %USERPROFILE%\Desktop\erp-tools-ragic
python -m venv venv
venv\Scripts\activate
```

**Windows（PowerShell）：**
```powershell
cd $env:USERPROFILE\Desktop\erp-tools-ragic
python -m venv venv
venv\Scripts\Activate.ps1
```

> 啟動成功後，命令列前方會出現 `(venv)` 字樣。

---

### 3. 安裝套件

```bash
pip install -r requirements.txt
```

---

### 4. 設定 `.env`

**Mac：**
```bash
cp .env.example .env
```

**Windows：**
```cmd
copy .env.example .env
```

一般情況下 `.env` 內容不需要修改。若 Ragic 帳號或表單有異動，再開啟 `.env` 編輯。

---

### 5. 設定 API 金鑰

**Mac — 讓 start.command 可執行（只需一次）：**
```bash
chmod +x ~/Desktop/erp-tools-ragic/start.command
```

第一次執行程式時，若未偵測到 API 金鑰，程式會自動提示輸入。貼上後**自動儲存到本機**（`~/.boptoys-ai_key`），下次不需再輸入。

> 若需重設金鑰，執行：
> ```bash
> python3 app/ragic_upload.py --reset-key   # Mac
> python app\ragic_upload.py --reset-key    # Windows
> ```

---

## 執行方式

### 互動模式（日常使用）

**Mac — 雙擊啟動：**
直接在 Finder 雙擊 `start.command`。

**Windows — 雙擊啟動：**
直接在檔案總管雙擊 `start.bat`。

> 若 Windows 出現「Windows 已保護您的電腦」警告，點「更多資訊」→「仍要執行」。

**Mac — 終端機：**
```bash
source venv/bin/activate
python3 app/ragic_upload.py
```

**Windows — 命令提示字元：**
```cmd
venv\Scripts\activate
python app\ragic_upload.py
```

---

### 指定單一檔案

```bash
# Mac
python3 app/ragic_upload.py "client_order/LE/0324T221.xlsx"

# Windows
python app\ragic_upload.py "client_order\LE\0324T221.xlsx"
```

### Dry-run 模式（測試，不實際上傳）

```bash
# Mac
python3 app/ragic_upload.py --dry-run

# Windows
python app\ragic_upload.py --dry-run
```

---

## 日常使用流程

### 步驟 1：放入採購單

將客戶的採購單 Excel 放入對應資料夾：

| 客戶 | 資料夾 |
|---|---|
| 麗嬰國際（LE） | `client_order/LE/` |
| 玩具反斗城（TRU） | `client_order/TRU/` |

### 步驟 2：啟動程式

按上方「執行方式」操作。

### 步驟 3：勾選要處理的檔案

```
請選擇要處理的採購單（空白鍵勾選，Enter 確認）：
 ○ LE/0829-T221-0903到店.xlsx
 ○ TRU/578029潮玩波普新品.xlsx
```

### 步驟 4：依照提示回答訂單資訊

1. 訂單單別
2. 訂單狀態
3. 稅率
4. 運費
5. 備註（選填）
6. 最終確認

### 步驟 5：完成

上傳成功後 Excel 自動移至 `done/` 資料夾，登入 Ragic 確認銷貨單已建立。

---

## 如何新增客戶格式

每個客戶的採購單格式不同，透過以下 3 個步驟即可支援新客戶：

### 步驟 1：新增 Parser

在 `app/parsers/` 建立 `xxx_parser.py`（xxx 為客戶代碼小寫），繼承 `BaseParser` 並實作 `parse()` 方法：

```python
# app/parsers/abc_parser.py
from .base import BaseParser, Order, OrderItem

class AbcParser(BaseParser):
    def parse(self) -> list[Order]:
        # 讀取 self.filepath 的 Excel，解析出訂單列表
        orders = []
        # ... 解析邏輯（參考 le_parser.py 或 tru_parser.py）...
        return orders
```

### 步驟 2：登記到 PARSERS

開啟 `app/parsers/__init__.py`，加入新的 parser：

```python
from .le_parser  import LeParser
from .tru_parser import TruParser
from .abc_parser import AbcParser   # ← 新增

PARSERS = {
    "LE":  LeParser,
    "TRU": TruParser,
    "ABC": AbcParser,               # ← 新增（key 必須大寫）
}
```

### 步驟 3：建立採購單資料夾

在 `client_order/` 建立與 key 同名的資料夾，並放入 `.gitkeep` 讓 git 追蹤空資料夾：

```
client_order/
└── ABC/
    └── .gitkeep
```

完成後，程式啟動時會自動掃描並用對應的 parser 處理。處理完的訂單會自動歸檔至 `ABC/done/`。

---

## 常見問題

**Q：Mac 雙擊 start.command 沒有反應**
A：需先執行 `chmod +x ~/Desktop/erp-tools-ragic/start.command`（步驟 5）。

**Q：Windows 雙擊 start.bat 出現亂碼**
A：確認系統已安裝 Python，且已執行 `pip install -r requirements.txt`。

**Q：找不到客戶**
A：程式會讓你輸入關鍵字搜尋，打客戶代碼或名稱片段即可。

**Q：找不到條碼對應商品**
A：該商品會顯示警告並跳過，需在 Ragic 商品單價管理補充資料後重新執行。

**Q：想重新上傳已處理過的檔案**
A：刪除 `upload_log.json` 中對應記錄（或整個檔案清空所有記錄）。

**Q：Windows 執行 PowerShell 出現「無法載入」錯誤**
A：以系統管理員身分執行 PowerShell，輸入：
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

**Q：程式結束後 Terminal 視窗直接關掉（Mac）**
A：Terminal → 偏好設定 → 描述檔 → Shell → 「Shell 結束時」改為「不要關閉視窗」。
