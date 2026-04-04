# Ragic ERP 自動化工具 — 使用說明

---

## 目錄

- [第一次設定（只需做一次）](#第一次設定只需做一次)
- [日常使用](#日常使用)
  - [功能一：新建銷售單](#功能一新建銷售單)
  - [功能二：建立出貨單](#功能二建立出貨單銷貨單拋轉)
  - [功能三：建立出庫單](#功能三建立出庫單出貨單拋轉)
- [常見問題](#常見問題)

---

## 第一次設定（只需做一次）

### 步驟 1：安裝 Python

先確認電腦有沒有安裝 Python。

**Mac：** 打開「終端機」，輸入：
```
python3 --version
```

**Windows：** 打開「命令提示字元」，輸入：
```
python --version
```

如果出現版本號（如 `Python 3.11.0`）代表已安裝，跳到步驟 2。

如果出現錯誤，請至 https://www.python.org/downloads/ 下載安裝，安裝時勾選 **「Add Python to PATH」**。

---

### 步驟 2：下載程式

在 GitHub 頁面點右上角的綠色「**Code**」按鈕 → 選「**Download ZIP**」。

解壓縮後，將資料夾放到**桌面**，資料夾名稱保持 `erp-tools-ragic`。

---

### 步驟 3：安裝必要套件

**Mac：** 打開「終端機」，貼上以下指令後按 Enter：
```
cd ~/Desktop/erp-tools-ragic && python3 -m venv venv && source venv/bin/activate && pip install -r requirements.txt
```

**Windows：** 打開「命令提示字元」，貼上以下指令後按 Enter：
```
cd %USERPROFILE%\Desktop\erp-tools-ragic && python -m venv venv && venv\Scripts\activate && pip install -r requirements.txt
```

看到一堆文字跑完，最後沒有 `ERROR` 代表成功。

---

### 步驟 4：設定 .env 檔案

**Mac：** 終端機執行：
```
cp .env.example .env
```

**Windows：** 命令提示字元執行：
```
copy .env.example .env
```

一般情況下這個檔案不需要修改。

---

### 步驟 5：Mac 額外設定（Windows 跳過）

讓 `start.command` 可以雙擊執行，在終端機執行：
```
chmod +x ~/Desktop/erp-tools-ragic/start.command
```

---

### 步驟 6：設定 Ragic API 金鑰

第一次啟動程式時，程式會自動要求輸入 API 金鑰：

```
請輸入 Ragic API Key：
```

貼上金鑰後按 Enter，金鑰會自動儲存，**下次不需要再輸入**。

> API 金鑰請向 Woody 索取。

---

## 日常使用

雙擊啟動程式：

| 平台 | 方式 |
|------|------|
| Mac | 雙擊桌面的 `start.command` |
| Windows | 雙擊桌面的 `start.bat` |

> Windows 若出現「Windows 已保護您的電腦」→ 點「更多資訊」→「仍要執行」。

啟動後用 **↑↓ 方向鍵** 選擇功能，按 **Enter** 確認：

```
請選擇功能：
 » 新建銷售單
   建立出貨單（銷貨單拋轉）
   建立出庫單（出貨單拋轉）
   退出
```

> 任何步驟選錯，可以選「← 返回」退回上一步。執行前程式也會再次確認。

---

### 功能一：新建銷售單

**適用情境：** 收到客戶的採購單 Excel，要建立 Ragic 銷貨單。

**操作步驟：**

**1. 把採購單 Excel 放進對應資料夾**

| 客戶 | 放到這個資料夾 |
|------|--------------|
| 麗嬰國際（LE） | `client_order/LE/` |
| 玩具反斗城（TRU） | `client_order/TRU/` |

**2. 啟動程式，選「新建銷售單」**

**3. 用空白鍵勾選要處理的採購單，按 Enter 確認**

```
請選擇要處理的採購單：
 ○ LE/0829-T221-0903到店.xlsx
 ○ TRU/578029潮玩波普新品.xlsx
```

**4. 依照提示回答訂單資訊**

程式會問幾個問題（訂單單別、狀態、稅率、運費、備註），用方向鍵選擇或直接輸入。

**5. 確認後自動上傳**

上傳成功後，Excel 會自動移入 `done/` 資料夾。登入 Ragic 確認銷貨單已建立。

---

### 功能二：建立出貨單（銷貨單拋轉）

**適用情境：** 銷貨單已建立，要一鍵在 Ragic 建立對應的出貨單。

**操作步驟：**

1. 選「建立出貨單（銷貨單拋轉）」
2. 程式自動列出狀態為「未出貨 / 預接單 / 已收款未出貨」的銷貨單
3. 用**空白鍵**勾選要建立出貨單的銷貨單，按 **Enter** 確認
4. 確認摘要後按 **Y** 執行
5. 至 Ragic 出貨單頁面確認

> 已建立過出貨單的銷貨單，Ragic 會自動擋掉重複建立，不用擔心。

---

### 功能三：建立出庫單（出貨單拋轉）

**適用情境：** 出貨單已建立，要一鍵在 Ragic 建立出庫單，並自動填入倉庫資訊。

**操作步驟：**

1. 選「建立出庫單（出貨單拋轉）」
2. 用**空白鍵**勾選要建立出庫單的出貨單，按 **Enter** 確認
3. 選擇倉庫（TW01 台灣總部預設在最上方）
4. 確認每個商品的庫存編號（只有一個選項時程式自動帶入，不用選）
5. 確認摘要後按 **Y** 執行
6. 至 Ragic 出庫單頁面確認，倉庫代碼與庫存編號已自動填入

---

## 常見問題

**Q：Mac 雙擊 start.command 沒有反應**

A：需先在終端機執行一次步驟 5 的指令（`chmod +x ...`）。

---

**Q：Windows 雙擊 start.bat 出現亂碼或錯誤**

A：確認已安裝 Python，且步驟 3 的指令執行成功（沒有 ERROR）。

---

**Q：程式說「找不到客戶」**

A：輸入客戶代碼或名稱的一部分搜尋，例如輸入「TRU」或「麗嬰」。若確認客戶存在但找不到，請聯絡 Woody 確認 Ragic 客戶資料是否有建檔。

---

**Q：程式說「找不到條碼對應商品」**

A：該商品在 Ragic 商品單價管理尚未建檔，會自動跳過。請請 Woody 補建資料後重新執行。

---

**Q：想重新上傳已處理過的同一個 Excel**

A：找到程式資料夾內的 `upload_log.json`，用記事本打開，刪除該筆記錄後存檔，重新執行即可。或直接刪除整個 `upload_log.json` 清空所有記錄。

---

**Q：Windows 出現「無法載入，因為這個系統上已停用指令碼執行」**

A：以**系統管理員**身分執行 PowerShell，輸入：
```
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```
按 Enter，輸入 `Y` 確認。

---

**Q：Mac 程式跑完視窗直接關掉，看不到結果**

A：打開終端機 → 偏好設定（Cmd+,）→ 描述檔 → Shell → 「Shell 結束時」改為「不要關閉視窗」。

---

**Q：API 金鑰要重新設定**

A：在終端機（Mac）或命令提示字元（Windows）執行：
```
# Mac
python3 app/ragic_upload.py --reset-key

# Windows
python app\ragic_upload.py --reset-key
```
