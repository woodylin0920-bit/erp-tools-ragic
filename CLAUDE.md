# CLAUDE.md — agent 開機說明

給進到這個資料夾的 AI agent（Claude Code / Codex 等）的快速上手文件。
目標：agent 進來就能**操作這套 ERP 工具、查資料、開單、除錯、修跑版**；沒有 AI 時人也能 `./start.command` 手動跑。

> Codex 讀的是同內容的 `AGENTS.md`（本檔的鏡像）。改動請兩邊同步。

---

## 這是什麼

潮玩波普（Boptoys）的 Ragic ERP 自動化工具（純 Python + Ragic API）。
主程式 `app/ragic_upload.py` 是互動式選單（questionary）。功能：

- **新建銷售單**：讀客戶訂購 Excel（TRU/LE…）→ 自動開 Ragic 銷貨單
- **建立出貨單 / 出庫單**：銷貨單→出貨單→出庫單拋轉（出庫才扣庫存）
- **匯出庫存報表**：客戶現貨表、月度庫存金額表（`app/export_inventory_value.py`）
- **在途查詢**：採購單未到貨（`app/in_transit_query.py`）
- **Agent mode**：工具內建的唯讀 AI 查詢（`app/ai_assistant.py`，走 Anthropic API，只能查不能寫）

## 怎麼跑
- 人手動：`./start.command`（第一次自動建 venv、裝 `requirements.txt`）。互動選單第一個提示輸入 `debug` = 預覽不寫入。
- 指定檔：`python3 app/ragic_upload.py <檔路徑> --dry-run`（檔的**父資料夾名 = 客戶代碼**，如 `client_order/TRU/xxx.xlsx`，否則選不到 parser）。
- agent 注意：questionary 選單**需要真人 TTY，agent 無法代點**會卡住。要完整代開單請走 Ragic API 直接做，或先加非互動模式。

## Ragic 存取（關鍵）
- 帳號 `toybebop` @ `https://ap12.ragic.com`，API key 在 `~/.boptoys-ai_key`（已是 Base64，直接當 `Authorization: Basic <key>`）。失效時所有表回 `status=ERROR / code=106`。
- ⚠️ 這把 key 能**讀寫正式 ERP**。寫入動作（開單/改庫存/建客戶）一律**先把內容給使用者確認再送出**。
- 詳細 sheet IDs、庫存流向、`history=true` 查異動：見 `~/.claude/projects/.../memory/ragic-data-model.md`（這是個人記憶，換機/換 Codex 讀不到，核心已濃縮於下）。

### 常用 sheet
- 倉庫庫存 `ragicinventory/20008`（數量欄 `3001107`，倉庫代碼/商品編號/數量/種類/規格）
- 出庫單 `ragicinventory/20009`、入庫單 `20005`、調撥 `20010`
- 銷貨單 `ragicsales-order-management/20001`、客戶主檔 `ragicsales-order-management/20004`、商品單價 `.../20006`
- 採購單 `ragicpurchasing/20003`

### Ragic 寫入 API（很重要）
- POST/PATCH `/{sheet}?api&doLinkLoad=true`，body 的 key **只能是數字欄位 ID（CID），不認中文名稱**（用名稱回 `INVALID code=202 Field id ... not found`）。
- **取 CID 的方法**：任何讀取網址加 `&naming=EID` → 回傳改用 CID 當 key，跟名稱版對照即得「名稱↔CID」表。
- 客戶主檔 20004 CID：客戶名稱`3000479`(必填)、客戶簡稱`3001873`、客戶負責人`3000480`、主要聯絡窗口`3001449`、窗口手機`3000909`、電話號碼`3000483`、送貨地址`3000903`(可寫)、送貨完整地址`3000904`(公式勿寫)、備註`3000913`。客戶編號`3000666` **自動產生勿填**（= _ragicId+1）。
- 客戶命名慣例：TRU 門市 = `TRU-<門市店號>`（如 `TRU-4463`），簡稱=門市名，負責人=Woody。

## 客戶格式眉角（parser 在 `app/parsers/`）
- **TRU**（`tru_parser.py`）：一檔多門市，每門市一張單。用 **PCS（個）下單，我們發中盒** → 開單時規格選中盒。欄位已改成**讀標題自動定位**（條碼/單價/PO號碼/門市區起點），TRU 改版跑版能自動對上。同店多品項用整檔主要 PO 歸成一張單。
- **PCS 非整中盒**（如 98÷8）通常是客戶填錯：不擋單，`resolve_items` 會標 `box_note` 寫進訂單內部備注提醒人工確認。
- **LE**（`le_parser.py`）：一檔多分頁（總表+每門市一個工作表），店號格式 `AD227`，PO 格式 `PO-xxxxxxx`。LE 不帶單價，價格開單時從 Ragic 商品表帶（所以解析出價=0 是正常）。
- **Template**（`template_parser.py`）：通用報價模板。
- `le_parser.py` / `template_parser.py` 仍有寫死欄位 index，是潛在跑版風險（TRU 已改健壯，這兩個尚未）。

## 開發約定
- 改 parser 後**務必用真實舊檔回歸測試**（今天的跑版 bug 就是這樣抓到的）。對照「改動前 vs 改動後」訂單數/總量是否一致。歷史檔在 `~/Desktop/雜物/erp備份/.../client_order/{TRU,LE}/done/`。
- 不要把真實客戶訂單檔 commit 進 repo（含客戶/價格資料）。
- 寫入正式 ERP 的動作先給使用者確認。
- 商品/門市的數字統一用 PCS 思考，報客戶/出貨才換中盒。

## 方向（使用者期望）
在電腦前開 Claude/Codex 進這個資料夾 → 用講的操作工具（查庫存、開單、建客戶）、又能即時除錯/修跑版；沒有 AI 時 `./start.command` 人工照樣能跑。三層：手動選單 / AI 代操作 / （待加）非互動自動化。
