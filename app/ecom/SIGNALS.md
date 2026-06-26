# 電商訊號 ↔ 信箱對照（重要）

各平台/狀態的通知信**分散在兩個信箱**，對帳系統需同時讀取兩者。

| 訊號 | 寄件者 | 主要信箱 | 用途 |
|---|---|---|---|
| ShopStore 新訂單 | service@shopstore.tw | **info@boptoys.com.tw** | 開單來源（官網） |
| Pinkoi 新訂單 | notifications@…pinkoi | **info@boptoys.com.tw** | （無明細，人工） |
| 蝦皮 訂單 | info@mail.shopee.tw | **toybebop@gmail.com**（歷史 253）／新單已改寄 info@ | 開單來源（蝦皮） |
| PAYUNi 取貨成功（**領貨狀態**） | 統一金流 PAYUNi | **toybebop@gmail.com** | 標記已領貨／已收款 |
| PAYUNi 退件門市（未領） | 統一金流 PAYUNi | **toybebop@gmail.com** | 未領佐證（僅物流序號，無訂單號） |

## 關鍵注意
- **訂單（info@）與領貨狀態（PAYUNi→toybebop）在不同信箱** → 系統要讀兩個信箱才能完整對帳。
- 蝦皮新訂單已改寄 info@，但 **PAYUNi 仍寄 toybebop**；若要全部統一，需另至 PAYUNi 後台改通知信箱。
- 帳號密碼：`~/.boptoys-info_app_password`、`~/.boptoys-gmail_app_password`（皆應用程式密碼）。

## 領貨通知延遲
PAYUNi「取貨成功」信在實際領貨後 **約 1~2 小時**才送達（實測 1:00~1:45）。系統輪詢再加約 10 分。
→「領貨後約 1~2 小時內可得知」，非即時。

## 對照鍵
- ShopStore/Pinkoi：訂單號（存 Ragic 銷貨單『備註』）。PAYUNi 取貨成功信含「商店自訂單號」可直接對到。
- 蝦皮：買家帳號（存『備註』）+ 日期（+金額）回推。
- PAYUNi 退件信僅含物流序號、無訂單號 → 未領以「逾期未取貨成功」反推為主。
