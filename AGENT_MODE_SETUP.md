# Agent Mode - API Key 設定指南

Agent Mode 需要 Anthropic API Key 才能運作。本文檔說明如何設定和管理 API Key。

---

## 📋 快速開始

### 第一次使用

1. **複製環境變數範本**
   ```bash
   cp .env.example .env
   ```

2. **填入你的 API Key**
   ```bash
   # 編輯 .env 檔案
   ANTHROPIC_API_KEY=sk-ant-api03-xxx...
   ```
   
   取得 Key：https://console.anthropic.com/account/keys

3. **進入 Agent Mode**
   ```bash
   python3 app/main.py
   # 選擇「2. Agent Mode - 數據分析」
   ```

---

## 🔄 換 API Key

### 方式 1：在 Agent Mode 中快速更換 ⭐（推薦）

在 Agent Mode 提示符中輸入：
```
重設 key
```
或
```
reset key
```

系統會提示輸入新的 API Key，自動保存到本地。

### 方式 2：編輯 `.env` 檔案

```bash
# 用編輯器打開
vim .env
```

修改：
```
ANTHROPIC_API_KEY=新的key
```

重啟 Agent Mode 生效。

### 方式 3：清除快取重新輸入

```bash
# 刪除本地快取
rm ~/.boptoys-anthropic_key

# 下次進入 Agent Mode 會重新要求輸入
```

---

## 🔐 安全提醒

### ✅ 正確做法
- API Key 存在 `.env` 或 `~/.boptoys-anthropic_key`（本地檔案）
- `.env` 已在 `.gitignore` 中，**不會上傳到 GitHub**
- 每個開發者有自己的 API Key

### ❌ 避免
- 不要把 API Key 粘貼在聊天或代碼中
- 不要提交 `.env` 到版本控制
- 如果不小心洩露，立即輪換 Key

---

## 🐛 Debug

### 檢查 API Key 來源

進入 Agent Mode 時，會顯示：
```
🔍 DEBUG: 開始查詢 API Key...
  └─ 環境變數 ANTHROPIC_API_KEY: [有設定/未設定]
  └─ 本地檔案 ~/.boptoys-anthropic_key: [存在/不存在]
✓ 使用環境變數中的 API Key
```

### 常見問題

| 問題 | 解決方案 |
|------|--------|
| 信用額度不足 | 前往 https://console.anthropic.com/account/billing 充值 |
| API Key 無效 | 確認 Key 未過期，重新複製完整的 Key |
| 讀不到 `.env` | 確認 `.env` 在項目根目錄，重啟 Agent Mode |

---

## 📚 相關資源

- [Anthropic API 文檔](https://docs.anthropic.com)
- [API Key 管理](https://console.anthropic.com/account/keys)
- [計費與額度](https://console.anthropic.com/account/billing)
