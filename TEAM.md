# 物流內勤團隊 (Logistics Back-office Team)

本文件定義傑夫物流事業的**內勤自動化團隊編制**：由各個專職 Agent 組成，每個 Agent 是獨立的 Hermes profile（實體），彼此透過共用資料（Excel / 報表 / API）協作。

## 團隊成員

### 1. JOJO — 路線規劃 Agent ✅ 已實體化
- **實體**：Hermes profile `jojo`（`SOUL.md` 在 `~/AppData/Local/hermes/profiles/jojo/`）
- **本體程式**：`logistics_agent`（GitHub `qazwsx00500-dot/milk-logistics-bot`）
- **部署**：Render `https://milk-logistics-bot.onrender.com` + LINE 頻道
- **職責**：讀每日配送 Excel → 每台車最優停靠順序 → 產路線總表/派車單/路線圖 + 每站 ETA
- **鐵律**：純路線規劃不做載貨量估算；每車從各自出發點回該點（倉庫=台中大雅101-1號）；下貨=瓶數×15秒折進 ETA；資料走 OneDrive
- **輸入**：`OneDrive/桌面/路線規劃/每日配送.xlsx`
- **輸出**：`OneDrive/桌面/當日車輛報表/<日期>/`、`OneDrive/桌面/當日派車單/<日期>/`
- **詳細約束**：見 JOJO 的 `SOUL.md`

### 2. 客服助理 (Customer Service Assistant) — 🚧 規劃中
- **預計實體**：Hermes profile（待建，建議命名 `jojo-cs` 或獨立名）
- **職責**：面向客戶（便利商店/茶飲店/零售點）的問答與改單受理
- **與 JOJO 連動方式**（待確認，3 選 1）：
  - **(A) 只讀查詢型**：讀 JOJO 產出的報表/ETA，回答「幾點到、送幾瓶」，**不動 Excel**（最簡單、最安全）
  - **(B) 讀＋改單回寫型**：收客戶改單/加瓶/取消 → 回寫 `每日配送.xlsx` → 觸發 JOJO 重排路線
  - **(C) 工單分流型**：先分類客戶問題，物流類轉 JOJO、其他（退換貨/帳務）自己答
- **共用介面**：讀取 JOJO 的 `當日車輛報表/<日期>/route_report.html` 與 `當日派車單/<日期>/dispatch.csv` 取得 ETA/瓶數；(B)(C) 模式則需要寫入 `每日配送.xlsx` 並呼叫 JOJO 重排
- **狀態**：2026-07-18 啟動，模式待傑夫確認

## 協作資料流
```
每日配送.xlsx (傑夫編輯)
        │
        ▼
   [JOJO] 路線規劃 + ETA
        │
        ├─► 當日車輛報表/<日期>/   (路線總表 + 路線圖)
        ├─► 當日派車單/<日期>/     (每台車一份，司機用)
        │
        ▼
   [客服助理] 讀報表答客戶 / (可選) 改單回寫 → 觸發 JOJO 重排
```

## 擴充原則
- 新成員 = 新增一個 Hermes profile + 在本檔加一條成員定義
- 每個成員有自己的 `SOUL.md` 定義職責與邊界，避免職責重疊
- 成員間**只透過共用資料/API 協作**，不直接互改彼此程式碼
- 所有成員共用 `OneDrive/桌面` 下的資料夾慣例
