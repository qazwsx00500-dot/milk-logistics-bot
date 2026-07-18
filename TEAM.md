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

### 2. ANN — 客服業助 / 銷貨明細整理 Agent ✅ 已實體化
- **實體**：Hermes profile `ann`（`SOUL.md` 在 `~/AppData/Local/hermes/profiles/ann/`）
- **本體程式**：`sales_to_dispatch.py`（`logistics_agent` 專案內，ANN 的轉檔工具）
- **職責**：把公司銷貨日報表（ERP/會計出）整理成 JOJO 能吃的 `每日配送.xlsx`——去重/續行/退貨相抵/品項彙總。是 JOJO 的「前置作業」成員。
- **鐵律**：純資料整理不做路線規劃；退貨相抵、續行繼承、同店合併、品項（非鮮奶）進「品項」欄；相抵後 0 瓶且無品項的店跳過。
- **輸入**：`OneDrive/桌面/客服AI/1150720銷貨明細(中).xlsx`（銷貨日報表）
- **輸出**：`OneDrive/桌面/路線規劃/每日配送_<日期>.xlsx` → 交給 JOJO 跑
- **詳細規則**：見 ANN 的 `SOUL.md`

## 協作資料流
```
銷貨日報表 (客服AI/1150720銷貨明細(中).xlsx)
        │
        ▼
   [ANN] 去重/續行/退貨相抵/品項彙總 → 每日配送_<日期>.xlsx
        │
        ▼
每日配送.xlsx (路線規劃/)
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
