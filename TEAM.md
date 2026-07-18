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
- **本體程式**：`sales_to_dispatch.py`（前置作業轉檔）、`ann_archive.py`（本機存檔 / 整合 Excel，JOJO 成果歸檔）
- **職責**：
  1. **前置作業** — 把公司銷貨日報表（ERP/會計出）整理成 JOJO 能吃的 `每日配送.xlsx`——去重/續行/退貨相抵/品項彙總。
  2. **本機存檔（JOJO 成果歸檔）** — JOJO 跑完規劃並在 LINE 回傳後，由 ANN 跑 `ann_archive.py` 把 JOJO 雲端複本抓回本機 OneDrive 桌面（當日車輛報表/當日派車單）並產**整合 Excel**。
- **鐵律**：純資料整理不做路線規劃；退貨相抵、續行繼承、同店合併、品項（非鮮奶）進「品項」欄；相抵後 0 瓶且無品項的店跳過。**本機存檔與整合 Excel 是 ANN 的活，JOJO 不再自己上傳 OneDrive。**
- **輸入（前置）**：`OneDrive/桌面/客服AI/1150720銷貨明細(中).xlsx`
- **輸出（前置）**：`OneDrive/桌面/路線規劃/每日配送_<日期>.xlsx` → 交給 JOJO 跑
- **輸出（存檔）**：`OneDrive/桌面/當日車輛報表/<日期>/`、`OneDrive/桌面/當日派車單/<日期>/`（含 `整合報表_<日期>.xlsx`）
- **詳細約束**：見 ANN 的 `SOUL.md`

## 協作資料流
```
銷貨日報表 (客服AI/1150720銷貨明細(中).xlsx)
        │
        ▼
   [ANN ①前置] 去重/續行/退貨相抵/品項彙總 → 每日配送_<日期>.xlsx
        │
        ▼
每日配送.xlsx (路線規劃/)
        │
        ▼
   [JOJO] 路線規劃 + ETA → 產報表/派車單/路線圖
        │
        ├─► 在 LINE 回傳結果給傑夫（流程終點：回傳即停止，等待下一次分配）
        └─► 成果複本留在雲端端點 (/report /dispatch /route_map /workbook)
                    │
                    ▼
            [ANN ②存檔] ann_archive.py 抓回本機 OneDrive 桌面 + 產整合 Excel
                    │
                    ▼
        當日車輛報表/<日期>/   當日派車單/<日期>/ (含 整合報表_<日期>.xlsx)
                    │
                    ▼
           [客服助理] 讀報表答客戶 / (可選) 改單回寫 → 觸發 JOJO 重排
```

## 擴充原則
- 新成員 = 新增一個 Hermes profile + 在本檔加一條成員定義
- 每個成員有自己的 `SOUL.md` 定義職責與邊界，避免職責重疊
- 成員間**只透過共用資料/API 協作**，不直接互改彼此程式碼
- 所有成員共用 `OneDrive/桌面` 下的資料夾慣例
