"""
excel_normalizer.py — 把「任意來源 Excel」轉成標準每日配送表

標準格式 (4 欄)：車號 / 店家名稱 / 店家地址 / 瓶數
來源可能：
  - 欄位名不同 (車輛/店名/地址/數量...) → 寬鬆對應
  - 沒有「車號」欄 → 留空，由 caller 決定怎麼分車
  - 地址欄可能夾雜「-隨貨附發票」之類後綴 → 保留原值(地理編碼時再清理)
  - 多餘列/空列 → 跳過

產出：DataFame(標準4欄) + 存檔 每日配送_YYYYMMDD.xlsx 到 out_dir
"""

import os
import re
import datetime

try:
    import openpyxl
except ImportError:
    openpyxl = None

# 欄位別名對應（小寫去空白後比對）
_VEH = ["車號", "車輛", "路線編號", "路線", "route", "vehicle", "車", "車牌"]
_NAME = ["店家名稱", "名稱", "店名", "客戶名稱", "店家", "name", "店"]
_ADDR = ["店家地址", "地址", "客戶地址", "address", "addr", "位置"]
_QTY = ["瓶數", "數量", "箱數", "瓶量", "qty", "bottles", "count", "件數", "瓶"]

# 台灣縣市關鍵字（用於無車號時的地理分車）
_REGION_KEYWORDS = [
    ("台北", "台北"), ("新北", "新北"), ("基隆", "基隆"), ("桃園", "桃園"),
    ("新竹", "新竹"), ("苗栗", "苗栗"), ("台中", "台中"), ("彰化", "彰化"),
    ("南投", "南投"), ("雲林", "雲林"), ("嘉義", "嘉義"), ("台南", "台南"),
    ("高雄", "高雄"), ("屏東", "屏東"), ("宜蘭", "宜蘭"), ("花蓮", "花蓮"),
    ("台東", "台東"), ("澎湖", "澎湖"),
]


def _norm(h):
    return (h or "").strip().lower().replace(" ", "").replace("　", "")


def _match_col(headers_norm, aliases):
    for a in aliases:
        if _norm(a) in headers_norm:
            return headers_norm[_norm(a)]
    return None


def _clean_int(v):
    try:
        if v is None or v == "":
            return 0
        if isinstance(v, (int, float)):
            return int(v)
        # 去掉逗號/空白後取數字
        m = re.search(r"\d+", str(v))
        return int(m.group()) if m else 0
    except Exception:
        return 0


def normalize_excel(path, out_dir, default_vehicle="車01", date_str=None):
    """
    讀任意來源 Excel，轉成標準 4 欄 DataFrame。
    回傳 (rows, skipped, out_path)
      rows: [(車號, 店家名稱, 店家地址, 瓶數), ...]  (車號可能為 "")
      skipped: [(原始店名或標記, 原因), ...]
      out_path: 存的檔路徑 (每日配送_YYYYMMDD.xlsx) 或 None(若沒裝 openpyxl)
    """
    if openpyxl is None:
        raise RuntimeError("請先安裝 openpyxl: uv pip install openpyxl")

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    raw = list(ws.iter_rows(values_only=True))
    if not raw:
        return [], [], None

    headers = [str(h) if h is not None else "" for h in raw[0]]
    hnorm = {_norm(h): h for h in headers}

    col_name = _match_col(hnorm, _NAME)
    col_addr = _match_col(hnorm, _ADDR)
    col_qty = _match_col(hnorm, _QTY)
    col_veh = _match_col(hnorm, _VEH)

    skipped = []
    rows = []
    for i, r in enumerate(raw[1:], 1):
        if all(v is None or v == "" for v in r):
            continue  # 空列跳過
        name = str(r[headers.index(col_name)] if col_name else "") or ""
        addr = str(r[headers.index(col_addr)] if col_addr else "") or ""
        qty = _clean_int(r[headers.index(col_qty)] if col_qty else "")
        veh = str(r[headers.index(col_veh)] if col_veh else "") if col_veh else ""
        veh = veh.strip()
        name = name.strip()
        addr = addr.strip()
        if not name and not addr:
            skipped.append((f"第{i}列", "店家與地址皆空"))
            continue
        if not addr:
            skipped.append((name or f"第{i}列", "缺少店家地址"))
            continue
        # 車號若來源沒有 → 留空，由 caller 決定
        rows.append((veh or "", name or f"店家{i}", addr, qty))

    if not rows:
        return [], skipped, None

    # 存成標準檔：每日配送_YYYYMMDD.xlsx
    date_str = date_str or datetime.datetime.now().strftime("%Y%m%d")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f"每日配送_{date_str}.xlsx")
    from openpyxl import Workbook
    wb2 = Workbook()
    ws2 = wb2.active
    ws2.title = "每日配送"
    ws2.append(["車號", "店家名稱", "店家地址", "瓶數"])
    for veh, name, addr, qty in rows:
        ws2.append([veh, name, addr, qty])
    wb2.save(out_path)
    return rows, skipped, out_path


def region_of(addr: str) -> str:
    """從地址判斷縣市(用於無車號時的地理分車)。"""
    for kw, label in _REGION_KEYWORDS:
        if kw in addr:
            return label
    return "其他"


def auto_assign_vehicles(rows):
    """
    來源無車號時，由 Agent 自己決定車號。
    策略：先全部併為 1 台車（車01）；
          若 caller 發現 1 台車跑完超過 17:30 回倉，
          再呼叫 split_by_region 按地理區域拆成多台。
    這裡回傳「全併 1 台車」版本（最保守起點），
    拆車由 plan 層回傳超時資訊後再呼叫 split_by_region。
    """
    return [("車01", n, a, q) for (_, n, a, q) in rows]


_REGION_LABEL_TO_VEH = {}
def split_by_region(rows):
    """把無車號 rows 按縣市拆成多台車：每台車 = 一個縣市。"""
    groups = {}
    order = []
    for (_, n, a, q) in rows:
        reg = region_of(a)
        if reg not in groups:
            groups[reg] = []
            order.append(reg)
        groups[reg].append((n, a, q))
    out = []
    for i, reg in enumerate(order, 1):
        veh = f"車{i:02d}"
        for (n, a, q) in groups[reg]:
            out.append((veh, n, a, q))
    return out


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        p = sys.argv[1]
        out = sys.argv[2] if len(sys.argv) > 2 else os.path.dirname(p)
        rows, skipped, outp = normalize_excel(p, out)
        print(f"讀到 {len(rows)} 筆，跳過 {len(skipped)} 筆")
        for r in rows[:5]:
            print(" ", r)
        if outp:
            print("標準檔已存:", outp)
