"""
data_loader.py — 讀取店家資料 (Excel / CSV)，照車號分組

預期欄位（中文或英文皆可）：
  車號       / 車輛 / 路線編號 / route / vehicle      -> 分組依據
  店家名稱   / 名稱 / 店名 / name
  店家地址   / 地址 / address
  瓶數       / 數量 / qty / bottles
  出發點地址 / 起點地址 / 倉庫地址 / start_address / depot_address  -> 每台車出發點

自動：
  - 地理編碼 (店家地址 + 出發點地址 -> 座標)
  - 由瓶數計算 下貨時間 (瓶數 * 15 秒)
  - 照車號分組，每台車獨立回傳

回傳 (vehicles, stops_by_vehicle, skipped)
  vehicles: [Vehicle]           每台的起點座標
  stops_by_vehicle: {車號: [Stop]}
  skipped: [(名稱, 原因)]
"""

import os
import csv

try:
    import openpyxl
except ImportError:
    openpyxl = None

from geocoder import geocode
from route_planner import Stop, Vehicle

SERVICE_SEC_PER_BOTTLE = 15.0


def _norm_header(h):
    return (h or "").strip().lower().replace(" ", "").replace("　", "")


_VEH_KEYS = ["車號", "車輛", "路線編號", "路線", "route", "vehicle", "車"]
_NAME_KEYS = ["店家名稱", "名稱", "店名", "客戶名稱", "name", "店家"]
_ADDR_KEYS = ["店家地址", "地址", "客戶地址", "address", "addr"]
_QTY_KEYS = ["瓶數", "數量", "箱數", "瓶量", "qty", "bottles", "count"]
_START_KEYS = ["出發點地址", "起點地址", "倉庫地址", "出發地址",
               "start_address", "depot_address", "start", "depot"]
_FUEL_KEYS = ["油資單價", "油資", "油錢單價", "fuel", "fuel_cost", "fuel_cost_per_km", "元每km"]


def _read_fuel_from_rows(rows, headers):
    """從資料列讀『油資單價』欄：回傳第一個有效值(float)，全空則回 None。全車共用。"""
    idx_by_norm = {_norm_header(h): h for h in headers}
    for row in rows:
        v = _get(row, _FUEL_KEYS, idx_by_norm)
        if v not in (None, ""):
            try:
                f = float(v)
                if f > 0:
                    return f
            except (ValueError, TypeError):
                pass
    return None


def read_fuel_cost(path):
    """從資料檔(xlsx/csv)讀『油資單價』欄，回傳 float 或 None（供 main 覆蓋 .env）。"""
    try:
        if str(path).lower().endswith(".csv"):
            with open(path, encoding="utf-8-sig", newline="") as f:
                reader = csv.DictReader(f)
                return _read_fuel_from_rows([dict(r) for r in reader], reader.fieldnames or [])
        if openpyxl is None:
            return None
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        rows_raw = list(ws.iter_rows(values_only=True))
        if not rows_raw:
            return None
        headers = [str(h) if h is not None else "" for h in rows_raw[0]]
        rows = [{headers[i]: r[i] for i in range(len(headers))}
                for r in rows_raw[1:] if not all(v is None for v in r)]
        return _read_fuel_from_rows(rows, headers)
    except Exception:
        return None


def _get(row, keys, idx_by_norm):
    for k in keys:
        orig = idx_by_norm.get(k)
        if orig is not None and orig in row:
            return row[orig]
    return None


def load_from_rows(rows, headers, depot=None):
    idx_by_norm = {_norm_header(h): h for h in headers}
    vehicles = {}            # 車號 -> Vehicle (暫存)
    stops_by_vehicle = {}    # 車號 -> [Stop]
    skipped = []
    geo_jobs = []            # [(label, address, callback)]

    for i, row in enumerate(rows):
        veh_raw = _get(row, _VEH_KEYS, idx_by_norm)
        name = str(_get(row, _NAME_KEYS, idx_by_norm) or f"店家{i+1}").strip()
        addr = str(_get(row, _ADDR_KEYS, idx_by_norm) or "").strip()
        start_raw = _get(row, _START_KEYS, idx_by_norm)
        start_addr = str(start_raw or "").strip() if start_raw is not None else ""
        qty_raw = _get(row, _QTY_KEYS, idx_by_norm)
        try:
            qty = float(qty_raw) if qty_raw not in (None, "") else 0.0
        except ValueError:
            qty = 0.0
        svc = qty * SERVICE_SEC_PER_BOTTLE

        veh = str(veh_raw).strip() if veh_raw not in (None, "") else "未分車"
        if not addr:
            skipped.append((name, "缺少店家地址"))
            continue

        stop = Stop(id=f"{veh}-{len(stops_by_vehicle.get(veh, []))+1:03d}",
                    name=name, lat=0.0, lon=0.0, demand=qty,
                    service_time=svc, address=addr, vehicle=veh)
        stops_by_vehicle.setdefault(veh, []).append(stop)
        # 起點：若有總倉 depot 則統一用總倉；否則用 Excel 出發點地址欄
        if depot is not None:
            vehicles.setdefault(veh, {"name": veh, "depot": True,
                                      "start_addr": depot.address,
                                      "lat": depot.lat, "lon": depot.lon})
        else:
            if veh not in vehicles:
                vehicles[veh] = {"name": veh, "start_addr": start_addr}
            geo_jobs.append(("start", veh, None, start_addr))
        # 地理編碼任務（僅店家）
        geo_jobs.append(("stop", veh, len(stops_by_vehicle[veh]) - 1, addr))

    # 批量地理編碼（僅店家；depot 模式不起點）
    addrs = [job[3] for job in geo_jobs if job[0] == "stop"]
    geo_result = geocode_batch(addrs)
    addr_to_coord = dict(zip(addrs, geo_result))

    # 回填
    for kind, veh, sidx, addr in geo_jobs:
        if kind != "stop":
            continue
        coord = addr_to_coord.get(addr)
        st = stops_by_vehicle[veh][sidx]
        if coord:
            st.lat, st.lon = coord
        else:
            skipped.append((st.name, f"店家地址無法定位：{addr}"))
            st.lat = None

    # 去掉無效 stop
    for veh in list(stops_by_vehicle.keys()):
        stops_by_vehicle[veh] = [s for s in stops_by_vehicle[veh] if s.lat is not None]
        if not stops_by_vehicle[veh]:
            del stops_by_vehicle[veh]

    # 建 Vehicle 物件
    vlist = []
    for veh, info in vehicles.items():
        if info.get("lat") is None:
            sk = stops_by_vehicle.get(veh)
            if sk:
                info["lat"], info["lon"] = sk[0].lat, sk[0].lon
                info["start_addr"] = info.get("start_addr") or sk[0].address
            else:
                continue
        vlist.append(Vehicle(id=veh, name=veh,
                             start_lat=info["lat"], start_lon=info["lon"],
                             start_addr=info.get("start_addr", "")))

    return vlist, stops_by_vehicle, skipped


# 用 geocoder 的批次（含快取）
def geocode_batch(addresses):
    from geocoder import geocode_many
    return [geocode_many([a]).get(a) for a in addresses]


def load_excel(path, depot=None):
    if openpyxl is None:
        raise RuntimeError("尚未安裝 openpyxl：請執行 uv pip install openpyxl")
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows_raw = list(ws.iter_rows(values_only=True))
    if not rows_raw:
        return [], {}, []
    headers = [str(h) if h is not None else "" for h in rows_raw[0]]
    rows = []
    for r in rows_raw[1:]:
        if all(v is None for v in r):
            continue
        rows.append({headers[i]: r[i] for i in range(len(headers))})
    return load_from_rows(rows, headers, depot=depot)


def load_csv(path, depot=None):
    with open(path, encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        headers = reader.fieldnames or []
        rows = [dict(r) for r in reader]
    return load_from_rows(rows, headers, depot=depot)


def load(path, depot=None):
    ext = os.path.splitext(path)[1].lower()
    # OneDrive / Excel 開啟中常鎖住 xlsx 導致 PermissionError。
    # 先嘗試直接讀；失敗則複製到本地 temp 再讀（並重試，因鎖檔通常短暫）。
    read_path = _copy_to_temp(path)
    if ext in (".xlsx", ".xlsm"):
        return load_excel(read_path, depot=depot)
    if ext in (".csv", ".txt"):
        return load_csv(read_path, depot=depot)
    raise ValueError(f"不支援的檔案格式：{ext}")


def _copy_to_temp(path):
    """複製到本地 temp 再回傳副本路徑（繞過 OneDrive/Excel 鎖檔）。
    若原檔可直讀就直接用；否則重試複製幾次（鎖檔通常短暫），都失敗則退回原路徑。"""
    import shutil, tempfile, time
    # 先試原檔能否直開
    try:
        with open(path, "rb") as f:
            f.read(1)
        return path
    except (PermissionError, OSError):
        pass
    d = tempfile.gettempdir()
    dst = os.path.join(d, "logi_" + os.path.basename(path))
    last_err = None
    for attempt in range(5):
        try:
            shutil.copy2(path, dst)
            return dst
        except Exception as e:
            last_err = e
            time.sleep(1.5)   # 等 OneDrive/Excel 釋放鎖
    # 都失敗：嘗試原檔（讓上層報錯時能給明確訊息）
    return path
