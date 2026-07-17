"""
logistics_agent.py — 鮮奶物流「分組路線規劃」Agent

模式：照 Excel 的「車號」分組，每台車獨立排出最順拜訪順序。
     每台車有各自的出發點(出發點地址欄)。計算每間店下貨時間(瓶數*15秒)
     與到店預估時間(ETA)。

用法：
  python logistics_agent.py --data 每日配送.xlsx
  python logistics_agent.py --data 每日配送.csv
  python logistics_agent.py --make-sample        # 產生範例 Excel 模板
  python logistics_agent.py --straight           # 強制直線距離/估速
  python logistics_agent.py --no-google          # 暫不走 Google(帳單未啟用時)
  python logistics_agent.py --start 7.5          # 出發時間 07:30
"""

import os
import argparse

from route_planner import Stop, Vehicle, solve_grouped, _hhmm
from data_loader import load
import report as report_mod
from google_maps import distance_matrix as g_distance_matrix
from osrm_client import get_matrix_dur as osrm_matrix_dur


# 總倉（所有車統一出發點 / 收班點）
DEPOT_ADDR = "台中市大雅區101-1號"
DEPOT = Stop("DEPOT", f"總倉 ({DEPOT_ADDR})", 0.0, 0.0, address=DEPOT_ADDR)

# 出車時間與目標回倉
DEFAULT_START_HOUR = 9.5        # 司機正常出車 09:30
TARGET_RETURN_HOUR = 17.5       # 目標下午 17:30 回到倉庫

# 預設資料/報表資料夾：使用者的 OneDrive 桌面「路線規劃」
# (你的桌面是 OneDrive 同步桌面，故預設指向此路徑，方便直接在桌面操作)
def _default_data_dir():
    base = os.path.join(os.path.expanduser("~"), "OneDrive", "桌面", "路線規劃")
    if not os.path.isdir(base):
        base = os.path.join(os.path.expanduser("~"), "Desktop", "路線規劃")
    os.makedirs(base, exist_ok=True)
    return base

DATA_DIR = _default_data_dir()

def _default_report_dir():
    onedrive_desk = os.path.join(os.path.expanduser("~"), "OneDrive", "桌面")
    if os.path.isdir(onedrive_desk):
        base = os.path.join(onedrive_desk, "當日車輛報表")
    else:
        base = os.path.join(os.path.expanduser("~"), "Desktop", "當日車輛報表")
    os.makedirs(base, exist_ok=True)
    return base

REPORT_DIR = _default_report_dir()

def _default_dispatch_dir():
    # 派車單 / 司機派遣清單：每台車一份，給司機/內勤拿著跑，與路線規劃總報表(當日車輛報表)分開。
    onedrive_desk = os.path.join(os.path.expanduser("~"), "OneDrive", "桌面")
    if os.path.isdir(onedrive_desk):
        base = os.path.join(onedrive_desk, "當日派車單")
    else:
        base = os.path.join(os.path.expanduser("~"), "Desktop", "當日派車單")
    os.makedirs(base, exist_ok=True)
    return base

DISPATCH_DIR = _default_dispatch_dir()


def make_sample(path):
    try:
        import openpyxl
    except ImportError:
        raise RuntimeError("請先安裝 openpyxl: uv pip install openpyxl")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "每日配送"
    ws.append(["車號", "店家名稱", "店家地址", "瓶數", "油資單價"])
    sample = [
        ("車01", "台中港區超市", "台中市梧棲區港埠路一段1號", 30, 5.0),
        ("車01", "沙鹿鮮奶坊", "台中市沙鹿區中山路1號", 45, None),
        ("車01", "清水門市", "台中市清水區中山路1號", 20, None),
        ("車01", "大雅市場店", "台中市大雅區中清路一段1號", 35, None),
        ("車02", "豐原車站店", "台中市豐原區中正路1號", 55, None),
        ("車02", "潭子門市", "台中市潭子區潭子街1號", 28, None),
        ("車02", "神岡便利商店", "台中市神岡區神岡路1號", 24, None),
        ("車02", "后里批發", "台中市后里區甲后路一段1號", 30, None),
        ("車03", "西屯辦公室", "台中市西屯區台灣大道四段1號", 40, None),
        ("車03", "南屯量販", "台中市南屯區公益路二段1號", 25, None),
        ("車03", "北屯門市", "台中市北屯區文心路四段1號", 18, None),
        ("車03", "太平商店", "台中市太平區太平路1號", 18, None),
    ]
    for r in sample:
        ws.append(list(r))
    ws2 = wb.create_sheet("填寫說明")
    notes = [
        ["欄位", "是否必填", "說明"],
        ["車號", "必填", "同一台車的點填相同車號，如 車01 / 車A。程式照車號分組，每台車獨立排順序"],
        ["店家名稱", "必填", "店家名稱，會顯示在報表與地圖"],
        ["店家地址", "必填", "完整地址(到門牌)，用 Google 精準定位。例: 台中市大雅區101-1號"],
        ["瓶數", "必填", "鮮奶瓶數。下貨時間=瓶數×10秒，並計入到店預估時間"],
        ["油資單價", "選填", "元/公里，全車共用；只需填第一列(其餘留空)。填了以此為準，留空則用 .env 設定。報表/派車單會顯示每車與總油資"],
        ["", "", ""],
        ["出發點", "總倉統一", f"所有車都從總倉出發並回到總倉：{DEPOT_ADDR}（已在程式設定，Excel 不需填出發點）"],
        ["用法", "", "填好後執行: python logistics_agent.py --data 路徑/每日配送.xlsx"],
        ["注意", "", "欄位名稱(第一列)請勿改；下方範例列可直接覆蓋成真實資料"],
    ]
    for r in notes:
        ws2.append(r)
    wb.save(path)
    print(f"✓ 範例模板已產生：{path}")
    print(f"   (所有車統一從總倉 {DEPOT_ADDR} 出發)")
    print("  請用真實資料替換後執行：")
    print(f"  python logistics_agent.py --data {os.path.basename(path)}")


def build_matrices(vehicles, stops_by_vehicle, use_google, no_google):
    """每台車獨立取得真實距離+時間矩陣。回傳 (matrix_km, duration_matrix, source)。"""
    matrix_km = {}
    duration_matrix = {}
    source = "haversine"
    if not use_google:
        # 回 None（不是空 dict）：solve_grouped 的 _Dist 在 matrix is None 時
        # 才會 fallback 到 haversine 真算；回空 dict 會讓 .get 全取 0 → 0km 假數據。
        return None, None, source

    for v in vehicles:
        stops = stops_by_vehicle.get(v.id, [])
        if not stops:
            continue
        coords = [(v.start_lat, v.start_lon)] + [(s.lat, s.lon) for s in stops]
        # 優先 Google
        ok = False
        if not no_google:
            try:
                m, d, src = g_distance_matrix(coords)
                matrix_km, duration_matrix = _merge(matrix_km, duration_matrix, v.id, m, d)
                source = src if source == "haversine" else source
                ok = True
            except Exception:
                pass
        if not ok:
            # 降級 OSRM
            try:
                m, d, src = osrm_matrix_dur(coords)
                if src != "fallback":
                    matrix_km, duration_matrix = _merge(matrix_km, duration_matrix, v.id, m, d)
                    if source == "haversine":
                        source = "osrm"
                    ok = True
            except Exception:
                pass
        if not ok:
            # 都失敗：降級直線
            matrix_km, duration_matrix = _merge_line(matrix_km, duration_matrix, v.id, coords)
            source = "fallback"
    return matrix_km, duration_matrix, source


def _merge(m_km, m_dur, vid, mat, dur):
    n = len(mat)
    for i in range(n):
        for j in range(n):
            m_km[(vid, i, j)] = mat[i][j]
            m_dur[(vid, i, j)] = dur[i][j]
    return m_km, m_dur


def _merge_line(m_km, m_dur, vid, coords):
    from route_planner import haversine
    n = len(coords)
    for i in range(n):
        for j in range(n):
            km = haversine(coords[i][0], coords[i][1], coords[j][0], coords[j][1])
            m_km[(vid, i, j)] = km
            m_dur[(vid, i, j)] = km / 30.0 * 3600.0
    return m_km, m_dur


def _load_fuel_cost():
    """讀油資參數：優先環境變數(Render 用)，其次 .env 檔(本機用)；
    優先 FUEL_COST_PER_KM，否則用 油耗×油價(FUEL_EFFICIENCY×FUEL_PRICE)推算。"""
    env = {}
    p = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if os.path.exists(p):
        for line in open(p, encoding="utf-8"):
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                env[k.strip()] = v.strip().strip('"').strip("'")

    def get(key):
        # 環境變數優先(Render)，再退回 .env(本機)
        return os.environ.get(key) or env.get(key)

    if get("FUEL_COST_PER_KM"):
        try:
            return float(get("FUEL_COST_PER_KM"))
        except ValueError:
            pass
    eff = get("FUEL_EFFICIENCY")
    price = get("FUEL_PRICE")
    if eff and price:
        try:
            eff = float(eff); price = float(price)
            if eff > 0:
                return price / eff
        except ValueError:
            pass
    return 0.0


def plan(start_hour, data_path, use_google, no_google, fuel_cost_per_km=0.0):
    here = os.path.dirname(os.path.abspath(__file__))
    # 若呼叫方沒給油資，從 .env 讀
    if not fuel_cost_per_km:
        fuel_cost_per_km = _load_fuel_cost()
    # 先定位總倉座標
    if DEPOT.lat == 0.0 and DEPOT.lon == 0.0:
        try:
            from geocoder import geocode
            c = geocode(DEPOT.address)
            if c:
                DEPOT.lat, DEPOT.lon = c
        except Exception:
            pass
    print(f"🏭 總倉出發點：{DEPOT.address}  ({DEPOT.lat:.4f}, {DEPOT.lon:.4f})")

    print(f"📂 讀取店家資料：{data_path}")
    vehicles, stops_by_vehicle, skipped = load(data_path, depot=DEPOT)
    if not vehicles:
        print("⚠ 沒有可規劃的車輛/店家，請檢查資料或改 --straight 試。")
        return None, skipped
    total_stops = sum(len(v) for v in stops_by_vehicle.values())
    print(f"   共 {len(vehicles)} 台車，{total_stops} 個配送點，跳過 {len(skipped)} 筆。")

    print("🌐 取得真實道路距離 + 行車時間 ...")
    matrix_km, duration_matrix, source = build_matrices(
        vehicles, stops_by_vehicle, use_google, no_google)
    src_map = {"haversine": "直線估算", "google": "Google Maps",
               "google-cache": "本機快取(零費用)", "osrm": "OSRM",
               "fallback": "直線估算(降級)"}
    print(f"   距離來源：{src_map.get(source, source)}")
    try:
        import geo_cache
        print("   " + geo_cache.stats_line())
    except Exception:
        pass

    result = solve_grouped(
        vehicles, stops_by_vehicle,
        matrix_km=matrix_km, duration_matrix=duration_matrix,
        distance_source=source, start_hour=start_hour,
        fuel_cost_per_km=fuel_cost_per_km,
    )
    # 目標回倉標註：每台車預計回倉是否 <= 17:30
    for rt in result.routes:
        rt["return_hour"] = rt.get("end_hour", 0)
        rt["on_time"] = rt["return_hour"] <= TARGET_RETURN_HOUR
    result.skipped = skipped
    return result, skipped


def plan_auto_assign(start_hour, data_path, use_google, no_google, fuel_cost_per_km=0.0,
                      max_vehicles=3, force_vehicles=None):
    """無車號時的單輪自動分車：全站只打「一次」真實距離矩陣，
    拆群(1→3台)時直接從同一次矩陣切出子矩陣複用，避免重複打 Google/OSRM。
    force_vehicles: 指定車數(如 2) → 強制分剛好 N 台，且忽略 17:30 時間窗
                    （只求各車最快回倉、最短路線）。None=由 Agent 自動決定。
    """
    here = os.path.dirname(os.path.abspath(__file__))
    if not fuel_cost_per_km:
        fuel_cost_per_km = _load_fuel_cost()
    # 總倉座標
    if DEPOT.lat == 0.0 and DEPOT.lon == 0.0:
        try:
            from geocoder import geocode
            c = geocode(DEPOT.address)
            if c:
                DEPOT.lat, DEPOT.lon = c
        except Exception:
            pass
    print(f"🏭 總倉出發點：{DEPOT.address}  ({DEPOT.lat:.4f}, {DEPOT.lon:.4f})")

    print(f"📂 讀取店家資料：{data_path}")
    vehicles, stops_by_vehicle, skipped = load(data_path, depot=DEPOT)
    if not vehicles:
        print("⚠ 沒有可規劃的車輛/店家，請檢查資料。")
        return None, skipped

    # 合并所有站點（無車號 → 全部歸到同一臨時群）
    all_stops = []
    for v in vehicles:
        all_stops.extend(stops_by_vehicle.get(v.id, []))

    # ---- 嘉義列為例外：剔除不參與分車/最佳化（單程過遠會拖垮時間窗）----
    chiayi_stops = [s for s in all_stops
                    if "嘉義" in (getattr(s, "address", "") or "")
                    or "嘉義" in (getattr(s, "name", "") or "")]
    if chiayi_stops:
        all_stops = [s for s in all_stops if s not in chiayi_stops]
        for s in chiayi_stops:
            skipped.append((s.name, "嘉義例外(單程過遠，不納入時間窗最佳化)"))
        print(f"   ⚠ 嘉義例外：{len(chiayi_stops)} 站不納入自動分車（列為例外）。")

    total_stops = len(all_stops)
    if total_stops == 0:
        return None, skipped
    print(f"   共 {total_stops} 個配送點（無車號 → 由 Agent 自動分車，最多 {max_vehicles} 台）。")

    # ---- 全站只打一次真實矩陣 ----
    coords = [(DEPOT.lat, DEPOT.lon)] + [(s.lat, s.lon) for s in all_stops]  # 0=倉, 1..n=站
    FULL = "__ALL__"
    matrix_km_full = None
    duration_matrix_full = None
    source = "haversine"

    if use_google:
        ok = False
        if not no_google:
            try:
                m, d, src = g_distance_matrix(coords, fast_fail=True)
                n = len(coords)
                matrix_km_full = {}
                duration_matrix_full = {}
                for i in range(n):
                    for j in range(n):
                        matrix_km_full[(FULL, i, j)] = m[i][j]
                        duration_matrix_full[(FULL, i, j)] = d[i][j]
                source = src
                ok = True
            except Exception:
                pass
        if not ok:
            try:
                m, d, src = osrm_matrix_dur(coords)
                if src != "fallback":
                    n = len(coords)
                    matrix_km_full = {}
                    duration_matrix_full = {}
                    for i in range(n):
                        for j in range(n):
                            matrix_km_full[(FULL, i, j)] = m[i][j]
                            duration_matrix_full[(FULL, i, j)] = d[i][j]
                    source = src
                    ok = True
            except Exception:
                pass
        if not ok:
            # 都失敗 → 降級直線（matrix 留 None，solver 用 Haversine）
            source = "fallback"
            matrix_km_full = None
            duration_matrix_full = None

    src_map = {"haversine": "直線估算", "google": "Google Maps",
               "google-cache": "本機快取(零費用)",
               "osrm": "OSRM", "fallback": "直線估算(降級)"}
    print(f"   距離來源：{src_map.get(source, source)}（全站單次矩陣）")
    try:
        import geo_cache as _gc
        print("   " + _gc.stats_line())
    except Exception:
        pass

    # ---- 依時間窗決定分幾群（均衡分車：k-means 初始 + 負載均衡） ----
    import auto_router
    # 先準備 (i,j) 直接索引矩陣（0=倉, 1..n=站），force 分支也要用
    m_ij = d_ij = None
    if matrix_km_full is not None:
        n = len(coords)
        m_ij = {(i, j): matrix_km_full[(FULL, i, j)] for i in range(n) for j in range(n)}
        d_ij = {(i, j): duration_matrix_full[(FULL, i, j)] for i in range(n) for j in range(n)}
    # 指定車數模式：強制剛好 N 台，且忽略 17:30 時間窗（只求最快回倉/最短路線）
    if force_vehicles:
        fk = max(1, min(int(force_vehicles), max_vehicles, len(all_stops)))
        target = 99.0   # 極大值 → 不算時間窗，純追求最短路線/最快回倉
        if matrix_km_full is not None:
            groups = auto_router.balanced_groups(
                all_stops, DEPOT, start_hour, target, max_vehicles=fk,
                matrix=m_ij, duration=d_ij, force_k=fk)
        else:
            groups = auto_router.balanced_groups(
                all_stops, DEPOT, start_hour, target, max_vehicles=fk, force_k=fk)
    elif matrix_km_full is not None:
        # balanced_groups 需要 (i,j) 直接索引的矩陣（0=倉, 1..n=站），
        # 這裡從 (FULL,i,j) 轉成 (i,j)
        groups = auto_router.balanced_groups(
            all_stops, DEPOT, start_hour, TARGET_RETURN_HOUR, max_vehicles=max_vehicles,
            matrix=m_ij, duration=d_ij)
    else:
        # 無矩陣（直線模式）：用 haversine 粗估均衡
        groups = auto_router.balanced_groups(
            all_stops, DEPOT, start_hour, TARGET_RETURN_HOUR, max_vehicles=max_vehicles)

    # ---- 拆成多台車，並從全矩陣切出各車子矩陣，逐群 solve ----
    new_vehicles = []
    new_stops = {}
    for gi, g in enumerate(groups, 1):
        veh_id = f"車{gi:02d}"
        sub = [all_stops[si] for si in g]
        new_vehicles.append(Vehicle(id=veh_id, name=veh_id,
                                    start_lat=DEPOT.lat, start_lon=DEPOT.lon,
                                    start_addr=DEPOT.address))
        new_stops[veh_id] = sub

    # 逐群用全域矩陣切片成局部矩陣後 solve（與 auto_router.group_end_hour_real 同一套切片，確保一致）
    per_vehicle_routes = []
    total_dist = 0.0
    total_load = 0.0
    total_fuel = 0.0
    for gi, g in enumerate(groups, 1):
        veh_id = f"車{gi:02d}"
        sub = new_stops[veh_id]
        km_local = None
        dur_local = None
        if matrix_km_full is not None:
            k = len(g)
            local_to_global = [0] + [si + 1 for si in g]
            km_local = {(veh_id, 0, 0): 0.0}
            dur_local = {(veh_id, 0, 0): 0.0}
            for li in range(k + 1):
                for lj in range(k + 1):
                    gn_i = local_to_global[li]
                    gn_j = local_to_global[lj]
                    km_local[(veh_id, li, lj)] = matrix_km_full[(FULL, gn_i, gn_j)]
                    dur_local[(veh_id, li, lj)] = duration_matrix_full[(FULL, gn_i, gn_j)]
        res = solve_grouped([new_vehicles[gi - 1]], {veh_id: sub},
                            matrix_km=km_local, duration_matrix=dur_local,
                            distance_source=source, start_hour=start_hour,
                            fuel_cost_per_km=fuel_cost_per_km)
        rt = res.routes[0]
        rt["vehicle"] = new_vehicles[gi - 1]
        rt["return_hour"] = rt.get("end_hour", 0)
        rt["on_time"] = rt["return_hour"] <= TARGET_RETURN_HOUR
        per_vehicle_routes.append(rt)
        total_dist += rt["distance_km"]
        total_load += rt["load"]
        total_fuel += rt.get("fuel_cost", 0)

    # 組成 PlanResult
    from route_planner import PlanResult
    result = PlanResult()
    result.routes = sorted(per_vehicle_routes, key=lambda r: r["vehicle"].id)
    result.distance_source = source
    result.fuel_cost_per_km = fuel_cost_per_km
    result.total_distance_km = total_dist
    result.total_load = total_load
    result.total_fuel_cost = total_fuel
    result.skipped = skipped
    result.summary = _make_summary_like(result)
    return result, skipped


def _make_summary_like(result):
    src_map = {"haversine": "直線估算", "google": "Google Maps",
               "osrm": "OSRM", "fallback": "直線估算(降級)"}
    src = src_map.get(result.distance_source, result.distance_source)
    lines = [f"出車數：{len(result.routes)} 台",
             f"距離來源：{src}",
             f"總配送距離：{result.total_distance_km:.1f} km",
             f"總瓶數：{result.total_load:.0f}"]
    if result.fuel_cost_per_km > 0:
        lines.append(f"油資單價：{result.fuel_cost_per_km:.1f} 元/km")
        lines.append(f"預估總油資：{result.total_fuel_cost:.0f} 元")
    for rt in result.routes:
        v = rt["vehicle"]
        ret = _hhmm(rt["end_hour"])
        tag = "準時回倉" if rt.get("on_time") else f"超過17:30({ret})"
        lines.append(f"  {v.id}：{len(rt['stops'])} 站 / {rt['distance_km']:.1f} km / "
                     f"{rt['load']:.0f} 瓶 / 回到起點 {ret}（{tag}）")
    return "\n".join(lines)


def self_check(result, args=None):
    """結果產出後自動複檢（每次跑完都呼叫，印在終端不污染報表）。
    檢查三個已知坑：下貨秒數一致性 / 里程非 0 / 回倉標註自洽。"""
    from data_loader import SERVICE_SEC_PER_BOTTLE
    print("\n🔍 產出自動複檢：")
    issues = []

    # 1) 下貨秒數一致性：footer 文字若硬寫秒數，必須與 SERVICE_SEC_PER_BOTTLE 對齊
    #    （report.py footer 寫『每瓶 N 秒』，改 SERVICE_SEC_PER_BOTTLE 時要同步）
    try:
        from report import __file__ as _rp
        rsrc = open(_rp, encoding="utf-8").read()
        import re
        m = re.search(r"每瓶\s*(\d+)\s*秒", rsrc)
        if m and int(m.group(1)) != int(SERVICE_SEC_PER_BOTTLE):
            issues.append(f"⚠ 報表 footer 寫『每瓶 {m.group(1)} 秒』但 SERVICE_SEC_PER_BOTTLE={int(SERVICE_SEC_PER_BOTTLE)}，兩者不一致！改 data_loader.py 秒數時要同步 report.py footer。")
        else:
            print(f"   ① 下貨秒數一致：{int(SERVICE_SEC_PER_BOTTLE)} 秒/瓶（footer 對齊）")
    except Exception as e:
        print(f"   ① 下貨秒數檢查跳過（{e}）")

    # 2) 里程非 0：每車 distance_km 應 > 0（曾踩坑：矩陣 key 漏 veh_id → 0km 假準時）
    zero_km = [rt["vehicle"].id for rt in result.routes if rt.get("distance_km", 0) <= 0]
    if zero_km:
        issues.append(f"⚠ 以下車里程為 0km（疑距離矩陣沒吃到真實道路）：{', '.join(zero_km)}")
    else:
        parts = [f"{rt['vehicle'].id}={rt['distance_km']:.1f}km" for rt in result.routes]
        print(f"   ② 各車里程非零：{', '.join(parts)}")

    # 3) 回倉標註自洽：on_time 布林應 == (end_hour <= TARGET_RETURN_HOUR)
    bad = []
    for rt in result.routes:
        calc = rt.get("end_hour", 0) <= TARGET_RETURN_HOUR
        if rt.get("on_time") != calc:
            bad.append(rt["vehicle"].id)
    if bad:
        issues.append(f"⚠ 回倉準時標註與計算不符：{', '.join(bad)}")
    else:
        print(f"   ③ 回倉標註自洽：準時窗 {_hhmm(TARGET_RETURN_HOUR)}（目標回倉）")

    # 4) 回倉時間：基於報表自身 etas 做健全性檢查（不引入第二種距離度量，
    #    避免 haversine 直線 vs 真實道路的系統偏差誤報）
    #    - etas 須單調遞增（下一站到達 >= 本站離開）
    #    - end_hour 須 >= 最後一站離開（回倉必晚於最後離開）
    #    - 回程間隔須在合理範圍 (0 < gap < 12h)，抓 end_hour 被硬設的異常
    back_warn = []
    for rt in result.routes:
        v = rt["vehicle"]; etas = rt.get("etas", [])
        if not etas:
            continue
        mono = all(etas[i + 1][0] >= etas[i][1] for i in range(len(etas) - 1))
        last_leave = etas[-1][1]
        end = rt.get("end_hour", 0)
        gap = end - last_leave
        if not mono:
            back_warn.append(f"{v.id}(ETA 非單調遞增)")
        elif not (0 < gap < 12):
            back_warn.append(f"{v.id}(回程間隔 {gap*60:.0f}分不合理)")
        else:
            tag = "準時" if end <= TARGET_RETURN_HOUR else "超過17:30"
            print(f"   ④ 回倉時間核對：{v.id} 預計 {_hhmm(end)} 回倉（{tag}）；ETA 鏈單調 ✅")
    if back_warn:
        issues.append("⚠ 回倉時間異常：" + "; ".join(back_warn))

    # 5) 路線最佳化：用 multi-restart 2-opt（多隨機種子取最小）獨立重優化，
    #    跟 solve_grouped 產出的現有順序比。若還能降 >5% 行車 → 路線未充分最佳化。
    #    （單次 2-opt 太弱、連明顯繞遠都解不開，故用 multi-restart 做有意義的交叉驗證）
    try:
        import random
        from route_planner import Stop as _Stop, _Dist as _DistR, _route_duration_sec, _two_opt  # noqa
        opt_warn = []
        for rt in result.routes:
            v = rt["vehicle"]; stops = rt["stops"]
            if len(stops) < 4:
                continue
            nodes = [_Stop("START", v.start_addr or "起點", v.start_lat, v.start_lon)] + list(stops)
            dist = _DistR(nodes)
            n = len(nodes)
            ordered = list(range(1, n))  # rt["stops"] 已是規劃後順序
            cur_sec = _route_duration_sec(dist, 0, ordered)
            # multi-restart：8 個隨機種子各跑 2-opt，取最小行車秒
            best_sec = cur_sec
            for seed in range(8):
                rnd = list(range(1, n))
                random.Random(seed).shuffle(rnd)
                r = _two_opt(dist, 0, rnd)
                best_sec = min(best_sec, _route_duration_sec(dist, 0, r))
            if best_sec < cur_sec * 0.95:  # 還能優化 >5%
                opt_warn.append(f"{v.id}(現 {cur_sec/60:.0f}分→可降 {best_sec/60:.0f}分)")
        if opt_warn:
            issues.append("⚠ 路線可能未充分最佳化（multi-restart 仍可降 >5% 行車）：" + "; ".join(opt_warn))
        else:
            print("   ⑤ 路線最佳化：各車順序已接近最優（multi-restart 交叉驗證）")
    except Exception as e:
        print(f"   ⑤ 路線最佳化檢查跳過（{e}）")

    if issues:
        print("\n".join("   " + i for i in issues))
        print("   ⛔ 以上問題請先確認再交付報表。")
    else:
        print("   ✅ 複檢通過，無異常。")
    return len(issues) == 0


def main():
    ap = argparse.ArgumentParser(description="鮮奶物流分組路線規劃 Agent")
    ap.add_argument("--data", help="每日配送資料 (xlsx/csv)")
    ap.add_argument("--make-sample", action="store_true", help="產生範例 Excel 模板")
    ap.add_argument("--straight", action="store_true", help="強制直線距離/估速")
    ap.add_argument("--no-google", action="store_true", help="暫不走 Google，改 OSRM")
    ap.add_argument("--start", type=float, default=DEFAULT_START_HOUR, help="出發時間(24h, 預設9.5=09:30)")
    ap.add_argument("--auto", action="store_true", help="強制自動分車：忽略 Excel 車號欄，依 17:30 回倉自動分成多台車（最多3台）")
    ap.add_argument("--vehicles", type=int, default=None, help="搭配 --auto：強制分剛好 N 台車（忽略 17:30 時間窗，只求最短路線）")
    ap.add_argument("--excel", action="store_true",
                   help="額外產出整合 Excel（整合報表.xlsx，含路線圖分頁）。預設不產，需時再加此開關")
    args = ap.parse_args()

    here = os.path.dirname(os.path.abspath(__file__))

    if args.make_sample:
        make_sample(os.path.join(DATA_DIR, "每日配送.xlsx"))
        print(f"   (範本已產到桌面『路線規劃』資料夾：{os.path.join(DATA_DIR, '每日配送.xlsx')})")
        return

    if not args.data:
        # 預設找桌面『路線規劃』資料夾的每日配送.xlsx
        for cand in ["每日配送.xlsx", "每日配送.csv", "stores.xlsx"]:
            p = os.path.join(DATA_DIR, cand)
            if os.path.exists(p):
                args.data = p
                break
    if not args.data or not os.path.exists(args.data):
        print(f"⚠ 找不到資料（預設路徑：{DATA_DIR}）。")
        print("   請把填好的『每日配送.xlsx』放到桌面『路線規劃』資料夾，或加 --data 指定。")
        return

    use_google = not args.straight
    fuel_cost = _load_fuel_cost()
    # Excel『油資單價』欄優先於 .env（全車共用）
    try:
        from data_loader import read_fuel_cost as _read_fuel
        _excel_fuel = _read_fuel(args.data)
        if _excel_fuel:
            fuel_cost = _excel_fuel
            print(f"   (油資單價取自 Excel：{fuel_cost:.1f} 元/km)")
    except Exception:
        pass
    # 智慧判斷：--auto 強制自動分車；無車號 → 均衡自動分車(含嘉義例外)；有車號 → 照車號排序
    had_no_vehicle = False
    try:
        _v, _sbv, _sk = load(args.data, depot=DEPOT)
        had_no_vehicle = all((getattr(v, "id", "") or "").strip() == "未分車" for v in _v) if _v else False
    except Exception:
        had_no_vehicle = False
    if args.auto or had_no_vehicle:
        if args.auto and not had_no_vehicle:
            print("   (--auto：忽略 Excel 車號欄，強制走均衡自動分車，含嘉義例外)")
        else:
            print("   (偵測到無車號 → 走均衡自動分車，含嘉義例外)")
        result, skipped = plan_auto_assign(args.start, args.data, use_google, args.no_google,
                                           fuel_cost_per_km=fuel_cost, force_vehicles=args.vehicles)
    else:
        result, skipped = plan(args.start, args.data, use_google, args.no_google, fuel_cost_per_km=fuel_cost)
    if result is None:
        return

    print("=" * 60)
    print("📦 今日鮮奶配送路線規劃結果")
    print("=" * 60)
    print(result.summary)
    if result.fuel_cost_per_km > 0:
        print(f"\n⛽ 油資單價 {result.fuel_cost_per_km:.1f} 元/km ｜ 預估總油資 {result.total_fuel_cost:.0f} 元")
    print()
    for i, rt in enumerate(result.routes, 1):
        v = rt["vehicle"]
        ret = _hhmm(rt["end_hour"])
        tag = "✅ 準時回倉" if rt.get("on_time") else f"⚠ 超過 17:30（{ret}）"
        fuel_txt = f" ｜ 油資 {rt.get('fuel_cost',0):.0f} 元" if result.fuel_cost_per_km > 0 else ""
        print(f"【{v.id}｜起點 {v.start_addr}】{len(rt['stops'])} 站, "
              f"{rt['distance_km']:.1f} km, 回到起點 {ret}{fuel_txt}")
        print(f"   {tag}")
        for si, s in enumerate(rt["stops"]):
            a, lv = rt["etas"][si]
            qty = s.demand
            print(f"   {si+1:>2}. {s.name} | {int(qty)}瓶 | 到 {_hhmm(a)} 離 {_hhmm(lv)}")
        print()
    if skipped:
        print("跳過：", "; ".join(f"{n}({r})" for n, r in skipped))
    self_check(result, args)

    # 依執行日期分資料夾：REPORT_DIR/YYYY-MM-DD
    from datetime import datetime
    day_dir = os.path.join(REPORT_DIR, datetime.now().strftime("%Y-%m-%d"))
    os.makedirs(day_dir, exist_ok=True)

    html = report_mod.build_html_grouped(result, os.path.join(day_dir, "route_report.html"),
                                         meta={"start_hour": args.start})
    csvp = report_mod.build_csv_grouped(result, os.path.join(day_dir, "route_report.csv"))
    gmap = build_map(result, day_dir, use_google)
    # 路線圖截 PNG（本機有 Edge/Chrome 才成功；雲端無瀏覽器則回 None）
    from map_capture import capture_map_png
    map_png = capture_map_png(gmap, os.path.join(day_dir, "route_map.png"))
    print(f"\n📁 報表輸出資料夾：{day_dir}")
    print(f"📄 路線報表(HTML)：{html}")
    print(f"📄 路線報表(CSV) ：{csvp}")
    print(f"🌐 互動地圖      ：{gmap}")
    if map_png:
        print(f"🖼️ 路線圖PNG     ：{map_png}")

    # 派車單（每台車一份）→ 獨立 DISPATCH_DIR/日期/
    dispatch_dir = os.path.join(DISPATCH_DIR, datetime.now().strftime("%Y-%m-%d"))
    os.makedirs(dispatch_dir, exist_ok=True)
    # 把路線圖 PNG 也複製到 dispatch_dir（供同步器統一抓）
    if map_png:
        import shutil as _sh
        _sh.copy(map_png, os.path.join(dispatch_dir, "route_map.png"))
    dhtml, dcsv = report_mod.build_dispatch_grouped(result, dispatch_dir, meta={"start_hour": args.start})
    # 整合 Excel 預設不產（使用者決定：需時加 --excel）；LINE/雲端也不自動產
    xlsx = None
    if args.excel:
        xlsx = report_mod.build_workbook(result, os.path.join(dispatch_dir, "整合報表.xlsx"),
                                         meta={"start_hour": args.start},
                                         map_png=(os.path.join(dispatch_dir, "route_map.png") if map_png else None))
    print(f"🚚 派車單資料夾  ：{dispatch_dir}")
    print(f"🚚 派車單(HTML)  ：{dhtml}")
    print(f"🚚 派車單(CSV)   ：{dcsv}")
    if xlsx:
        print(f"📊 整合Excel     ：{xlsx}")
    else:
        print(f"📊 整合Excel     ：（未產出，需加 --excel 才產）")


def build_map(result, here, use_google):
    import json
    from osrm_client import get_route_geometry
    routes_geo = []
    # 路線顏色：依車數產 N 條，依序 紅/黃/藍（最多 3 台）
    ROUTE_COLORS = ["#e53935", "#fdd835", "#1e88e5"]  # 紅、黃、藍
    for i, rt in enumerate(result.routes):
        v = rt["vehicle"]
        color = ROUTE_COLORS[i % len(ROUTE_COLORS)]
        points = [(v.start_lat, v.start_lon)] + [(s.lat, s.lon) for s in rt["stops"]] + [(v.start_lat, v.start_lon)]
        # 真實道路折線：逐段向 OSRM 取幾何，失敗則退回直線連接
        line = get_route_geometry(points)
        if not line:
            line = [[a, b] for a, b in points]   # 退回直線
        stops_info = [{"name": s.name, "demand": s.demand, "addr": s.address} for s in rt["stops"]]
        routes_geo.append({"color": color, "vehicle": v.id,
                           "distance": round(rt["distance_km"], 1),
                           "load": round(rt["load"], 1),
                           "points": [[a, b] for a, b in points],   # 站點座標（標記用）
                           "line": [[a, b] for a, b in line],        # 道路折線（路線用）
                           "stops": stops_info})
    data_json = json.dumps({"depot": None, "routes": routes_geo}, ensure_ascii=False)
    out = _MAP_TPL.replace("/*__DATA__*/", data_json)
    p = os.path.join(here, "route_map.html")
    with open(p, "w", encoding="utf-8") as f:
        f.write(out)
    return p


_MAP_TPL = """<!DOCTYPE html>
<html lang="zh-TW"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>鮮奶配送路線地圖</title>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<script src="https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js"></script>
<style>html,body{margin:0;height:100%;font-family:-apple-system,"Microsoft JhengHei",sans-serif;}
#map{height:100%;width:100%;}
.panel{position:absolute;top:10px;right:10px;z-index:1000;background:#fff;padding:10px 12px;
border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,.2);max-height:90%;overflow:auto;min-width:230px;font-size:13px;}
.panel h3{margin:0 0 6px;font-size:14px;}
.leg{display:flex;align-items:center;gap:6px;margin:3px 0;}
.sw{width:14px;height:14px;border-radius:3px;flex:0 0 auto;}
.num-marker{background:#fff;border:2px solid #333;border-radius:50%;width:24px;height:24px;
  display:flex;align-items:center;justify-content:center;font-weight:800;font-size:13px;color:#111;
  box-shadow:0 1px 4px rgba(0,0,0,.5);}
.num-marker.start{background:#ffd700;}
.zooms{position:absolute;top:10px;left:10px;z-index:1000;display:flex;flex-direction:column;gap:6px;}
.zbtn{background:#1a73e8;color:#fff;border:none;border-radius:6px;padding:7px 10px;font-size:13px;
  font-weight:600;cursor:pointer;box-shadow:0 1px 4px rgba(0,0,0,.3);}
.zbtn:hover{background:#1557b0;}
</style></head>
<body><div id="map"></div>
<div class="zooms">
  <button class="zbtn" onclick="map.fitBounds(allBounds,{padding:[30,30]})">全覽路線</button>
  <button class="zbtn" onclick="if(firstStop)map.setView([firstStop[0],firstStop[1]],13)">聚焦起點</button>
  <button class="zbtn" style="background:#0b6b3a" onclick="dlMapPNG()">📥 下載地圖 PNG</button>
</div>
<div class="panel" id="panel"></div>
<script>
const DATA = /*__DATA__*/;
const map = L.map('map').setView([24.15, 120.67], 12);
L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {maxZoom:19, crossOrigin:true, attribution:'&copy; OpenStreetMap'}).addTo(map);
const panel = document.getElementById('panel');
function numIcon(label, isStart, color){
  return L.divIcon({className:'', html:'<div class="num-marker'+(isStart?' start':'')+'" style="'+(isStart?'':'background:'+color)+'">'+label+'</div>',
    iconSize:[24,24], iconAnchor:[12,12]});
}
let html = '<h3>🚚 路線清單（點標記看詳情）</h3>';
let firstStop = null;
let allBounds = null;
DATA.routes.forEach((r) => {
  L.polyline(r.line, {color: r.color, weight:4, opacity:.85}).addTo(map);
  const start = r.points[0];
  L.marker(start, {icon: numIcon('起', true, r.color)}).addTo(map)
    .bindPopup('<b>'+r.vehicle+' 起點（總倉）</b>');
  if (!firstStop) firstStop = start;
  r.stops.forEach((s, idx) => {
    const pos = r.points[idx+1];
    L.marker(pos, {icon: numIcon(String(idx+1), false, r.color)}).addTo(map)
      .bindPopup('<b>'+r.vehicle+' · 第 '+(idx+1)+' 站</b><br>'+s.name+'<br>'+s.addr+'<br>瓶數: '+s.demand);
  });
  html += '<div class="leg"><span class="sw" style="background:'+r.color+'"></span><span><b>'+r.vehicle+'</b> · '+r.distance+' km · '+r.load+' 瓶 · 共 '+r.stops.length+' 站</span></div>';
});
panel.innerHTML = '<h3>🚚 路線清單（顏色：紅=第1台 / 黃=第2台 / 藍=第3台）</h3>' + html.replace('<h3>🚚 路線清單（點標記看詳情）</h3>','');
const all = []; DATA.routes.forEach(r => r.points.forEach(p => all.push(p)));
if (all.length) {
  allBounds = L.latLngBounds(all);
  // 初始視圖：不強制全覽(會讓點擠在一起)，改設適中縮放並以起點為中心
  if (firstStop) map.setView([firstStop[0], firstStop[1]], 12);
}
function dlMapPNG(){
  var btns=document.querySelector('.zooms'); btns.style.visibility='hidden';
  setTimeout(function(){
    html2canvas(document.getElementById('map'),{useCORS:true,allowTaint:false,scale:2}).then(function(c){
      var a=document.createElement('a');
      a.download='配送地圖_'+new Date().toISOString().slice(0,10)+'.png';
      a.href=c.toDataURL('image/png'); a.click();
      btns.style.visibility='visible';
    }).catch(function(e){alert('產生地圖 PNG 失敗:'+e);btns.style.visibility='visible';});
  }, 600);
}
</script></body></html>"""


if __name__ == "__main__":
    main()
