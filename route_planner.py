"""
route_planner.py — 物流路線「分組順序規劃」核心

純 Python，零第三方依賴（地理編碼/距離靠外部 client）。
模式：照「車號」分組，每台車獨立把配送點排出最順拜訪順序。
     每台車有各自的「出發點/起點地址」，從起點出發、最後回起點。
     不做容量限制（使用者自行決定載貨）。

距離來源可注入：
  - 預設 Haversine 直線
  - 或由呼叫方傳入距離/時間矩陣 (用真實道路)

用法見 logistics_agent.py
"""

import math
from dataclasses import dataclass, field


@dataclass
class Stop:
    id: str
    name: str
    lat: float
    lon: float
    demand: float = 0.0          # 瓶數 (鮮乳類)
    service_time: float = 0.0    # 下貨秒數 (= 瓶數*15 + 有非鮮奶品項? 180 : 0)
    address: str = ""
    vehicle: str = ""            # 所屬車號
    items: dict = None           # 非鮮奶品項: {品名: {"qty": float, "unit": str}}
                                  #   (鮮乳已計入 demand，不在此)
    # ── 特殊需求約束（Excel「特殊需求」欄解析）──────────────
    constraint: dict = None      # {"time_lb": float|None, "time_ub": float|None,
                                  #   "first": bool, "last": bool, "raw": str}
    def __post_init__(self):
        if self.items is None:
            self.items = {}
        if self.constraint is None:
            self.constraint = {}


@dataclass
class Vehicle:
    id: str
    name: str
    start_lat: float
    start_lon: float
    start_addr: str = ""


@dataclass
class PlanResult:
    routes: list = field(default_factory=list)   # 每台車一條
    unassigned: list = field(default_factory=list)
    total_distance_km: float = 0.0
    total_load: float = 0.0
    total_fuel_cost: float = 0.0     # 總油資成本(元)
    fuel_cost_per_km: float = 0.0    # 每公里油資(元/km)
    summary: str = ""
    distance_source: str = "haversine"


def haversine(lat1, lon1, lat2, lon2):
    R = 6371.0088
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlam = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dlam / 2) ** 2
    return 2 * R * math.asin(math.sqrt(a))


class _Dist:
    def __init__(self, nodes, matrix_km=None, duration_matrix=None, est_speed_kmh=30.0):
        self.nodes = nodes
        self.matrix = matrix_km
        self.dur = duration_matrix
        self.est_speed = est_speed_kmh

    def d(self, i, j):
        if self.matrix is not None:
            return self.matrix[i][j]
        a, b = self.nodes[i], self.nodes[j]
        return haversine(a.lat, a.lon, b.lat, b.lon)

    def t(self, i, j):
        if self.dur is not None:
            return self.dur[i][j]
        return self.d(i, j) / self.est_speed * 3600.0


def _route_duration_sec(dist, start_idx, stop_idxs):
    if not stop_idxs:
        return 0.0
    seq = [start_idx] + list(stop_idxs) + [start_idx]
    total = 0.0
    for a, b in zip(seq, seq[1:]):
        total += dist.t(a, b)
    for k in stop_idxs:
        total += getattr(dist.nodes[k], "service_time", 0) or 0.0
    return total


def _edge_cost(dist, a, b, region_fn=None, lam=0.0):
    """單段成本。lam>0 時跨區段加權懲罰：成本 = t(a,b)*(1+lam)。
    lam=0 → 純行車時間（等同 2-opt）；lam 大 → 強聚簇。中間值即平衡點。"""
    base = dist.t(a, b)
    if lam and region_fn is not None:
        ra = region_fn(dist.nodes[a].address)
        rb = region_fn(dist.nodes[b].address)
        if ra != rb:
            return base * (1.0 + lam)
    return base


def _route_cost_sec(dist, start_idx, stop_idxs, region_fn=None, lam=0.0):
    """含區域懲罰的總成本（供 2-opt 比較用）。"""
    if not stop_idxs:
        return 0.0
    seq = [start_idx] + list(stop_idxs) + [start_idx]
    total = 0.0
    for a, b in zip(seq, seq[1:]):
        total += _edge_cost(dist, a, b, region_fn, lam)
    for k in stop_idxs:
        total += getattr(dist.nodes[k], "service_time", 0) or 0.0
    return total


def _nearest_neighbor(dist, start_idx, stop_idxs, region_fn=None, lam=0.0):
    remaining = list(stop_idxs)
    route = []
    current = start_idx
    while remaining:
        nxt = min(remaining, key=lambda k: _edge_cost(dist, current, k, region_fn, lam))
        route.append(nxt)
        remaining.remove(nxt)
        current = nxt
    return route


def _two_opt(dist, start_idx, stop_idxs, max_iter=200, region_fn=None, lam=0.0):
    route = list(stop_idxs)
    if len(route) < 4:
        return route
    best = _route_cost_sec(dist, start_idx, route, region_fn, lam)
    improved = True
    it = 0
    while improved and it < max_iter:
        improved = False
        it += 1
        for i in range(len(route) - 1):
            for j in range(i + 1, len(route)):
                new = route[:i] + route[i:j + 1][::-1] + route[j + 1:]
                d = _route_cost_sec(dist, start_idx, new, region_fn, lam)
                if d + 1e-9 < best:
                    route, best = new, d
                    improved = True
    return route


def _multi_restart_two_opt(dist, start_idx, stop_idxs, n_restart=8, max_iter=200,
                           seed=12345, region_fn=None, lam=0.0):
    """multi-restart 2-opt：多個隨機起點各跑 NN+2-opt，取成本最小者。
    站數 <= 3 直接回傳。lam>0 時跨區段加權（region_lambda 平衡機制）。"""
    if len(stop_idxs) <= 3:
        return list(stop_idxs)
    import random
    rng = random.Random(seed)
    best_route = None
    best_dur = float("inf")
    for r in range(n_restart):
        if r == 0:
            cand = _nearest_neighbor(dist, start_idx, list(stop_idxs), region_fn, lam)
        else:
            shuffled = list(stop_idxs)
            rng.shuffle(shuffled)
            cand = _nearest_neighbor(dist, start_idx, shuffled, region_fn, lam)
        cand = _two_opt(dist, start_idx, cand, max_iter=max_iter, region_fn=region_fn, lam=lam)
        dur = _route_cost_sec(dist, start_idx, cand, region_fn, lam)
        if dur + 1e-9 < best_dur:
            best_route, best_dur = cand, dur
    return best_route


def _constraint_sort(stops, ordered, start_hour, dist, start_idx=0):
    """把 2-opt 排序後的路線，依『特殊需求』約束調整順序。
    - 首站(first): 移到最前
    - 末站(last): 移到最後
    - 時間窗(time_lb/time_ub): 軟約束，把早送(時間窗上界小)往前挪、晚送(下界大)往後挪
      用插入排序：依 (time_ub or time_lb or 99) 升序盡量排，但不破壞太嚴重（仍貪心插入到合法位）
    回傳 (new_ordered, violations)；violations 列出仍違反時間窗的站名。
    """
    stops_map = {i: stops[i - 1] for i in ordered}  # node index -> Stop
    # 分類
    firsts, lasts, mids = [], [], []
    for k in ordered:
        c = getattr(stops_map[k], "constraint", None) or {}
        if c.get("first"):
            firsts.append(k)
        elif c.get("last"):
            lasts.append(k)
        else:
            mids.append(k)
    # 中間段按時間窗上界升序（早送排前面）。無時間窗者用大值排後。
    def _key(k):
        c = getattr(stops_map[k], "constraint", None) or {}
        ub = c.get("time_ub")
        lb = c.get("time_lb")
        # 早送(time_ub 小) 排最前(group 0)；無約束居中(group 1)；晚送(time_lb 大) 排最後(group 2)
        if ub is not None:
            return (0, ub)
        if lb is not None:
            return (2, lb)
        return (1, 99.0)
    mids.sort(key=_key)
    new_ordered = firsts + mids + lasts

    # 檢查違反時間窗
    violations = []
    t = start_hour
    prev = start_idx
    # 用同一 dist 重算 ETA 檢查
    for k in new_ordered:
        t += dist.t(prev, k) / 3600.0
        arrive = t
        c = getattr(stops_map[k], "constraint", None) or {}
        if c.get("time_ub") is not None and arrive > c["time_ub"] + 1/60.0:
            violations.append((stops_map[k].name,
                               f"應 {_hhmm(c['time_ub'])}前到，實際 {_hhmm(arrive)}"))
        if c.get("time_lb") is not None and arrive < c["time_lb"] - 1/60.0:
            violations.append((stops_map[k].name,
                               f"應 {_hhmm(c['time_lb'])}後到，實際 {_hhmm(arrive)}"))
        svc = getattr(stops_map[k], "service_time", 0) or 0.0
        t += svc / 3600.0
        prev = k
    return new_ordered, violations


    return best_route


def _region_of(addr):
    """從地址解析『區/鎮/市』層級的聚簇 key（比縣市更細，符合『一區跑完再下一區』）。
    例：台中市西屯區中工三路 → '西屯區'；苗栗縣竹南鎮新生路 → '竹南鎮'。
    抓不到區則退回前兩段（縣市+剩餘）避免全併成一團。"""
    import re
    a = addr or ""
    # 先去掉縣市前綴（台中市/臺中市/苗栗縣...），再抓區/鎮/市
    a2 = re.sub(r"^(臺[北中高][市縣]|台[北中高][市縣]|苗栗縣|彰化縣|雲林縣|嘉義縣|南投縣)", "", a)
    # 區（如 西屯區/北區/東區/太平區/烏日區/豐原區）
    m = re.search(r"(.+?區)", a2)
    if m:
        return m.group(1)
    # 鎮 / 市（如 竹南鎮/頭份市/苑裡鎮/卓蘭鎮/通霄鎮）
    m = re.search(r"(.+?鎮|.+?市)", a2)
    if m:
        return m.group(1)
    # 兜底：前兩段
    parts = [p for p in re.split(r"[市縣]", a) if p]
    return parts[0][:4] if parts else a[:6]


def solve_grouped_regional(
    vehicles: list[Vehicle],
    stops_by_vehicle: dict,
    matrix_km: dict = None,
    duration_matrix: dict = None,
    distance_source: str = "haversine",
    start_hour: float = 8.0,
    fuel_cost_per_km: float = 0.0,
    region_lambda: float = 0.0,
) -> PlanResult:
    """區域平衡版：全站用 multi-restart 2-opt，但跨區段邊成本加權 region_lambda。
    λ=0 → 純 2-opt（最短距離、可能跳區）；λ 大 → 強聚簇（不跳區、里程增）；
    中間值 → 平衡點（不跳區且總里程不暴增）。產出結構與 solve_grouped 一致。"""
    result = PlanResult()
    result.distance_source = distance_source
    result.fuel_cost_per_km = fuel_cost_per_km
    total_dist = 0.0
    total_load = 0.0
    total_fuel = 0.0

    for v in vehicles:
        stops = stops_by_vehicle.get(v.id, [])
        if not stops:
            continue
        nodes = [Stop("START", v.start_addr or "起點", v.start_lat, v.start_lon)] + list(stops)
        start_idx = 0
        stop_idxs = list(range(1, len(nodes)))

        # 該車矩陣
        m_km = None
        m_dur = None
        if matrix_km is not None:
            n = len(stops)
            m_km = [[0.0] * (n + 1) for _ in range(n + 1)]
            for i in range(n + 1):
                for j in range(n + 1):
                    m_km[i][j] = matrix_km.get((v.id, i, j), 0.0)
        if duration_matrix is not None:
            n = len(stops)
            m_dur = [[0.0] * (n + 1) for _ in range(n + 1)]
            for i in range(n + 1):
                for j in range(n + 1):
                    m_dur[i][j] = duration_matrix.get((v.id, i, j), 0.0)
        dist = _Dist(nodes, m_km, m_dur)

        # --- 區域平衡：全站用 2-opt，但跨區段邊成本加權 region_lambda ---
        # λ=0 → 純 2-opt（最短距離，可能跳區）
        # λ 大 → 跨區移動被懲罰，趨向「同區跑完再下一區」
        # 中間值 → 平衡點（不跳區且總里程不暴增）
        ordered = _multi_restart_two_opt(
            dist, start_idx, stop_idxs, region_fn=_region_of, lam=region_lambda)

        # 套用特殊需求約束
        ordered, violations = _constraint_sort(stops, ordered, start_hour, dist, start_idx)

        # 距離 / ETA 計算（與 solve_grouped 同）
        dist_km = 0.0
        seq = [start_idx] + ordered + [start_idx]
        for a, b in zip(seq, seq[1:]):
            dist_km += dist.d(a, b)
        etas = []
        t = start_hour
        prev = start_idx
        for k in ordered:
            t += dist.t(prev, k) / 3600.0
            arrive = t
            svc = getattr(nodes[k], "service_time", 0) or 0.0
            t += svc / 3600.0
            etas.append((arrive, t))
            prev = k
        t += dist.t(prev, start_idx) / 3600.0
        end_hour = t

        load = sum(s.demand for s in stops)
        fuel = dist_km * fuel_cost_per_km
        total_dist += dist_km
        total_load += load
        total_fuel += fuel
        result.routes.append({
            "vehicle": v,
            "stops": [stops[k - 1] for k in ordered],
            "etas": etas,
            "distance_km": dist_km,
            "end_hour": end_hour,
            "load": load,
            "fuel_cost": fuel,
            "violations": violations,
        })

    result.routes.sort(key=lambda r: r["vehicle"].id)
    result.total_distance_km = total_dist
    result.total_load = total_load
    result.total_fuel_cost = total_fuel
    result.summary = _make_summary(result)
    return result


def solve_grouped(
    vehicles: list[Vehicle],
    stops_by_vehicle: dict,           # {車號: [Stop...]}
    matrix_km: dict = None,           # {(v_id, i, j): km} 或不傳
    duration_matrix: dict = None,     # {(v_id, i, j): sec}
    distance_source: str = "haversine",
    start_hour: float = 8.0,
    fuel_cost_per_km: float = 0.0,   # 每公里油資(元/km)
) -> PlanResult:
    """
    每台車獨立排序。stops_by_vehicle[車號] 內的 Stop 已地理編碼。
    matrix_km/duration_matrix 為選用：{(車號, i, j): 值}，i/j 為該車 stops 的 index。
    若提供，則用真實距離/時間；否則 Haversine+估速。
    fuel_cost_per_km: 每公里油資(元)，用於估算油資成本。
    """
    result = PlanResult()
    result.distance_source = distance_source
    result.fuel_cost_per_km = fuel_cost_per_km
    total_dist = 0.0
    total_load = 0.0
    total_fuel = 0.0

    for v in vehicles:
        stops = stops_by_vehicle.get(v.id, [])
        if not stops:
            continue
        # 節點：index 0 = 起點，1..n = stops
        nodes = [Stop("START", v.start_addr or "起點", v.start_lat, v.start_lon)] + list(stops)
        start_idx = 0
        stop_idxs = list(range(1, len(nodes)))

        # 取該車矩陣
        m_km = None
        m_dur = None
        if matrix_km is not None:
            n = len(stops)
            m_km = [[0.0] * (n + 1) for _ in range(n + 1)]
            for i in range(n + 1):
                for j in range(n + 1):
                    m_km[i][j] = matrix_km.get((v.id, i, j), 0.0)
        if duration_matrix is not None:
            n = len(stops)
            m_dur = [[0.0] * (n + 1) for _ in range(n + 1)]
            for i in range(n + 1):
                for j in range(n + 1):
                    m_dur[i][j] = duration_matrix.get((v.id, i, j), 0.0)

        dist = _Dist(nodes, m_km, m_dur)
        ordered = _multi_restart_two_opt(dist, start_idx, stop_idxs)
        # 套用特殊需求約束（首站/末站硬約束 + 時間窗軟排序）
        ordered, violations = _constraint_sort(stops, ordered, start_hour, dist, start_idx)

        # 距離
        dist_km = 0.0
        seq = [start_idx] + ordered + [start_idx]
        for a, b in zip(seq, seq[1:]):
            dist_km += dist.d(a, b)

        # ETA
        etas = []
        t = start_hour
        prev = start_idx
        for k in ordered:
            t += dist.t(prev, k) / 3600.0
            arrive = t
            svc = getattr(nodes[k], "service_time", 0) or 0.0
            t += svc / 3600.0
            etas.append((arrive, t))
            prev = k
        t += dist.t(prev, start_idx) / 3600.0
        end_hour = t

        load = sum(s.demand for s in stops)
        fuel = dist_km * fuel_cost_per_km
        total_dist += dist_km
        total_load += load
        total_fuel += fuel
        result.routes.append({
            "vehicle": v,
            "stops": [stops[k - 1] for k in ordered],   # 轉回 Stop
            "etas": etas,
            "distance_km": dist_km,
            "end_hour": end_hour,
            "load": load,
            "fuel_cost": fuel,
            "violations": violations,   # 特殊需求時間窗未達成清單
        })

    result.routes.sort(key=lambda r: r["vehicle"].id)
    result.total_distance_km = total_dist
    result.total_load = total_load
    result.total_fuel_cost = total_fuel
    result.summary = _make_summary(result)
    return result


def _make_summary(result):
    src_map = {"haversine": "直線估算", "osrm": "OSRM 真實道路",
               "google": "Google Maps 真實道路",
               "fallback": "直線估算(降級)"}
    src = src_map.get(result.distance_source, result.distance_source)
    lines = []
    fpk = result.fuel_cost_per_km
    lines.append(f"出車數：{len(result.routes)} 台")
    lines.append(f"距離來源：{src}")
    lines.append(f"總配送距離：{result.total_distance_km:.1f} km")
    lines.append(f"總瓶數：{result.total_load:.0f}")
    if fpk > 0:
        lines.append(f"油資單價：{fpk:.1f} 元/km")
        lines.append(f"預估總油資：{result.total_fuel_cost:.0f} 元")
    for i, rt in enumerate(result.routes, 1):
        v = rt["vehicle"]
        fuel_txt = f" / 油資 {rt.get('fuel_cost', 0):.0f} 元" if fpk > 0 else ""
        lines.append(f"  {v.id}：{len(rt['stops'])} 站 / {rt['distance_km']:.1f} km / "
                     f"{rt['load']:.0f} 瓶 / 回到起點 {_hhmm(rt['end_hour'])}{fuel_txt}")
    return "\n".join(lines)


def _hhmm(h):
    hh = int(h) % 24
    mm = int(round((h - int(h)) * 60))
    if mm == 60:
        hh += 1; mm = 0
    return f"{hh:02d}:{mm:02d}"
