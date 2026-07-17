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
    demand: float = 0.0          # 瓶數
    service_time: float = 0.0    # 下貨秒數 (= 瓶數 * 15)
    address: str = ""
    vehicle: str = ""            # 所屬車號


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


def _nearest_neighbor(dist, start_idx, stop_idxs):
    remaining = list(stop_idxs)
    route = []
    current = start_idx
    t_acc = 0.0
    while remaining:
        def cost(k):
            return dist.t(current, k)
        nxt = min(remaining, key=cost)
        route.append(nxt)
        t_acc += dist.t(current, nxt) + (getattr(dist.nodes[nxt], "service_time", 0) or 0)
        remaining.remove(nxt)
        current = nxt
    return route


def _two_opt(dist, start_idx, stop_idxs, max_iter=200):
    route = list(stop_idxs)
    if len(route) < 4:
        return route
    best = _route_duration_sec(dist, start_idx, route)
    improved = True
    it = 0
    while improved and it < max_iter:
        improved = False
        it += 1
        for i in range(len(route) - 1):
            for j in range(i + 1, len(route)):
                new = route[:i] + route[i:j + 1][::-1] + route[j + 1:]
                d = _route_duration_sec(dist, start_idx, new)
                if d + 1e-9 < best:
                    route, best = new, d
                    improved = True
    return route


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
        ordered = _nearest_neighbor(dist, start_idx, stop_idxs)
        ordered = _two_opt(dist, start_idx, ordered)

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
