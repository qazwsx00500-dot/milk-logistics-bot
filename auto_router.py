"""
auto_router.py — 依「時間窗 + 最短距離」自動分配車輛路線

規則（來自使用者）：
  1. 每台車 09:30 出車、17:30 回倉（8 小時工作窗）
  2. 路線時間超過 → 自動多分配一台車
  3. 每條路線必須是最短行車距離（NN + 2-opt 已在 route_planner 做）
  4. 最多 3 條路線；一台跑不完才加下一台

作法：
  - 全部站點先估最短路線時間
  - 超時 → 用經緯度 k-means 拆成 2 群，每群各自最短路線
  - 仍超時 → 拆 3 群（最多 3 台）
  - 每群的「最短路線」由 route_planner.solve_grouped 負責
"""

import math
import random


def _haversine(lat1, lon1, lat2, lon2):
    R = 6371.0088
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlam = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dlam / 2) ** 2
    return 2 * R * math.asin(math.sqrt(a))


def kmeans_stops(stops, k, seed=42):
    """對 stops (有 .lat/.lon) 做 k-means 聚類，回傳 k 個群的索引分組。"""
    if k <= 1 or len(stops) <= k:
        return [list(range(len(stops)))]
    random.seed(seed)
    idxs = list(range(len(stops)))
    random.shuffle(idxs)
    centers = [(stops[i].lat, stops[i].lon) for i in idxs[:k]]
    for _ in range(50):
        groups = [[] for _ in range(k)]
        for si, s in enumerate(stops):
            best = min(range(k), key=lambda c: (s.lat - centers[c][0]) ** 2 + (s.lon - centers[c][1]) ** 2)
            groups[best].append(si)
        new_centers = []
        for g in groups:
            if g:
                new_centers.append((
                    sum(stops[i].lat for i in g) / len(g),
                    sum(stops[i].lon for i in g) / len(g),
                ))
            else:
                new_centers.append((stops[random.randrange(len(stops))].lat,
                                    stops[random.randrange(len(stops))].lon))
        if new_centers == centers:
            break
        centers = new_centers
    groups = [g for g in groups if g]
    return groups


def estimate_route_end_hour(stops, depot, start_hour=9.5, est_speed_kmh=30.0):
    """用 Haversine + 估速 + 下貨時間，粗估「全部站點一條最短路線」的回倉時刻。
    只用來決定「要不要加車」，不影響最終 route_planner 的精確路線。"""
    if not stops:
        return start_hour
    remaining = list(range(len(stops)))
    order = []
    cur = (depot.lat, depot.lon)
    t = start_hour * 3600.0
    while remaining:
        nxt = min(remaining, key=lambda i: _haversine(cur[0], cur[1], stops[i].lat, stops[i].lon))
        d_km = _haversine(cur[0], cur[1], stops[nxt].lat, stops[nxt].lon)
        t += d_km / est_speed_kmh * 3600.0 + (getattr(stops[nxt], "service_time", 0) or 0)
        order.append(nxt)
        cur = (stops[nxt].lat, stops[nxt].lon)
        remaining.remove(nxt)
    d_back = _haversine(cur[0], cur[1], depot.lat, depot.lon)
    t += d_back / est_speed_kmh * 3600.0
    return t / 3600.0


def decide_groups(stops, depot, start_hour=9.5, target_return=17.5, max_vehicles=3):
    """依時間窗決定分幾群。回傳 list of list（每群是 stop 索引）。
    原則：先試 1 台；超時才加車，最多 max_vehicles 台。"""
    end1 = estimate_route_end_hour(stops, depot, start_hour)
    if end1 <= target_return or max_vehicles <= 1:
        return [list(range(len(stops)))]
    for k in range(2, max_vehicles + 1):
        groups = kmeans_stops(stops, k)
        ok = True
        for g in groups:
            sub = [stops[i] for i in g]
            if estimate_route_end_hour(sub, depot, start_hour) > target_return:
                ok = False
                break
        if ok:
            return groups
    return kmeans_stops(stops, max_vehicles)
