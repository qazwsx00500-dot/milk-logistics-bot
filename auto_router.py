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


def decide_groups(stops, depot, start_hour=9.5, target_return=17.5, max_vehicles=3,
                  precomputed_end_hour=None, est_speed_kmh=30.0):
    """依時間窗決定分幾群。回傳 list of list（每群是 stop 索引）。
    原則：先試 1 台；超時才加車，最多 max_vehicles 台。

    precomputed_end_hour: 若提供(來自 solve_grouped 的真實 end_hour)，
        用它判斷「1 台車是否超時」，避免 haversine 粗估低估導致不分車。
    """
    if precomputed_end_hour is not None:
        one_over = precomputed_end_hour > target_return
    else:
        one_over = estimate_route_end_hour(stops, depot, start_hour, est_speed_kmh) > target_return

    if not one_over or max_vehicles <= 1:
        return [list(range(len(stops)))]
    for k in range(2, max_vehicles + 1):
        groups = kmeans_stops(stops, k)
        ok = True
        for g in groups:
            sub = [stops[i] for i in g]
            if estimate_route_end_hour(sub, depot, start_hour, est_speed_kmh) > target_return:
                ok = False
                break
        if ok:
            return groups
    return kmeans_stops(stops, max_vehicles)


def _centroid(group, stops):
    if not group:
        return (0.0, 0.0)
    lat = sum(stops[i].lat for i in group) / len(group)
    lon = sum(stops[i].lon for i in group) / len(group)
    return (lat, lon)


def _rebalance(groups, stops, depot, start_hour, target_return,
               matrix=None, duration=None, est_speed_kmh=30.0, max_iter=400):
    """全局貪心搬運：把站點從超時群攤到還有餘裕的群，直到每群都準時。

    每輪：
      1. 把所有「超時群」的站攤平——對每個超時群裡的每個站，找「剩餘時間最多、
         且接收後仍準時、且重心最近」的目標群搬過去。
      2. 一輪內儘可能多搬，直到沒有超時群能再倒給任何有餘裕群。
    直接操作全域 stop 索引。
    """
    def end_of(g):
        if not g:
            return start_hour
        if matrix is not None and duration is not None:
            # g 是全域 stop 索引；coords 中 0=depot, 1..n=站，故站索引 si 對應矩陣 (si+1)
            local = [0] + [i + 1 for i in g]
            k = len(g)
            km = {}; dur = {}
            for li in range(k + 1):
                for lj in range(k + 1):
                    a = local[li]; b = local[lj]
                    km[("c", li, lj)] = matrix[(a, b)]
                    dur[("c", li, lj)] = duration[(a, b)]
            # 用真實矩陣重算該群最短路線時間
            from route_planner import Vehicle, solve_grouped
            veh = [Vehicle(id="c", name="c", start_lat=depot.lat,
                           start_lon=depot.lon, start_addr=depot.address)]
            subs = [stops[i] for i in g]
            res = solve_grouped(veh, {"c": subs}, matrix_km=km, duration_matrix=dur,
                                distance_source="real", start_hour=start_hour,
                                fuel_cost_per_km=0.0)
            return res.routes[0]["end_hour"]
        return estimate_route_end_hour([stops[i] for i in g], depot, start_hour, est_speed_kmh)

    def remaining(g):
        return target_return - end_of(g)

    for _ in range(max_iter):
        overs = [gi for gi, g in enumerate(groups) if end_of(g) > target_return]
        if not overs:
            break
        moved_any = False
        for gi in overs:
            g = groups[gi]
            if not g:
                continue
            # 找「剩餘時間最多」的目標群（且有餘裕）
            cands = sorted(((remaining(groups[tj]), tj) for tj in range(len(groups))
                            if tj != gi and remaining(groups[tj]) > 0), reverse=True)
            if not cands:
                continue
            # 在超時群 g 裡，找離任一候選群重心最近、且搬過去仍準時的站
            best = None
            for tj_rem, tj in cands:
                cd = _centroid(groups[tj], stops)
                for si in g:
                    if end_of(groups[tj] + [si]) <= target_return:
                        dist = _haversine(stops[si].lat, stops[si].lon, cd[0], cd[1])
                        # 評分：目標群餘裕越多、距離越近 越好
                        score = tj_rem - dist * 0.01
                        if best is None or score > best[0]:
                            best = (score, si, tj)
            if best is not None:
                _, si, tj = best
                groups[gi].remove(si)
                groups[tj].append(si)
                moved_any = True
        if not moved_any:
            break
    return groups


def balanced_groups(stops, depot, start_hour=9.5, target_return=17.5, max_vehicles=3,
                    matrix=None, duration=None, est_speed_kmh=30.0):
    """時間窗驅動的均衡分車：儘量讓每群 ≤ 目標回倉時間。

    先試 k=2、再 k=3（最多 max_vehicles 台）。每個 k 跑多個 seed 的 k-means，
    各自 _rebalance，選「所有群準時 且 最忙群最早回倉」的最佳解；
    若都無法全準時，選「超時群數最少、超時幅度最小」的解。
    回傳 list of list（全域 stop 索引），最少車數優先。
    """
    if not stops:
        return []
    # 1 台就準時 → 直接一台
    if estimate_route_end_hour(stops, depot, start_hour, est_speed_kmh) <= target_return:
        return [list(range(len(stops)))]
    if max_vehicles <= 1:
        return [list(range(len(stops)))]

    def _score(groups):
        """越低越好。優先：全準時 > 超時群少 > 超時幅度小 > 最忙群早回倉。"""
        if matrix is not None:
            ends = [group_end_hour_real(g, depot, start_hour, matrix, duration, stops, est_speed_kmh)
                    for g in groups]
        else:
            ends = [estimate_route_end_hour([stops[i] for i in g], depot, start_hour, est_speed_kmh)
                    for g in groups]
        over = sum(1 for e in ends if e > target_return)
        over_amt = sum(e - target_return for e in ends if e > target_return)
        max_end = max(ends)
        # 全準時時 max_end 越小越好；有超時時 over/over_amt 主導
        return (over, round(over_amt, 3), round(max_end, 3))

    def _try_k(k):
        best_g = None
        best_s = None
        for seed in range(8):  # 多 seed 避免壞初始(如 2/56/5)
            groups = kmeans_stops(stops, k, seed=seed)
            groups = _rebalance(groups, stops, depot, start_hour, target_return,
                                matrix=matrix, duration=duration, est_speed_kmh=est_speed_kmh)
            s = _score(groups)
            if best_s is None or s < best_s:
                best_s = s; best_g = groups
        return best_g, best_s

    for k in range(2, max_vehicles + 1):
        g, s = _try_k(k)
        # 全準時(over==0)就接受，且 k 最小優先
        if s[0] == 0:
            return g
    # 都無法全準時：回傳 k=max_vehicles 的最佳均衡解（至少盡量均衡、標超時）
    g, s = _try_k(max_vehicles)
    return g


def group_end_hour_real(group, depot, start_hour, matrix, duration, stops, est_speed_kmh=30.0):
    """用真實矩陣算某群最短路線回倉時刻。group=全域 stop 索引 list；stops=全部 Stop 列表。"""
    if not group:
        return start_hour
    from route_planner import Vehicle, solve_grouped
    k = len(group)
    local = [0] + [i + 1 for i in group]   # 站索引 si 對應矩陣 (si+1)
    km = {}; dur = {}
    for li in range(k + 1):
        for lj in range(k + 1):
            a = local[li]; b = local[lj]
            km[("c", li, lj)] = matrix[(a, b)]
            dur[("c", li, lj)] = duration[(a, b)]
    veh = [Vehicle(id="c", name="c", start_lat=depot.lat,
                   start_lon=depot.lon, start_addr=depot.address)]
    subs = [stops[i] for i in group]
    res = solve_grouped(veh, {"c": subs}, matrix_km=km, duration_matrix=dur,
                        distance_source="real", start_hour=start_hour,
                        fuel_cost_per_km=0.0)
    return res.routes[0]["end_hour"]

