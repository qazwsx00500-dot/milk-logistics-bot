"""
osrm_client.py — 真實道路距離客戶端

使用 OSRM 公開路由服務 (https://router.project-osrm.org)，
免 API Key、含台灣路網資料。

功能：
  - get_distance_matrix(coords): 一次算出所有點對的真實道路距離矩陣 (公里)
  - get_route_geometry(coords):  取得一連串點的真實道路折線 (供地圖繪製)
  - 失敗時自動降級為 Haversine 直線距離/直線連接，並標註來源

座標格式： [(lat, lon), ...]
"""

import json
import math
import urllib.request
import urllib.error

OSRM_BASE = "https://router.project-osrm.org"

# 分類錯誤：是否值得重試
_RETRYABLE = (urllib.error.URLError, TimeoutError, ConnectionError)


def haversine(lat1, lon1, lat2, lon2):
    R = 6371.0088
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlam = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dlam / 2) ** 2
    return 2 * R * math.asin(math.sqrt(a))


def _post_json(url, payload=None, timeout=40):
    data = json.dumps(payload).encode() if payload is not None else None
    req = urllib.request.Request(
        url, data=data,
        headers={"User-Agent": "logistics-agent/0.1", "Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode())


def get_matrix_dur(coords, timeout=60):
    """
    同時取得距離矩陣(公里) 與 行車時間矩陣(秒)。
    回傳 (matrix_km, duration_sec, source)
      source = 'osrm' 或 'fallback'(直線/估速)
    """
    n = len(coords)
    def fallback():
        m = [[0.0] * n for _ in range(n)]
        d = [[0.0] * n for _ in range(n)]
        for i in range(n):
            for j in range(i + 1, n):
                km = haversine(coords[i][0], coords[i][1], coords[j][0], coords[j][1])
                m[i][j] = m[j][i] = km
                d[i][j] = d[j][i] = km / 30.0 * 3600.0   # 30km/h 估速
        return m, d, "fallback"

    if n <= 1:
        return [[0.0]], [[0.0]], "fallback"

    coord_str = ";".join(f"{lon:.6f},{lat:.6f}" for lat, lon in coords)
    url = (f"{OSRM_BASE}/table/v1/driving/{coord_str}"
           f"?annotations=distance,duration")
    try:
        data = _post_json(url, timeout=timeout)
        if data.get("code") != "Ok":
            return fallback()
        matrix_km = [[x / 1000.0 for x in row] for row in data["distances"]]
        duration_sec = [[x for x in row] for row in data["durations"]]
        return matrix_km, duration_sec, "osrm"
    except _RETRYABLE:
        return fallback()
    except Exception:
        return fallback()


def get_distance_matrix(coords, timeout=60):
    """
    輸入 coords: [(lat, lon), ...]，index 0 通常為倉庫。
    回傳 (matrix_km: list[list[float]], source: str)
      - matrix_km[i][j] = 從點 i 到點 j 的真實道路距離 (公里)
      - source = 'osrm' 或 'fallback'(直線估算)
    """
    n = len(coords)
    # 失敗降級用直線
    def fallback():
        m = [[0.0] * n for _ in range(n)]
        for i in range(n):
            for j in range(i + 1, n):
                d = haversine(coords[i][0], coords[i][1], coords[j][0], coords[j][1])
                m[i][j] = m[j][i] = d
        return m, "fallback"

    if n == 0:
        return [], "fallback"
    if n == 1:
        return [[0.0]], "osrm"

    # OSRM 期望 lon,lat
    coord_str = ";".join(f"{lon:.6f},{lat:.6f}" for lat, lon in coords)
    url = (f"{OSRM_BASE}/table/v1/driving/{coord_str}"
           f"?annotations=distance")
    try:
        data = _post_json(url, timeout=timeout)
        if data.get("code") != "Ok":
            return fallback()
        # 距離單位為公尺 → 轉公里
        dists = data["distances"]
        matrix_km = [[d / 1000.0 for d in row] for row in dists]
        return matrix_km, "osrm"
    except _RETRYABLE:
        return fallback()
    except Exception:
        return fallback()


def get_route_geometry(ordered_coords, timeout=40, max_pts=50):
    """
    輸入 ordered_coords: [(lat, lon), ...] 已排好順序 (含起終點倉庫)。
    回傳 list[[lat, lon], ...] 的真實道路折線，或 None（失敗時由呼叫方退回直線）。

    做法：逐段 (點i -> 點i+1) 各呼叫一次 OSRM route 拿該段道路折線，再串接。
          比「一次丟全部點」穩（OSRM route 一次 via 點過多會 400 / 路徑變形）。
          max_pts: 若點數 <= max_pts 仍用一次呼叫（更快），否則才逐段。
    """
    if len(ordered_coords) < 2:
        return None
    # 點數不多就一次呼叫（OSRM 免費服務可承受 ~50 個 via 點）
    if len(ordered_coords) <= max_pts:
        return _osrm_route_once(ordered_coords, timeout)
    # 逐段
    full = []
    for a, b in zip(ordered_coords, ordered_coords[1:]):
        seg = _osrm_route_once([a, b], timeout)
        if seg is None:
            # 這段失敗：用直線兩點代替，保證地圖不斷
            full.append(list(a))
        else:
            # 避免重複銜接點
            if full:
                full.extend(seg[1:])
            else:
                full.extend(seg)
    return full if len(full) >= 2 else None


def _osrm_route_once(coords, timeout=40):
    """一次 OSRM route 呼叫，回傳 [[lat,lon],...] 或 None。"""
    if len(coords) < 2:
        return None
    coord_str = ";".join(f"{lon:.6f},{lat:.6f}" for lat, lon in coords)
    url = (f"{OSRM_BASE}/route/v1/driving/{coord_str}"
           f"?overview=full&geometries=geojson")
    try:
        data = _post_json(url, timeout=timeout)
        if data.get("code") != "Ok" or not data.get("routes"):
            return None
        geom = data["routes"][0]["geometry"]["coordinates"]
        # geojson 是 [lon, lat] → 轉 [lat, lon]
        return [[lat, lon] for lon, lat in geom]
    except Exception:
        return None
