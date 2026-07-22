"""
google_maps.py — Google Maps Platform 後端

功能：
  - geocode(address): 門牌級地理編碼 (Geocoding API)，台灣地址可精準到號
  - distance_matrix(coords): 真實道路距離矩陣(公里) + 行車時間矩陣(秒)
    (Distance Matrix API, driving 模式)

注意：
  - 需啟用帳單 + 開通 Geocoding API / Distance Matrix API
  - 讀取 .env 的 GOOGLE_MAPS_API_KEY
  - 失敗(REQUEST_DENIED / 配額)時拋出例外，由呼叫方降級 OSM
"""

import json
import os
import urllib.request
import urllib.parse

_BASE = "https://maps.googleapis.com/maps/api"

def haversine(lat1, lon1, lat2, lon2):
    """直線距離(公里)，cache_only 模式填 miss 格子用。"""
    import math
    R = 6371.0088
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlam = math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dlam / 2) ** 2
    return 2 * R * math.asin(math.sqrt(a))


def _load_key():
    key = os.environ.get("GOOGLE_MAPS_API_KEY")
    if key:
        return key
    # 從 .env 讀
    try:
        here = os.path.dirname(os.path.abspath(__file__))
        with open(os.path.join(here, ".env"), encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line.startswith("GOOGLE_MAPS_API_KEY="):
                    return line.split("=", 1)[1].strip().strip('"').strip("'")
    except Exception:
        pass
    raise RuntimeError("找不到 GOOGLE_MAPS_API_KEY（請在 .env 設定或匯出環境變數）")


def geocode(address: str):
    """門牌級地理編碼。回傳 (lat, lon) 或 None。失敗拋例外。"""
    if not address:
        return None
    key = _load_key()
    url = (f"{_BASE}/geocode/json?address={urllib.parse.quote(address)}"
           f"&region=tw&language=zh-TW&key={key}")
    req = urllib.request.Request(url, headers={"User-Agent": "logistics-agent/0.1"})
    data = json.loads(urllib.request.urlopen(req, timeout=25).read().decode())
    status = data.get("status")
    if status == "OK" and data.get("results"):
        loc = data["results"][0]["geometry"]["location"]
        return float(loc["lat"]), float(loc["lng"])
    if status == "ZERO_RESULTS":
        return None
    # REQUEST_DENIED / OVER_QUERY_LIMIT / 其他 → 拋出讓呼叫方降級
    raise RuntimeError(f"Google Geocoding 失敗: {status} {data.get('error_message','')}")


def distance_matrix(coords, timeout=30, batch=10, fast_fail=False, cache_only=False):
    """
    輸入 coords=[(lat,lon)...]，index 0 = depot。
    回傳 (matrix_km, duration_sec, source)。失敗拋例外。
    分塊呼叫：Google 單次上限 100 元素(origins×destinations)，
    故每塊 batch×batch 呼叫後拼回完整矩陣。

    timeout: 單一區塊最長等待秒數（預設 30，避免卡死）。
    fast_fail: 若為 True，遇到 OVER_QUERY_LIMIT / REQUEST_DENIED 立即拋出，
               不再對剩餘區塊做無謂嘗試（呼叫方會直接降級 OSRM/直線）。
    cache_only: 減 Google Cloud 費用硬原則——只讀 geo_cache，cache miss 的格子
                直接用 Haversine 直線填，絕不呼叫 Google。Render 環境自動強制。
    """
    # Render 環境：絕對不打 Google（減費用 + 避免對外連線卡死）
    if os.environ.get("RENDER") or os.environ.get("IS_RENDER"):
        cache_only = True

    key = _load_key()
    n = len(coords)
    if n <= 1:
        return [[0.0]], [[0.0]], "google"
    coord_str = [f"{lat:.6f},{lon:.6f}" for lat, lon in coords]

    matrix_km = [[0.0] * n for _ in range(n)]
    duration_sec = [[0.0] * n for _ in range(n)]

    # ---- 持久化快取：先看有沒有已存的兩點結果 ----
    # 全部命中 → 完全不呼叫 Google（$0）。只要有任何一對沒存過，才打 Google。
    # 符合「除非是沒用過的地址（新路線）才呼叫 Google」的需求。
    try:
        import geo_cache
        _use_cache = True
    except Exception:
        geo_cache = None
        _use_cache = False

    if _use_cache:
        all_hit = True
        for i in range(n):
            for j in range(n):
                if i == j:
                    continue
                pr = geo_cache.get_pair(coords[i][0], coords[i][1],
                                        coords[j][0], coords[j][1])
                if pr is None:
                    all_hit = False
                else:
                    matrix_km[i][j] = pr[0]
                    duration_sec[i][j] = pr[1]
        if all_hit:
            return matrix_km, duration_sec, "google-cache"
        if cache_only:
            # 減 Google 費用硬原則：絕不呼叫 Google。
            # cache miss 的格子直接用 Haversine 直線填（零費用，距離不精但必跑得完）。
            for i in range(n):
                for j in range(n):
                    if i == j:
                        continue
                    if matrix_km[i][j] == 0.0:
                        km = haversine(coords[i][0], coords[i][1],
                                       coords[j][0], coords[j][1])
                        matrix_km[i][j] = km
                        duration_sec[i][j] = km / 30.0 * 3600.0   # 30km/h 估速
            any_real = any(matrix_km[i][j] != 0.0
                            for i in range(n) for j in range(n) if i != j)
            src = "google-cache" if any_real else "haversine"
            return matrix_km, duration_sec, src

    for rs in range(0, n, batch):
        re = min(rs + batch, n)
        for cs in range(0, n, batch):
            ce = min(cs + batch, n)
            orig = "|".join(coord_str[rs:re])
            dest = "|".join(coord_str[cs:ce])
            url = (f"{_BASE}/distancematrix/json?units=metric&mode=driving"
                   f"&language=zh-TW&origins={urllib.parse.quote(orig)}"
                   f"&destinations={urllib.parse.quote(dest)}&key={key}")
            req = urllib.request.Request(url, headers={"User-Agent": "logistics-agent/0.1"})
            data = json.loads(urllib.request.urlopen(req, timeout=timeout).read().decode())
            status = data.get("status")
            if status != "OK":
                # 配額耗盡 / 權限問題：fast_fail 時立即拋出，避免後續區塊空等
                if fast_fail and status in ("OVER_QUERY_LIMIT", "REQUEST_DENIED"):
                    raise RuntimeError(
                        f"Google Distance Matrix 失敗: {status} {data.get('error_message','')}")
                raise RuntimeError(f"Google Distance Matrix 失敗: {status} {data.get('error_message','')}")
            for i, row in enumerate(data["rows"]):
                gi = rs + i
                for j, el in enumerate(row["elements"]):
                    gj = cs + j
                    if el["status"] == "OK":
                        matrix_km[gi][gj] = el["distance"]["value"] / 1000.0
                        duration_sec[gi][gj] = el["duration"]["value"]
                        # 存進持久化快取：下次同一對就不必再花錢
                        if _use_cache and gi != gj:
                            geo_cache.put_pair(
                                coords[gi][0], coords[gi][1],
                                coords[gj][0], coords[gj][1],
                                matrix_km[gi][gj], duration_sec[gi][gj])
    if _use_cache:
        geo_cache.flush()   # 一次寫回磁碟
    return matrix_km, duration_sec, "google"
