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


def distance_matrix(coords, timeout=90, batch=10):
    """
    輸入 coords=[(lat,lon)...]，index 0 = depot。
    回傳 (matrix_km, duration_sec, source)。失敗拋例外。
    分塊呼叫：Google 單次上限 100 元素(origins×destinations)，
    故每塊 batch×batch 呼叫後拼回完整矩陣。
    """
    key = _load_key()
    n = len(coords)
    if n <= 1:
        return [[0.0]], [[0.0]], "google"
    coord_str = [f"{lat:.6f},{lon:.6f}" for lat, lon in coords]

    matrix_km = [[0.0] * n for _ in range(n)]
    duration_sec = [[0.0] * n for _ in range(n)]

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
                raise RuntimeError(f"Google Distance Matrix 失敗: {status} {data.get('error_message','')}")
            for i, row in enumerate(data["rows"]):
                gi = rs + i
                for j, el in enumerate(row["elements"]):
                    gj = cs + j
                    if el["status"] == "OK":
                        matrix_km[gi][gj] = el["distance"]["value"] / 1000.0
                        duration_sec[gi][gj] = el["duration"]["value"]
    return matrix_km, duration_sec, "google"
