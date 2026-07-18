"""
geo_cache.py — 持久化快取（讓 agent「學會」用過的地址/路線，減少 Google Cloud 費用）

核心概念：
  Google 查過的東西幾乎不變——一間店的座標永遠一樣、兩點間的道路距離也幾乎一樣。
  把查過的結果永久存在本機 JSON 檔，下次先查快取、命中就完全不呼叫 Google（$0）。
  只有「沒用過的新地址」才會真的打 Google，符合使用者「除非是沒用過的地址」的需求。

兩個快取檔（存在本專案目錄，可 git commit → push 到 Render 讓雲端也受惠）：
  - geo_cache.json     : 地址 -> [lat, lon]              （地理編碼結果）
  - matrix_cache.json  : "lat,lon|lat,lon" -> [km, sec]  （兩點道路距離+時間）

設計原則：
  - 純 stdlib，零第三方依賴。
  - 寫入用「先寫暫存檔再 rename」避免中途中斷損毀 JSON。
  - 座標一律四捨五入到小數 6 位當 key，避免浮點誤差造成快取 miss。
  - 任何讀寫失敗都安全吞掉（快取只是加速，壞了頂多退回呼叫 Google，不能讓主流程掛掉）。
"""

import json
import os
import threading

_HERE = os.path.dirname(os.path.abspath(__file__))
_GEO_PATH = os.path.join(_HERE, "geo_cache.json")
_MATRIX_PATH = os.path.join(_HERE, "matrix_cache.json")

_lock = threading.Lock()          # LINE 後台執行緒 + 主流程可能同時寫
_geo = None                       # 地址 -> (lat, lon)
_matrix = None                    # "a|b" -> (km, sec)

# 統計（本次執行期間），方便印出「省了幾次 Google 呼叫」
STATS = {"geo_hit": 0, "geo_miss": 0, "matrix_hit": 0, "matrix_miss": 0}


def _load(path):
    try:
        if os.path.exists(path):
            with open(path, encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def _ensure_loaded():
    global _geo, _matrix
    if _geo is None:
        _geo = _load(_GEO_PATH)
    if _matrix is None:
        _matrix = _load(_MATRIX_PATH)


def _save(path, data):
    """原子寫入：先寫 .tmp 再 rename，避免中途中斷把 JSON 寫壞。"""
    try:
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
        os.replace(tmp, path)
    except Exception:
        pass


# ---------------- 地理編碼快取 ----------------
def get_geo(address):
    """回傳 (lat, lon) 或 None（未快取）。"""
    _ensure_loaded()
    with _lock:
        v = _geo.get(address)
        if v:
            STATS["geo_hit"] += 1
            return float(v[0]), float(v[1])
        STATS["geo_miss"] += 1
        return None


def put_geo(address, latlon):
    """存入地址座標並立即寫檔。latlon=(lat,lon)。
    寫檔前先 reload 磁碟現有內容再 overlay，避免多程序/歷史條目被本程序 subset 覆寫而萎縮。"""
    if not address or not latlon:
        return
    _ensure_loaded()
    with _lock:
        # 重新從磁碟讀最新全量（其他程序可能寫了新條目），再疊加本次變更
        try:
            disk = json.load(open(_GEO_PATH, encoding="utf-8"))
        except Exception:
            disk = {}
        disk[address] = [float(latlon[0]), float(latlon[1])]
        _geo[address] = disk[address]
        _save(_GEO_PATH, disk)


# ---------------- 距離矩陣快取 ----------------
def _key(lat1, lon1, lat2, lon2):
    return f"{lat1:.6f},{lon1:.6f}|{lat2:.6f},{lon2:.6f}"


def get_pair(lat1, lon1, lat2, lon2):
    """回傳 (km, sec) 或 None（未快取）。"""
    _ensure_loaded()
    with _lock:
        v = _matrix.get(_key(lat1, lon1, lat2, lon2))
        if v is not None:
            STATS["matrix_hit"] += 1
            return float(v[0]), float(v[1])
        STATS["matrix_miss"] += 1
        return None


def put_pair(lat1, lon1, lat2, lon2, km, sec):
    _ensure_loaded()
    with _lock:
        _matrix[_key(lat1, lon1, lat2, lon2)] = [float(km), float(sec)]
    # 批次寫入時逐筆存檔太慢，改由 flush() 統一寫


def flush():
    """把距離矩陣快取一次寫回磁碟（批次呼叫後呼叫一次即可）。
    寫檔前先 reload 磁碟現有全量再 overlay 記憶體變更，避免本程序只 touched
    部分 key 就把全域快取（其他程序/歷史寫入的條目）砍掉而萎縮。"""
    _ensure_loaded()
    with _lock:
        try:
            disk = json.load(open(_MATRIX_PATH, encoding="utf-8"))
        except Exception:
            disk = {}
        disk.update(_matrix)   # 記憶體變更為權威，疊加到磁碟全量上
        _save(_MATRIX_PATH, disk)


def stats_line():
    g_tot = STATS["geo_hit"] + STATS["geo_miss"]
    m_tot = STATS["matrix_hit"] + STATS["matrix_miss"]
    return (f"🗃 快取命中：地理編碼 {STATS['geo_hit']}/{g_tot}、"
            f"距離 {STATS['matrix_hit']}/{m_tot}"
            f"（miss 才會呼叫 Google，命中=$0）")
