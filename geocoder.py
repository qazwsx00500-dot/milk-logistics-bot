"""
geocoder.py — 地址 -> 經緯度 (地理編碼)

使用 OpenStreetMap 的 Nominatim 公開服務 (免 API Key)。
限制與注意：
  - 免費服務速率限制 1 req/sec，本模組已內建 1 秒間隔。
  - 台灣門牌層級資料稀疏：完整門牌(如『忠孝東路四段1號』)常解析失敗，
    會自動降級到『路段級』(如『忠孝東路四段』)，誤差約一個街區。
  - 若需精準到門牌，請接 Google Geocoding / 台灣國土測繪中心 API (需金鑰)。
  - 內建記憶體快取：同一地址不重複查詢。
"""

import json
import re
import time
import urllib.request
import urllib.parse

NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
_HEADERS = {"User-Agent": "logistics-agent/0.1 (route planning prototype)"}

_cache = {}          # address -> (lat, lon) | None
_last_call = 0.0
_MIN_INTERVAL = 1.1  # 秒，遵守 Nominatim 速率限制


def _rate_limit():
    global _last_call
    wait = _MIN_INTERVAL - (time.time() - _last_call)
    if wait > 0:
        time.sleep(wait)
    _last_call = time.time()


def _query(q: str):
    url = (f"{NOMINATIM_URL}?format=jsonv2&countrycodes=TW&limit=1&q="
           + urllib.parse.quote(q))
    _rate_limit()
    req = urllib.request.Request(url, headers=_HEADERS)
    data = json.loads(urllib.request.urlopen(req, timeout=25).read().decode())
    if not data:
        return None
    return float(data[0]["lat"]), float(data[0]["lon"])


def _street_level(addr: str) -> str:
    """去掉門牌號碼，只保留到路段 (路/街/段)。"""
    # 去掉「數字號」及其後內容，例如『忠孝東路四段1號』->『忠孝東路四段』
    m = re.match(r"^(.*?(路|街|大道|巷|弄)[^0-9]*?[一二三四五六七八九十]?段?).*$", addr)
    if m:
        return m.group(1)
    # 去掉結尾的 數字號
    return re.sub(r"\d+號.*$", "", addr).strip()


def geocode(address: str):
    """
    輸入中文地址，回傳 (lat, lon) 或 None。
    優先用 Google Maps (門牌級精準)；失敗自動降級 OSM (路段級)。
    """
    if not address:
        return None
    if address in _cache:
        return _cache[address]

    # 0) 持久化快取（跨執行保存）：命中就完全不呼叫 Google（$0）。
    #    這是「除非是沒用過的新地址才呼叫 Google」的核心。
    try:
        import geo_cache
        cached = geo_cache.get_geo(address)
        if cached:
            _cache[address] = cached
            return cached
    except Exception:
        pass

    # 1) 優先 Google（門牌級）—— 只有快取沒有的新地址才會走到這
    try:
        from google_maps import geocode as g_geocode
        res = g_geocode(address)
        if res:
            _cache[address] = res
            try:
                import geo_cache
                geo_cache.put_geo(address, res)   # 存檔，下次不再花錢
            except Exception:
                pass
            return res
    except Exception:
        pass  # 降級 OSM

    # 2) OSM 完整地址
    try:
        res = _query(address)
        if res:
            _cache[address] = res
            return res
    except Exception:
        pass
    # 3) OSM 路段級降級
    try:
        street = _street_level(address)
        if street and street != address:
            res = _query(street)
            if res:
                _cache[address] = res
                return res
    except Exception:
        pass
    _cache[address] = None
    return None


def geocode_many(addresses: list) -> dict:
    """批次地理編碼，回傳 {address: (lat,lon)|None}。"""
    out = {}
    for a in addresses:
        out[a] = geocode(a)
    return out


def clear_cache():
    _cache.clear()
