"""
logistics_agent.py — 鮮奶物流「分組路線規劃」Agent

模式：照 Excel 的「車號」分組，每台車獨立排出最順拜訪順序。
     每台車有各自的出發點(出發點地址欄)。計算每間店下貨時間(瓶數*20秒)
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
    base = os.path.join(os.path.expanduser("~"), "OneDrive", "桌面", "當日車輛報表")
    if not os.path.isdir(base):
        base = os.path.join(os.path.expanduser("~"), "Desktop", "當日車輛報表")
    os.makedirs(base, exist_ok=True)
    return base

REPORT_DIR = _default_report_dir()


def make_sample(path):
    try:
        import openpyxl
    except ImportError:
        raise RuntimeError("請先安裝 openpyxl: uv pip install openpyxl")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "每日配送"
    ws.append(["車號", "店家名稱", "店家地址", "瓶數"])
    sample = [
        ("車01", "台中港區超市", "台中市梧棲區港埠路一段1號", 30),
        ("車01", "沙鹿鮮奶坊", "台中市沙鹿區中山路1號", 45),
        ("車01", "清水門市", "台中市清水區中山路1號", 20),
        ("車01", "大雅市場店", "台中市大雅區中清路一段1號", 35),
        ("車02", "豐原車站店", "台中市豐原區中正路1號", 55),
        ("車02", "潭子門市", "台中市潭子區潭子街1號", 28),
        ("車02", "神岡便利商店", "台中市神岡區神岡路1號", 24),
        ("車02", "后里批發", "台中市后里區甲后路一段1號", 30),
        ("車03", "西屯辦公室", "台中市西屯區台灣大道四段1號", 40),
        ("車03", "南屯量販", "台中市南屯區公益路二段1號", 25),
        ("車03", "北屯門市", "台中市北屯區文心路四段1號", 18),
        ("車03", "太平商店", "台中市太平區太平路1號", 18),
    ]
    for r in sample:
        ws.append(list(r))
    ws2 = wb.create_sheet("填寫說明")
    notes = [
        ["欄位", "是否必填", "說明"],
        ["車號", "必填", "同一台車的點填相同車號，如 車01 / 車A。程式照車號分組，每台車獨立排順序"],
        ["店家名稱", "必填", "店家名稱，會顯示在報表與地圖"],
        ["店家地址", "必填", "完整地址(到門牌)，用 Google 精準定位。例: 台中市大雅區101-1號"],
        ["瓶數", "必填", "鮮奶瓶數。下貨時間=瓶數×20秒，並計入到店預估時間"],
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
        return matrix_km, duration_matrix, source

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
    """從 .env 讀油資參數：優先用 FUEL_COST_PER_KM，否則用 油耗×油價 推算。"""
    env = {}
    p = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if os.path.exists(p):
        for line in open(p, encoding="utf-8"):
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                env[k.strip()] = v.strip().strip('"').strip("'")
    if env.get("FUEL_COST_PER_KM"):
        try:
            return float(env["FUEL_COST_PER_KM"])
        except ValueError:
            pass
    eff = env.get("FUEL_EFFICIENCY")
    price = env.get("FUEL_PRICE")
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
               "osrm": "OSRM", "fallback": "直線估算(降級)"}
    print(f"   距離來源：{src_map.get(source, source)}")

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


def main():
    ap = argparse.ArgumentParser(description="鮮奶物流分組路線規劃 Agent")
    ap.add_argument("--data", help="每日配送資料 (xlsx/csv)")
    ap.add_argument("--make-sample", action="store_true", help="產生範例 Excel 模板")
    ap.add_argument("--straight", action="store_true", help="強制直線距離/估速")
    ap.add_argument("--no-google", action="store_true", help="暫不走 Google，改 OSRM")
    ap.add_argument("--start", type=float, default=DEFAULT_START_HOUR, help="出發時間(24h, 預設9.5=09:30)")
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

    # 依執行日期分資料夾：REPORT_DIR/YYYY-MM-DD
    from datetime import datetime
    day_dir = os.path.join(REPORT_DIR, datetime.now().strftime("%Y-%m-%d"))
    os.makedirs(day_dir, exist_ok=True)

    html = report_mod.build_html_grouped(result, os.path.join(day_dir, "route_report.html"),
                                         meta={"start_hour": args.start})
    csvp = report_mod.build_csv_grouped(result, os.path.join(day_dir, "route_report.csv"))
    gmap = build_map(result, day_dir, use_google)
    print(f"\n📁 報表輸出資料夾：{day_dir}")
    print(f"📄 路線報表(HTML)：{html}")
    print(f"📄 路線報表(CSV) ：{csvp}")
    print(f"🌐 互動地圖      ：{gmap}")


def build_map(result, here, use_google):
    import json
    from osrm_client import get_route_geometry
    routes_geo = []
    # 統一藍色路線（使用者要求）
    ROUTE_COLOR = "#1a73e8"
    colors = [ROUTE_COLOR] * 10
    for i, rt in enumerate(result.routes):
        v = rt["vehicle"]
        color = colors[i % len(colors)]
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
</div>
<div class="panel" id="panel"></div>
<script>
const DATA = /*__DATA__*/;
const map = L.map('map').setView([24.15, 120.67], 12);
L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {maxZoom:19, attribution:'&copy; OpenStreetMap'}).addTo(map);
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
panel.innerHTML = html;
const all = []; DATA.routes.forEach(r => r.points.forEach(p => all.push(p)));
if (all.length) {
  allBounds = L.latLngBounds(all);
  // 初始視圖：不強制全覽(會讓點擠在一起)，改設適中縮放並以起點為中心
  if (firstStop) map.setView([firstStop[0], firstStop[1]], 12);
}
</script></body></html>"""


if __name__ == "__main__":
    main()
