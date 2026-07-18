"""
line_bot.py — LINE 機器人 Webhook 伺服器 (line-bot-sdk v3 寫法)

功能：
  1. 收到 Excel 檔案 → 下載 → 存成 DATA_DIR/每日配送.xlsx → 自動規劃路線
  2. 收到文字指令：
       跑 / 排程 / 安排 → 用 DATA_DIR/每日配送.xlsx 排今日路線
       倉庫 <地址>      → 設定總倉出發點
       狀態             → 顯示總倉與資料路徑
       幫助 / help      → 指令清單

報表檢視（經 cloudflared/ngrok 隧道對外開放）：
   https://<PUBLIC_URL>/report       (HTML 報表)
   https://<PUBLIC_URL>/report.csv   (CSV)
   /report 直接讀 REPORT_DIR/今天/ 的檔案，不依賴進程記憶體。

依賴：flask, line-bot-sdk (v3), openpyxl
設定：.env 裡 LINE_CHANNEL_SECRET / LINE_CHANNEL_ACCESS_TOKEN / PUBLIC_URL(可選)
隧道：cloudflared tunnel --url http://localhost:5000，把 <網址>/callback 填進 LINE Console。

啟動： python line_bot.py
"""

import os
import re
import io
import sys
import json
import time
import subprocess as _sp
import traceback
from datetime import datetime, timedelta, timezone


def _today_tw():
    """台灣時間(UTC+8)的今日日期字串 YYYY-MM-DD。
    報表/派車單/地圖的日期資料夾統一用台灣時間，避免 Render(UTC) 與使用者(台灣)跨午夜時
    產出的資料夾名稱與 sync_from_render 抓檔、使用者認知對不上（地圖連結失效）。"""
    return (datetime.now(timezone.utc) + timedelta(hours=8)).strftime("%Y-%m-%d")

from flask import Flask, request, abort, send_file, Response

# ---- line-bot-sdk v3 ----
from linebot.v3.webhook import WebhookParser
from linebot.v3.webhooks.models import MessageEvent, TextMessageContent, FileMessageContent
from linebot.v3.messaging import (
    Configuration,
    ApiClient,
    MessagingApi,
    MessagingApiBlob,
    ReplyMessageRequest,
    TextMessage,
)
from linebot.v3.exceptions import InvalidSignatureError

# ---- 載入本機 Agent ----
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

app = Flask(__name__)

# ---- 讀取設定 ----
def _load_env():
    env = {}
    p = os.path.join(HERE, ".env")
    if os.path.exists(p):
        for line in open(p, encoding="utf-8"):
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                env[k.strip()] = v.strip().strip('"').strip("'")
    return env

_ENV = _load_env()
CHANNEL_SECRET = os.environ.get("LINE_CHANNEL_SECRET") or _ENV.get("LINE_CHANNEL_SECRET", "")
CHANNEL_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN") or _ENV.get("LINE_CHANNEL_ACCESS_TOKEN", "")
PUBLIC_URL = os.environ.get("PUBLIC_URL") or _ENV.get("PUBLIC_URL", "").rstrip("/")

if not CHANNEL_SECRET or not CHANNEL_TOKEN:
    print("⚠ 警告：LINE_CHANNEL_SECRET / LINE_CHANNEL_ACCESS_TOKEN 未設定，"
          "請在 .env 加入後重啟。")

configuration = Configuration(access_token=CHANNEL_TOKEN)
parser = WebhookParser(CHANNEL_SECRET)
messaging_api = MessagingApi(ApiClient(configuration))
blob_api = MessagingApiBlob(ApiClient(configuration))


# ---- 健康檢查 ----
@app.route("/", methods=["GET"])
def index():
    return "OK - 鮮奶物流 LINE Bot 運作中", 200


def _git_head():
    """回傳當前 git HEAD 短 hash（部署在 Render 時能確認線上跑哪版）。"""
    try:
        return _sp.check_output(
            ["git", "rev-parse", "--short", "HEAD"],
            cwd=os.path.dirname(os.path.abspath(__file__)),
            stderr=_sp.DEVNULL).decode().strip()
    except Exception:
        return "unknown"

@app.route("/version", methods=["GET"])
def view_version():
    # 版本號與 deployment 標記；改動後 push 即更新，方便確認線上是否為最新
    import logistics_agent as L
    return Response(
        f"version=2026-07-19b | git={_git_head()} | "
        f"target_return={_hhmm(L.TARGET_RETURN_HOUR)} | geocode=parallel",
        mimetype="text/plain; charset=utf-8")


# ---- 報表檢視路由（直接讀今天日期資料夾的檔案） ----
def _today_report(which):
    """which: 'html' | 'csv' → 回傳今天報表檔路徑或 None。"""
    import logistics_agent as L
    day_dir = os.path.join(L.REPORT_DIR, _today_tw())
    name = "route_report.html" if which == "html" else "route_report.csv"
    p = os.path.join(day_dir, name)
    return p if os.path.exists(p) else None

@app.route("/report", methods=["GET"])
def view_report():
    p = _today_report("html")
    if p:
        return send_file(p)
    return Response("尚未產生報表。請在 LINE 傳『每日配送.xlsx』或『跑』觸發規劃。",
                    mimetype="text/plain; charset=utf-8")

@app.route("/report.csv", methods=["GET"])
def view_report_csv():
    p = _today_report("csv")
    if p:
        return send_file(p, mimetype="text/csv")
    return Response("尚未產生報表。", mimetype="text/plain; charset=utf-8")


def _fix_vid(vid):
    """Render(gunicorn) 預設 ascii 環境下，中文 URL 段會以 latin-1 解碼，
    導致 '車' 等中文字炸 UnicodeEncodeError。統一還原成 utf-8。"""
    try:
        return vid.encode("latin-1").decode("utf-8")
    except Exception:
        return vid


def _resolve_vid_file(vid, day_dir, prefix):
    """把 URL 裡的 vid 解析成實際檔案路徑。

    - 短網址 vN（N=1,2,3...）：從 day_dir 的 {prefix}_*.html 依檔名排序取第 N-1 個
      （穩定對應：產報表順序 = 車輛順序），避免中文車號 URL 編碼過長。
    - 舊網址（原車號 / _safe_veh 後車號）：向後相容，直接試。
    回傳檔案路徑或 None。
    """
    import re as _re
    # 短網址 vN
    m = _re.fullmatch(r"v(\d+)", vid)
    if m:
        n = int(m.group(1))
        files = sorted(
            f for f in os.listdir(day_dir)
            if f.startswith(prefix + "_") and f.endswith(".html")
        )
        if 1 <= n <= len(files):
            return os.path.join(day_dir, files[n - 1])
        return None
    # 舊網址相容
    import report as report_mod
    safe = report_mod._safe_veh(vid)
    for name in (prefix + "_" + safe + ".html", prefix + "_" + vid + ".html"):
        fp = os.path.join(day_dir, name)
        if os.path.exists(fp):
            return fp
    return None


def _persist_cache_to_git():
    """跑完規劃後，把累積的快取(geo_cache.json/matrix_cache.json)自動 push 回 GitHub。
    根因：Render 免費容器每次部署會清空硬碟 → 雲端快取每次都從零開始 →
    LINE 第一次跑仍是全量打 Google（幾十個新地址 = 幾十秒慢）。
    解法：每次跑完若有新快取，commit+push 回 GitHub，下次部署自動帶上累積快取，
    云端的 LINE 路徑也能 0-Google 跑完。
    - 本機(SSH remote)照常運作，不干擾。
    - Render 上用 GITHUB_TOKEN 環境變數以 HTTPS(x-access-token) 推送（容器無 SSH key）。
    - 任何失敗都吞掉：快取回寫只是加速，絕不能讓主流程/LINE 回傳掛掉。
    """
    try:
        import subprocess
        here = os.path.dirname(os.path.abspath(__file__))
        # 防護：容器內若不是 git repo 或無 .git（例如 Render 用非 git 部署），直接跳過
        if not os.path.isdir(os.path.join(here, ".git")):
            return
        # 只追蹤這兩個快取檔
        caches = ["geo_cache.json", "matrix_cache.json"]
        # 有變動才 push
        st = subprocess.run(["git", "status", "--porcelain"] + caches,
                            cwd=here, capture_output=True, text=True, timeout=20)
        if not st.stdout.strip():
            return  # 無新快取，跳過
        # 設定 push URL：有 GITHUB_TOKEN 就用 token HTTPS（Render 用），否則沿用現有 remote
        token = os.environ.get("GITHUB_TOKEN") or os.environ.get("GH_TOKEN")
        env = dict(os.environ)
        if token:
            subprocess.run(["git", "config", "user.email", "bot@milk-logistics.local"],
                           cwd=here, capture_output=True, timeout=20)
            subprocess.run(["git", "config", "user.name", "cache-bot"],
                           cwd=here, capture_output=True, timeout=20)
            # 臨時把 origin push URL 換成 token HTTPS（不改本機 SSH 設定）
            subprocess.run(
                ["git", "remote", "set-url", "origin",
                 f"https://x-access-token:{token}@github.com/qazwsx00500-dot/milk-logistics-bot.git"],
                cwd=here, capture_output=True, timeout=20, env=env)
        subprocess.run(["git", "add"] + caches, cwd=here, capture_output=True, timeout=30)
        subprocess.run(["git", "commit", "-m",
                        "chore: 自動回寫快取(geo/matrix) 加速雲端 LINE 路徑"],
                       cwd=here, capture_output=True, timeout=30)
        subprocess.run(["git", "push", "origin", "main"], cwd=here,
                       capture_output=True, timeout=90, env=env)
        # 若用 token 改過 remote，還原回 SSH（本機開發不受影響）
        if token:
            subprocess.run(["git", "remote", "set-url", "origin",
                            "git@github.com:qazwsx00500-dot/milk-logistics-bot.git"],
                           cwd=here, capture_output=True, timeout=20, env=env)
        print("✅ 快取已自動回寫 GitHub（下次部署雲端即 0-Google）")
    except Exception as e:
        print(f"⚠ 快取回寫失敗（不影響 LINE 回傳）: {e}")



@app.route("/report/<vid>", methods=["GET"])
def view_report_vehicle(vid):
    """單一車輛的獨立報表（含 PNG 下載按鈕，可分別轉給司機）。支援短網址 vN 與舊車號網址。"""
    import logistics_agent as L
    vid = _fix_vid(vid)
    day_dir = os.path.join(L.REPORT_DIR, _today_tw())
    fp = _resolve_vid_file(vid, day_dir, "route_report")
    if fp:
        return send_file(fp)
    return Response("找不到該車報表。請先在 LINE 傳 Excel 或『跑』觸發規劃。",
                    mimetype="text/plain; charset=utf-8")


# ---- 派車單檢視路由（每台車一份，給司機/內勤） ----
def _today_dispatch(which):
    """which: 'html' | 'csv' → 回傳今天派車單檔路徑或 None。"""
    import logistics_agent as L
    day_dir = os.path.join(L.DISPATCH_DIR, _today_tw())
    name = "dispatch.html" if which == "html" else "dispatch.csv"
    p = os.path.join(day_dir, name)
    return p if os.path.exists(p) else None

@app.route("/dispatch", methods=["GET"])
def view_dispatch():
    p = _today_dispatch("html")
    if p:
        return send_file(p)
    return Response("尚未產生派車單。請在 LINE 傳『每日配送.xlsx』或『跑』觸發規劃。",
                    mimetype="text/plain; charset=utf-8")

@app.route("/dispatch.csv", methods=["GET"])
def view_dispatch_csv():
    p = _today_dispatch("csv")
    if p:
        return send_file(p, mimetype="text/csv")
    return Response("尚未產生派車單。", mimetype="text/plain; charset=utf-8")


# ---- 整合 Excel 檢視路由（3 分頁：路線總表/各車派車單/總計） ----
def _today_workbook():
    import logistics_agent as L
    day_dir = os.path.join(L.DISPATCH_DIR, _today_tw())
    p = os.path.join(day_dir, "整合報表.xlsx")
    return p if os.path.exists(p) else None

@app.route("/workbook", methods=["GET"])
def view_workbook():
    p = _today_workbook()
    if p:
        return send_file(p,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                         as_attachment=True, download_name="整合報表.xlsx")
    return Response("尚未產生整合報表。請在 LINE 傳『每日配送.xlsx』或『跑』觸發規劃。",
                    mimetype="text/plain; charset=utf-8")


# ---- 路線圖 PNG 檢視路由（供本機同步器抓取） ----
def _today_map_png():
    import logistics_agent as L
    day_dir = os.path.join(L.DISPATCH_DIR, _today_tw())
    p = os.path.join(day_dir, "route_map.png")
    return p if os.path.exists(p) else None

@app.route("/map_png", methods=["GET"])
def view_map_png():
    p = _today_map_png()
    if p:
        return send_file(p, mimetype="image/png",
                         as_attachment=True, download_name="route_map.png")
    return Response("尚未產生路線圖。請先傳 Excel 觸發規劃。",
                    mimetype="text/plain; charset=utf-8")


# ---- 指令處理 ----
def handle_text(text: str) -> str:
    """回傳要推回 LINE 的文字訊息。"""
    t = text.strip()
    cmd = t.lower()

    if cmd in ("幫助", "help", "?", "指令"):
        return (
            "🤖 鮮奶物流路線機器人指令：\n"
            "・直接傳『每日配送.xlsx』給我 → 我下載並自動排版\n"
            "・跑 / 排程 / 安排 → 用已存的 Excel 排今日路線\n"
            "・倉庫 <地址> → 設定總倉出發點 (例: 倉庫 台中市大雅區101-1號)\n"
            "・狀態 → 顯示目前總倉與資料路徑\n"
            "・幫助 → 顯示本清單"
        )

    if cmd in ("狀態", "status"):
        import logistics_agent as L
        return (f"📍 總倉：{L.DEPOT.address}\n"
                f"📂 資料夾：{L.DATA_DIR}\n"
                f"🔑 Google：{'已設定' if L.DEPOT.lat else '未定位'}\n"
                f"🌐 公網網址：{PUBLIC_URL or '（未設定 PUBLIC_URL）'}")

    if cmd.startswith("倉庫"):
        addr = re.sub(r"^(倉庫)\s*", "", t).strip()
        if not addr:
            return "請提供地址，例如：倉庫 台中市大雅區101-1號"
        import logistics_agent as L
        from geocoder import geocode
        try:
            coord = geocode(addr)
        except Exception:
            coord = None
        if not coord:
            return f"⚠ 無法定位此地址：{addr}\n請換個寫法或檢查網路/Google Key。"
        L.DEPOT.address = addr
        L.DEPOT.lat, L.DEPOT.lon = coord
        _save_depot(addr)
        return f"✅ 總倉已設為：{addr}\n({coord[0]:.4f}, {coord[1]:.4f})\n下次排程即從此出發。"

    if cmd in ("跑", "排程", "安排", "run", "go"):
        return run_plan()

    return ("❓ 我不懂這個指令。\n輸入「幫助」看可用指令；或直接傳 Excel 檔給我。")


def _save_depot(addr: str):
    """把總倉地址寫進 .env（保留其它設定）。"""
    p = os.path.join(HERE, ".env")
    lines = []
    if os.path.exists(p):
        lines = open(p, encoding="utf-8").read().splitlines()
    kept = [l for l in lines if not l.strip().startswith("DEPOT_ADDR=")]
    kept.append(f'DEPOT_ADDR={addr}')
    with open(p, "w", encoding="utf-8") as f:
        f.write("\n".join(kept) + "\n")
    import logistics_agent as L
    L.DEPOT_ADDR = addr


def run_plan(data_path=None, rows=None):
    """執行排程，回傳文字摘要。
    data_path: 標準 xlsx/csv 路徑（一般『跑』指令用）
    rows: normalize_excel 出來的 [(車號,店名,地址,瓶數)...]（LINE 傳檔用）
           若 rows 裡車號皆空 → Agent 自動決定車數（先1台，超時則按區域拆）
    """
    import logistics_agent as L
    import report as report_mod
    import excel_normalizer as en
    try:
        use_google = True
        no_google = False
        start_hour = L.DEFAULT_START_HOUR   # 9.5 = 09:30 出車
        fuel_cost = L._load_fuel_cost()
        # data_path 檔若含『油資單價』欄 → 優先
        if data_path:
            try:
                from data_loader import read_fuel_cost as _rfc
                _ef = _rfc(data_path)
                if _ef:
                    fuel_cost = _ef
            except Exception:
                pass

        # 把 rows 寫成臨時標準檔（供 data_loader 讀）
        if rows is not None and data_path is None:
            import excel_normalizer as en
            # 記錄「原始是否無車號」(後面分車判斷要用, 因為下面會先填車01)
            had_no_vehicle = all(r[0] == "" for r in rows)
            # 無車號 → 先填 車01 (auto_assign_vehicles 保守起點, 供 solve_grouped 跑第一次)
            if had_no_vehicle:
                rows = en.auto_assign_vehicles(rows)
            import openpyxl as _ox
            from openpyxl import Workbook as _WB
            tmp = os.path.join(HERE, "_normalized_tmp.xlsx")
            wb = _WB(); ws = wb.active; ws.title = "每日配送"
            ws.append(["車號", "店家名稱", "店家地址", "瓶數", "品項"])
            for veh, n, a, q, *rest in rows:
                item = rest[0] if rest else ""
                ws.append([veh, n, a, q, item])
            wb.save(tmp)
            data_path = tmp

        # 無車號 → 先填 車01 (auto_assign_vehicles 保守起點, 供 solve_grouped 跑第一次)
        # 註：改用 L.plan_auto_assign 單輪處理（不再雙重 L.plan），故無車號時直接走自動分車
        if not data_path or not os.path.exists(data_path):
            return (f"⚠ 找不到資料檔。\n請直接把 Excel 傳給我，或先放到『路線規劃』資料夾。")

        # 路線規劃：有車號走一般 plan；無車號走單輪自動分車(全站只打一次矩陣)
        if rows is not None and had_no_vehicle:
            result, skipped = L.plan_auto_assign(
                start_hour, data_path, use_google, no_google,
                fuel_cost_per_km=fuel_cost)
        else:
            result, skipped = L.plan(
                start_hour, data_path, use_google, no_google,
                fuel_cost_per_km=fuel_cost)
        if result is None:
            return "⚠ 排程失敗：沒有可規劃的車輛/店家。請檢查 Excel 欄位。"

        # 產報表到 日期子資料夾
        day_dir = os.path.join(L.REPORT_DIR, _today_tw())
        os.makedirs(day_dir, exist_ok=True)
        report_mod.build_html_grouped(result, os.path.join(day_dir, "route_report.html"),
                                      meta={"start_hour": start_hour})
        report_mod.build_csv_grouped(result, os.path.join(day_dir, "route_report.csv"))
        per_veh = report_mod.build_html_per_vehicle(result, day_dir, meta={"start_hour": start_hour})
        gmaps = L.build_map(result, day_dir, use_google)  # [總圖, 各車圖...] 雲端產互動HTML(無瀏覽器不截PNG)
        # 路線圖 PNG 由本機直跑時產生（本機有 Edge/Chrome）；雲端(Render)無瀏覽器，
        # 不嘗試截圖。使用者於本機執行 sync_from_render.py 會抓本機已產出的 PNG。
        map_png = None

        # 派車單（每台車一份）→ 獨立 DISPATCH_DIR/日期/
        dispatch_dir = os.path.join(L.DISPATCH_DIR, _today_tw())
        os.makedirs(dispatch_dir, exist_ok=True)
        report_mod.build_dispatch_grouped(result, dispatch_dir, meta={"start_hour": start_hour})
        # 結構化資料 JSON（供客服助理撈 ETA/貨品/載貨量）
        report_mod.build_dispatch_data(result, os.path.join(dispatch_dir, "dispatch_data.json"),
                                       meta={"start_hour": start_hour})
        # 整合 Excel 雲端不自動產（缺路線圖分頁），改由本機 --excel 直跑產出
        xlsx = None

        # 文字摘要
        veh_note = ""
        if rows is not None and all(r[0] == "" for r in rows):
            veh_note = f"\n🚚 由 Agent 依『09:30出車/{_hhmm(L.TARGET_RETURN_HOUR)}回倉』自動安排 {len(result.routes)} 台車（最多3台，一台跑不完才加車）"
        lines = [f"📦 路線規劃完成（{result.distance_source}）",
                 f"出發 09:30 ｜ 目標回倉 {_hhmm(L.TARGET_RETURN_HOUR)}{veh_note}",
                 f"車數 {len(result.routes)} 台 ｜ 總實際里程 {result.total_distance_km:.0f} km ｜ 總瓶數 {int(result.total_load)}"]
        if result.fuel_cost_per_km > 0:
            lines.append(f"⛽ 油資單價 {result.fuel_cost_per_km:.1f} 元/km ｜ 預估總油資 {result.total_fuel_cost:.0f} 元")
        for rt in result.routes:
            v = rt["vehicle"]
            ret = _hhmm(rt["end_hour"])
            tag = "✅準時回倉" if rt.get("on_time") else f"⚠超過{_hhmm(L.TARGET_RETURN_HOUR)}({ret})"
            fuel_txt = f" 油資{rt.get('fuel_cost',0):.0f}元" if result.fuel_cost_per_km > 0 else ""
            lines.append(f"\n【{v.id}】{len(rt['stops'])}站 {rt['distance_km']:.0f}km {ret}回 {tag}{fuel_txt}")
            for si, s in enumerate(rt["stops"][:5]):
                a, lv = rt["etas"][si]
                lines.append(f"  {si+1}. {s.name} 到{_hhmm(a)} 離{_hhmm(lv)}")
            if len(rt["stops"]) > 5:
                lines.append(f"  …其餘 {len(rt['stops'])-5} 站請看報表")
        if skipped:
            lines.append(f"\n⚠ 跳過 {len(skipped)} 筆：")
            lines.append("  " + "; ".join(f"{n}({r})" for n, r in skipped[:10]))

        # 附公網報表連結
        if PUBLIC_URL:
            lines.append(f"\n📄 總表：{PUBLIC_URL}/report")
            lines.append(f"📊 CSV ：{PUBLIC_URL}/report.csv")
            lines.append(f"🗺️ 地圖（總圖）：{PUBLIC_URL}/route_map")
            lines.append("   ※ 地圖頁內含「📥 下載地圖 PNG」按鈕，手機/電腦可直接存圖")
            lines.append("🚚 各車獨立報表（可分別轉給司機，可下載PNG）：")
            for i, rt in enumerate(result.routes, 1):
                vid = rt["vehicle"].id
                lines.append(f"  ・{vid} 報表：{PUBLIC_URL}/report/v{i}")
                lines.append(f"  ・{vid} 路線圖：{PUBLIC_URL}/route_map/v{i}")
        else:
            lines.append(f"\n📁 報表已產出：{day_dir}")

        # ☁ 若設了 OneDrive 憑證，把報表同步上傳到 OneDrive，
        #    讓本機「OneDrive 同步桌面」自動出現報表（雲端跑也看得到）。
        try:
            import onedrive_sync as ods
            if ods.upload_report_dir(day_dir, _today_tw()):
                lines.append(f"☁ 報表已同步至 OneDrive（本機桌面會自動出現）。")
        except Exception as e:
            print(f"⚠ OneDrive 同步失敗（不影響 LINE 回傳）: {e}")

        # 🔍 產出自動複檢：把 self_check 的終端輸出轉成 LINE 文字附上，
        #    讓手機端也能看到 ①~⑤ 檢查結果與 ⛔ 警告（本機 CLI 已有，這裡補齊 LINE 路徑）。
        try:
            import io as _io, contextlib as _cl
            _buf = _io.StringIO()
            with _cl.redirect_stdout(_buf):
                _ok = L.self_check(result, None)
            _chk = _buf.getvalue().strip()
            if _chk:
                lines.append(f"\n🔍 自動複檢：\n{_chk}")
        except Exception as e:
            print(f"⚠ 自動複檢失敗（不影響 LINE 回傳）: {e}")

        # 💾 跑完自動把累積快取(geo/matrix)回寫 GitHub（背景執行，不阻塞回傳）
        #    關鍵：git push 在 Render 雲端可能因 SSH/HTTPS 互動卡住，若同步呼叫會
        #    讓整個後台處理緒卡在 push 那行 → LAST_RESULT 寫不進、Push 發不出 → 使用者收不到結果。
        #    改丟進 daemon 背景線程，主流程先回傳，push 在背後跑（卡住只卡那個線程）。
        try:
            threading.Thread(target=_persist_cache_to_git, daemon=True).start()
        except Exception as e:
            print(f"⚠ 快取回寫啟動失敗（不影響 LINE 回傳）: {e}")

        return "\n".join(lines)
    except Exception as e:
        traceback.print_exc()
        return f"⚠ 排程時發生錯誤：{type(e).__name__}: {str(e)[:120]}"


def handle_file(event, file_msg) -> str:
    """收到 Excel/CSV 檔 → 下載 → 轉標準每日配送表 → 存檔(日期名) → 自動規劃。"""
    import logistics_agent as L
    import excel_normalizer as en
    fname = getattr(file_msg, "file_name", "") or ""
    ext = os.path.splitext(fname)[1].lower()
    if ext not in (".xlsx", ".xlsm", ".xls", ".csv"):
        return (f"⚠ 我只接受 Excel/CSV 檔（收到的是 {fname or '未知格式'}）。\n"
                f"請傳包含店家/地址/瓶數的 Excel，我會自動轉成每日配送表。")

    try:
        # 1) 下載檔案內容（blob API → bytearray）
        data = blob_api.get_message_content(file_msg.id)
        # 先存原始檔到暫存
        import tempfile
        tmp_in = os.path.join(tempfile.gettempdir(), "line_upload" + ext)
        with open(tmp_in, "wb") as f:
            f.write(bytes(data))
    except Exception as e:
        traceback.print_exc()
        return f"⚠ 下載檔案失敗：{type(e).__name__}: {str(e)[:100]}"

    # 2) 轉成標準每日配送表，存成 每日配送_YYYYMMDD.xlsx
    try:
        rows, skipped, out_path, excel_fuel = en.normalize_excel(tmp_in, L.DATA_DIR)
        if not rows:
            return ("⚠ 無法從這個 Excel 讀到店家資料。\n"
                    "需要的欄位：店家名稱 / 店家地址 / 瓶數（車號可省略，我會自動安排）。")
    except Exception as e:
        traceback.print_exc()
        return f"⚠ 轉換 Excel 失敗：{type(e).__name__}: {str(e)[:100]}"

    # 3) 不直接跑規劃，先問使用者要幾台車（Quick Reply 選單）
    # 轉檔後自行檢查 (避免把異常資料丟給 JOJO)
    chk_lines = []
    err_n = 0
    for veh, name, addr, qty, item in rows:
        if not addr or not str(addr).strip():
            chk_lines.append(f"  ❌ 「{name}」地址空白，無法定位")
            err_n += 1
        if isinstance(qty, (int, float)) and qty < 0:
            chk_lines.append(f"  ❌ 「{name}」瓶數為負({qty})")
            err_n += 1
    pending = {
        "rows": rows,
        "skipped_n": len(skipped),
        "fname": fname,
        "excel_fuel": excel_fuel,
        "ts": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    PENDING_FILES[user_id_key(event)] = pending
    _save_pending(PENDING_FILES)  # 落盤，避免 worker 重啟丟失選車狀態
    from linebot.v3.messaging import QuickReply, QuickReplyItem, MessageAction
    qr = QuickReply(items=[
        QuickReplyItem(action=MessageAction(label="🤖 自動安排", text="自動安排")),
        QuickReplyItem(action=MessageAction(label="1 台", text="1台")),
        QuickReplyItem(action=MessageAction(label="2 台", text="2台")),
        QuickReplyItem(action=MessageAction(label="3 台", text="3台")),
    ])
    # ★ 互動模式：傳檔後「先問車數」，使用者選完才照規則跑（不自動跑）。
    #   保留之前為修「收不到結果」加的健壯性：PENDING 已落盤(570-571)，
    #   使用者回「自動安排/1台/2台/3台」會經 _process_and_push →
    #   run_plan_choice 觸發規劃並 Push 結果；worker 重啟也不丟選車狀態。
    msg = (f"📥 已收到「{fname}」\n"
           f"✅ 已自動轉成標準每日配送表（{len(rows)} 筆店家，{len(skipped)} 筆跳過）\n"
           f"🔍 轉檔自檢：{'✅ 通過' if not chk_lines else str(len(chk_lines)) + ' 項需留意'}\n")
    if chk_lines:
        msg += "\n".join(chk_lines) + "\n"
    msg += ("\n🚚 請選擇車輛安排方式（點下方選單）：\n"
            "  • 自動安排 → 由 Agent 依『09:30出車/{_hhmm(L.TARGET_RETURN_HOUR)}回倉』自動決定車數(最多3台)\n"
            f"  • 指定台數(1/2/3台) → 只求最快回倉，不限制{_hhmm(L.TARGET_RETURN_HOUR)}")
    return (msg, qr)


def user_id_key(event):
    return getattr(getattr(event, "source", None), "user_id", None) or "default"


def run_plan_choice(user_id, choice_text, pending):
    """根據使用者選擇跑規劃。choice_text: '自動安排' / '1台' / '2台' / '3台'。"""
    import logistics_agent as L
    import report as report_mod
    import excel_normalizer as en
    fuel_cost = L._load_fuel_cost()
    if pending.get("excel_fuel"):   # Excel『油資單價』欄優先於 .env/環境變數
        fuel_cost = pending["excel_fuel"]
    start_hour = L.DEFAULT_START_HOUR
    rows = pending["rows"]

    # 寫成臨時標準檔
    import openpyxl as _ox
    from openpyxl import Workbook as _WB
    tmp = os.path.join(HERE, "_normalized_tmp.xlsx")
    wb = _WB(); ws = wb.active; ws.title = "每日配送"
    ws.append(["車號", "店家名稱", "店家地址", "瓶數", "品項"])
    for veh, n, a, q, *rest in rows:
        item = rest[0] if rest else ""
        ws.append([veh, n, a, q, item])
    wb.save(tmp)

    had_no_vehicle = all(r[0] == "" for r in rows)

    # 判斷車數模式
    force_v = None
    if choice_text in ("1台", "2台", "3台"):
        force_v = int(choice_text[0])
    # '自動安排' 或 其他 → 自動

    if force_v is not None:
        result, skipped = L.plan_auto_assign(
            start_hour, tmp, True, False, fuel_cost_per_km=fuel_cost,
            force_vehicles=force_v)
        mode_note = f"你指定 {force_v} 台（只求最快回倉，不限制{_hhmm(L.TARGET_RETURN_HOUR)}）"
    else:
        result, skipped = L.plan_auto_assign(
            start_hour, tmp, True, False, fuel_cost_per_km=fuel_cost)
        mode_note = f"由 Agent 依『09:30出車/{_hhmm(L.TARGET_RETURN_HOUR)}回倉』自動安排"

    if result is None:
        return "⚠ 排程失敗：沒有可規劃的車輛/店家。"
    return _format_result(result, skipped, fuel_cost, mode_note, PUBLIC_URL)


CHOICE_WORDS = {"自動安排", "1台", "2台", "3台"}


def _format_result(result, skipped, fuel_cost, mode_note, public_url):
    """把 PlanResult 組成 LINE 文字摘要（含報表連結）。"""
    import logistics_agent as L
    import report as report_mod
    # 產報表到 日期子資料夾
    from datetime import datetime
    day_dir = os.path.join(L.REPORT_DIR, _today_tw())
    os.makedirs(day_dir, exist_ok=True)
    report_mod.build_html_grouped(result, os.path.join(day_dir, "route_report.html"),
                                  meta={"start_hour": L.DEFAULT_START_HOUR})
    report_mod.build_csv_grouped(result, os.path.join(day_dir, "route_report.csv"))
    per_veh = report_mod.build_html_per_vehicle(result, day_dir, meta={"start_hour": L.DEFAULT_START_HOUR})
    L.build_map(result, day_dir, True)
    # 路線圖 PNG 由本機直跑時產生（本機有 Edge/Chrome）；雲端(Render)無瀏覽器，不嘗試截圖。
    map_png = None

    # 派車單（每台車一份）→ 獨立 DISPATCH_DIR/日期/
    dispatch_dir = os.path.join(L.DISPATCH_DIR, _today_tw())
    os.makedirs(dispatch_dir, exist_ok=True)
    if map_png:
        import shutil as _sh
        _sh.copy(map_png, os.path.join(dispatch_dir, "route_map.png"))
    report_mod.build_dispatch_grouped(result, dispatch_dir, meta={"start_hour": L.DEFAULT_START_HOUR})
    # 整合 Excel 雲端不自動產（缺路線圖分頁），改由本機 --excel 直跑產出
    xlsx = None

    lines = [f"📦 路線規劃完成（{result.distance_source}）",
             f"🚚 {mode_note}",
             f"車數 {len(result.routes)} 台 ｜ 總實際里程 {result.total_distance_km:.0f} km ｜ 總瓶數 {int(result.total_load)}"]
    if result.fuel_cost_per_km > 0:
        lines.append(f"⛽ 油資單價 {result.fuel_cost_per_km:.1f} 元/km ｜ 預估總油資 {result.total_fuel_cost:.0f} 元")
    for rt in result.routes:
        v = rt["vehicle"]
        ret = _hhmm(rt["end_hour"])
        if rt.get("on_time"):
            tag = "✅準時回倉"
        else:
            tag = f"⚠回倉 {ret}（指定台數模式不限制{_hhmm(L.TARGET_RETURN_HOUR)}）" if "指定" in mode_note else f"⚠超過{_hhmm(L.TARGET_RETURN_HOUR)}({ret})"
        fuel_txt = f" 油資{rt.get('fuel_cost',0):.0f}元" if result.fuel_cost_per_km > 0 else ""
        lines.append(f"\n【{v.id}】{len(rt['stops'])}站 {rt['distance_km']:.0f}km {ret}回 {tag}{fuel_txt}")
        for si, s in enumerate(rt["stops"][:5]):
            a, lv = rt["etas"][si]
            lines.append(f"  {si+1}. {s.name} 到{_hhmm(a)} 離{_hhmm(lv)}")
        if len(rt["stops"]) > 5:
            lines.append(f"  …其餘 {len(rt['stops'])-5} 站請看報表")
    if skipped:
        lines.append(f"\n⚠ 跳過 {len(skipped)} 筆：")
        lines.append("  " + "; ".join(f"{n}({r})" for n, r in skipped[:10]))

    if public_url:
        lines.append(f"\n📄 總表：{public_url}/report")
        lines.append(f"📊 CSV ：{public_url}/report.csv")
        lines.append(f"🗺️ 地圖（總圖）：{public_url}/route_map")
        lines.append("   ※ 地圖頁內含「📥 下載地圖 PNG」按鈕，手機/電腦可直接存圖")
        lines.append("🚚 各車獨立報表（可分別轉給司機，可下載PNG）：")
        for i, rt in enumerate(result.routes, 1):
            vid = rt["vehicle"].id
            # 短網址：用穩定序號 v1/v2/v3（避免中文車號 URL 編碼過長）
            lines.append(f"  ・{vid} 報表：{public_url}/report/v{i}")
            lines.append(f"  ・{vid} 路線圖：{public_url}/route_map/v{i}")
    else:
        lines.append(f"\n📁 報表已產出：{day_dir}")

    # 💾 跑完自動把累積快取(geo/matrix)回寫 GitHub（背景執行，不阻塞回傳）
    #    同 run_plan：git push 在雲端可能卡住，改丟 daemon 背景線程，主流程先回傳。
    try:
        threading.Thread(target=_persist_cache_to_git, daemon=True).start()
    except Exception as e:
        print(f"⚠ 快取回寫啟動失敗（不影響 LINE 回傳）: {e}")

    return "\n".join(lines)


def _run_choice_from_latest(choice_text):
    """fallback：使用者直接發車數指令(自動安排/1台/2台/3台)但無 PENDING 暫存時，
    找最近一次的每日配送檔直接跑規劃。避免 Render worker 重啟導致 PENDING 遺失後選車失效。"""
    import glob as _glob
    import logistics_agent as L
    cands = sorted(_glob.glob(os.path.join(L.DATA_DIR, "每日配送_*.xlsx")),
                   key=os.path.getmtime, reverse=True)
    if not cands:
        return ("⚠ 找不到可規劃的每日配送檔。請先傳 Excel 給我，我再幫你排程。")
    latest = cands[0]
    force_v = None
    if choice_text in ("1台", "2台", "3台"):
        force_v = int(choice_text[0])
    try:
        result, skipped = L.plan_auto_assign(
            L.DEFAULT_START_HOUR, latest, True, False,
            fuel_cost_per_km=L._load_fuel_cost(), force_vehicles=force_v)
        mode_note = (f"你指定 {force_v} 台" if force_v else "由 Agent 自動安排")
        return _format_result(result, skipped, L._load_fuel_cost(), mode_note, PUBLIC_URL)
    except Exception as e:
        traceback.print_exc()
        return f"⚠ 用最近檔案({os.path.basename(latest)})規劃失敗：{type(e).__name__}: {str(e)[:120]}"


def _hhmm(h):
    hh = int(h) % 24
    mm = int(round((h - int(h)) * 60))
    if mm == 60:
        hh += 1; mm = 0
    return f"{hh:02d}:{mm:02d}"


# ---- Webhook ----
import threading

# 最近一次處理結果（Push 失敗時的備援，可從 /last_result 網頁查看）
LAST_RESULT = {"ts": None, "text": "(尚無結果)", "pushed": None, "error": None}

# 待使用者選車數的暫存（user_id → {rows, fname, ...}）
# ⚠️ 必須持久化到磁碟：Render 免費版 worker 閒置會回收，記憶體 dict 會丟失，
#    導致「傳檔→選車」兩步互動在 worker 重啟後 PENDING_FILES 變空、選車不觸發規劃。
_PENDING_PATH = os.path.join(HERE, "_pending_files.json")
def _load_pending():
    try:
        with open(_PENDING_PATH, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}
def _save_pending(d):
    try:
        with open(_PENDING_PATH, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False)
    except Exception:
        pass
PENDING_FILES = _load_pending()  # 啟動時從磁碟恢復

def _push_to(user_id, text, quick_reply=None):
    """用 Push 推結果給使用者（繞過 reply_token 1 秒過期限制）。"""
    global LAST_RESULT
    try:
        from linebot.v3.messaging.models import PushMessageRequest
        if not user_id:
            raise ValueError("user_id 為空，無法 Push（可能事件結構不含 source.user_id）")
        msgs = [TextMessage(text=text[:1900])]
        if quick_reply is not None:
            msgs[0].quick_reply = quick_reply
        messaging_api.push_message(
            PushMessageRequest(
                to=user_id,
                messages=msgs,
            )
        )
        LAST_RESULT["pushed"] = True
    except Exception as e:
        LAST_RESULT["pushed"] = False
        LAST_RESULT["error"] = f"{type(e).__name__}: {str(e)[:200]}"
        traceback.print_exc()

def _process_and_push(user_id, kind, payload):
    """背景執行緒：實際處理（可能幾十秒），完成後用 Push 推結果。"""
    global LAST_RESULT
    import datetime as _dt
    try:
        if kind == "text":
            text = payload
            # 車數選擇互動：傳檔後使用者回「自動安排/1台/2台/3台」且有暫存檔
            if text.strip() in CHOICE_WORDS and user_id in PENDING_FILES:
                pending = PENDING_FILES.pop(user_id, None)
                _save_pending(PENDING_FILES)  # 取走後更新磁碟
                result = run_plan_choice(user_id, text.strip(), pending)
            elif text.strip() in CHOICE_WORDS:
                # fallback：無暫存（worker 重啟/PENDING 遺失）→ 讀最近的每日配送檔直接跑
                result = _run_choice_from_latest(text.strip())
            else:
                result = handle_text(text)
        elif kind == "file":
            result = handle_file(payload["event"], payload["file_msg"])
        else:
            result = "⚠ 不支援的訊息類型。"
        # 支援 (text, quick_reply) 回傳
        qr = None
        if isinstance(result, tuple):
            result, qr = result
    except Exception as e:
        traceback.print_exc()
        result = f"⚠ 處理時發生錯誤：{type(e).__name__}: {str(e)[:120]}"
        qr = None
    LAST_RESULT["ts"] = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    LAST_RESULT["text"] = result
    _push_to(user_id, result, quick_reply=qr)

@app.route("/callback", methods=["POST", "GET"])
def callback():
    if request.method == "GET":
        return "OK", 200
    signature = request.headers.get("X-Line-Signature", "")
    body = request.get_data(as_text=True)
    try:
        events = parser.parse(body, signature)
    except InvalidSignatureError:
        abort(400)
        return
    except Exception:
        traceback.print_exc()
        abort(400)
        return

    from linebot.v3.webhooks.models import PostbackEvent
    for event in events:
        # 選單回傳可能是 MessageEvent(文字) 或 PostbackEvent，兩者都處理
        if isinstance(event, PostbackEvent):
            # 把 postback 的 data 當作文字指令走同一分發邏輯
            _pb_text = getattr(getattr(event, "postback", None), "data", None) or ""
            user_id = getattr(getattr(event, "source", None), "user_id", None)
            try:
                messaging_api.reply_message(ReplyMessageRequest(
                    reply_token=event.reply_token,
                    messages=[TextMessage(text="✅ 收到，處理中…（完成後我會主動推播結果）")]))
            except Exception:
                traceback.print_exc()
            t = threading.Thread(target=_process_and_push,
                                 args=(user_id, "text", _pb_text), daemon=True)
            t.start()
            continue
        if not isinstance(event, MessageEvent):
            continue
        # 先拿到 user_id（Push 用）
        user_id = getattr(getattr(event, "source", None), "user_id", None)
        # 立即 reply 一個確認（reply_token 1 秒內必須用掉）
        try:
            extra = f"\n📋 結果也會貼在：{PUBLIC_URL}/last_result" if PUBLIC_URL else ""
            messaging_api.reply_message(
                ReplyMessageRequest(
                    reply_token=event.reply_token,
                    messages=[TextMessage(text="✅ 收到，處理中…（完成後我會主動推播結果）" + extra)],
                )
            )
        except Exception:
            traceback.print_exc()
        # 背景執行緒跑實際處理，完成後 Push 結果
        if isinstance(event.message, TextMessageContent):
            t = threading.Thread(target=_process_and_push,
                                 args=(user_id, "text", event.message.text), daemon=True)
            t.start()
        elif isinstance(event.message, FileMessageContent):
            t = threading.Thread(target=_process_and_push,
                                 args=(user_id, "file",
                                       {"event": event, "file_msg": event.message}), daemon=True)
            t.start()
    return "OK"


@app.route("/route_map", methods=["GET"])
def view_route_map():
    """當日路線地圖 (route_map.html)。"""
    import logistics_agent as L
    day = _today_tw()
    p = os.path.join(L.REPORT_DIR, day, "route_map.html")
    if os.path.exists(p):
        return send_file(p)
    return Response("尚無路線地圖。請先傳 Excel 觸發規劃。", mimetype="text/plain; charset=utf-8")


@app.route("/route_map/<vid>", methods=["GET"])
def view_route_map_vehicle(vid):
    """各車獨立路線地圖 (route_map_<safe車號>.html)。支援短網址 vN 與舊車號網址。"""
    import logistics_agent as L
    vid = _fix_vid(vid)
    day = _today_tw()
    day_dir = os.path.join(L.REPORT_DIR, day)
    fp = _resolve_vid_file(vid, day_dir, "route_map")
    if fp:
        return send_file(fp)
    return Response("尚無該車路線地圖。請先傳 Excel 觸發規劃。", mimetype="text/plain; charset=utf-8")


@app.route("/last_result", methods=["GET"])
def view_last_result():
    """最近一次處理結果（Push 失敗時的備援檢視）。"""
    pushed = LAST_RESULT.get("pushed")
    status = "已 Push" if pushed is True else ("Push 失敗" if pushed is False else "尚未處理")
    body = (f"處理時間: {LAST_RESULT.get('ts')}\n"
            f"Push 狀態: {status}\n"
            f"Push 錯誤: {LAST_RESULT.get('error') or '無'}\n"
            f"{'='*30}\n\n{LAST_RESULT.get('text')}")
    return Response(body, mimetype="text/plain; charset=utf-8")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"🤖 LINE 路線機器人啟動中 (port {port})...")
    print(f"   Webhook URL 需填: <你的隧道網址>/callback")
    print(f"   報表檢視: <你的隧道網址>/report  (PUBLIC_URL={PUBLIC_URL or '未設定'})")
    print(f"   功能: 傳 Excel → 自動下載並排版")
    if not CHANNEL_SECRET or not CHANNEL_TOKEN:
        print("   ⚠ 尚未設定 LINE Token，請先編輯 .env 後重啟。")

    # ---- Keep-alive：Render 免費版閒置會睡著，webhook 首次觸發需冷啟動(~數十秒)
    #    導致傳檔後很久才收到結果。每 10 分鐘自 ping 自己 / 端點，保持服務喚醒。
    #    daemon 背景線程，掛了也不影響主服務；本機開發照跑無副作用。
    def _keepalive_loop(interval=600):
        import urllib.request as _ur
        import urllib.error as _ue
        while True:
            try:
                _ur.urlopen(f"http://127.0.0.1:{port}/", timeout=10)
            except Exception:
                pass  # 啟動初期還沒 listen 也無妨，下輪再試
            time.sleep(interval)

    try:
        threading.Thread(target=_keepalive_loop, daemon=True).start()
        print("🔄 keep-alive 已啟動（每 10 分鐘自 ping，避免 Render 冷啟動變慢）")
    except Exception:
        pass

    app.run(host="0.0.0.0", port=port)
