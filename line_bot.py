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
import traceback
from datetime import datetime

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


# ---- 報表檢視路由（直接讀今天日期資料夾的檔案） ----
def _today_report(which):
    """which: 'html' | 'csv' → 回傳今天報表檔路徑或 None。"""
    import logistics_agent as L
    day_dir = os.path.join(L.REPORT_DIR, datetime.now().strftime("%Y-%m-%d"))
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

        # 把 rows 寫成臨時標準檔（供 data_loader 讀）
        if rows is not None and data_path is None:
            import excel_normalizer as en
            # 無車號 → 先填 車01 (auto_assign_vehicles 保守起點)
            if all(r[0] == "" for r in rows):
                rows = en.auto_assign_vehicles(rows)
            import openpyxl as _ox
            from openpyxl import Workbook as _WB
            tmp = os.path.join(HERE, "_normalized_tmp.xlsx")
            wb = _WB(); ws = wb.active; ws.title = "每日配送"
            ws.append(["車號", "店家名稱", "店家地址", "瓶數"])
            for veh, n, a, q in rows:
                ws.append([veh, n, a, q])
            wb.save(tmp)
            data_path = tmp

        if not data_path or not os.path.exists(data_path):
            return (f"⚠ 找不到資料檔。\n請直接把 Excel 傳給我，或先放到『路線規劃』資料夾。")

        result, skipped = L.plan(start_hour, data_path, use_google, no_google, fuel_cost_per_km=fuel_cost)
        if result is None:
            return "⚠ 排程失敗：沒有可規劃的車輛/店家。請檢查 Excel 欄位。"

        # 無車號 → 依「時間窗 + 最短距離」自動分車 (最多3台, 一台跑不完才加車)
        if rows is not None and all(r[0] == "" for r in rows):
            if len(result.routes) == 1 and not result.routes[0].get("on_time"):
                import auto_router
                stops = result.routes[0]["stops"]
                depot = L.DEPOT
                groups = auto_router.decide_groups(
                    stops, depot, start_hour, L.TARGET_RETURN_HOUR, max_vehicles=3)
                if len(groups) > 1:
                    # 按群重寫臨時 xlsx (車01/車02/車03)
                    import openpyxl as _ox
                    from openpyxl import Workbook as _WB
                    tmp = os.path.join(HERE, "_normalized_tmp.xlsx")
                    wb = _WB(); ws = wb.active; ws.title = "每日配送"
                    ws.append(["車號", "店家名稱", "店家地址", "瓶數"])
                    for gi, g in enumerate(groups, 1):
                        veh = f"車{gi:02d}"
                        for si in g:
                            s = stops[si]
                            ws.append([veh, s.name, s.address, int(s.demand or 0)])
                    wb.save(tmp)
                    result, skipped = L.plan(start_hour, tmp, use_google, no_google, fuel_cost_per_km=fuel_cost)

        # 產報表到 日期子資料夾
        day_dir = os.path.join(L.REPORT_DIR, datetime.now().strftime("%Y-%m-%d"))
        os.makedirs(day_dir, exist_ok=True)
        report_mod.build_html_grouped(result, os.path.join(day_dir, "route_report.html"),
                                      meta={"start_hour": start_hour})
        report_mod.build_csv_grouped(result, os.path.join(day_dir, "route_report.csv"))
        L.build_map(result, day_dir, use_google)

        # 文字摘要
        veh_note = ""
        if rows is not None and all(r[0] == "" for r in rows):
            veh_note = f"\n🚚 由 Agent 依『09:30出車/17:30回倉』自動安排 {len(result.routes)} 台車（最多3台，一台跑不完才加車）"
        lines = [f"📦 路線規劃完成（{result.distance_source}）",
                 f"出發 09:30 ｜ 目標回倉 17:30{veh_note}",
                 f"車數 {len(result.routes)} 台 ｜ 總實際里程 {result.total_distance_km:.0f} km ｜ 總瓶數 {int(result.total_load)}"]
        if result.fuel_cost_per_km > 0:
            lines.append(f"⛽ 油資單價 {result.fuel_cost_per_km:.1f} 元/km ｜ 預估總油資 {result.total_fuel_cost:.0f} 元")
        for rt in result.routes:
            v = rt["vehicle"]
            ret = _hhmm(rt["end_hour"])
            tag = "✅準時回倉" if rt.get("on_time") else f"⚠超過17:30({ret})"
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
            lines.append(f"\n📄 報表：{PUBLIC_URL}/report")
            lines.append(f"📊 CSV ：{PUBLIC_URL}/report.csv")
            lines.append(f"🗺️ 地圖：{PUBLIC_URL}/route_map")
        else:
            lines.append(f"\n📁 報表已產出：{day_dir}")

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
        rows, skipped, out_path = en.normalize_excel(tmp_in, L.DATA_DIR)
        if not rows:
            return ("⚠ 無法從這個 Excel 讀到店家資料。\n"
                    "需要的欄位：店家名稱 / 店家地址 / 瓶數（車號可省略，我會自動安排）。")
    except Exception as e:
        traceback.print_exc()
        return f"⚠ 轉換 Excel 失敗：{type(e).__name__}: {str(e)[:100]}"

    # 3) 自動規劃（rows 傳入 → 無車號時 Agent 自動分車）
    summary = (f"📥 已收到「{fname}」\n"
               f"✅ 已自動轉成標準每日配送表並存檔：\n"
               f"   {os.path.basename(out_path)}\n"
               f"   （{len(rows)} 筆店家資料，{len(skipped)} 筆跳過）\n\n")
    return summary + run_plan(rows=rows)


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

def _push_to(user_id, text):
    """用 Push 推結果給使用者（繞過 reply_token 1 秒過期限制）。"""
    global LAST_RESULT
    try:
        from linebot.v3.messaging.models import PushMessageRequest
        if not user_id:
            raise ValueError("user_id 為空，無法 Push（可能事件結構不含 source.user_id）")
        messaging_api.push_message(
            PushMessageRequest(
                to=user_id,
                messages=[TextMessage(text=text[:1900])],
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
            result = handle_text(payload)
        elif kind == "file":
            result = handle_file(payload["event"], payload["file_msg"])
        else:
            result = "⚠ 不支援的訊息類型。"
    except Exception as e:
        traceback.print_exc()
        result = f"⚠ 處理時發生錯誤：{type(e).__name__}: {str(e)[:120]}"
    LAST_RESULT["ts"] = _dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    LAST_RESULT["text"] = result
    _push_to(user_id, result)

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

    for event in events:
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
    day = datetime.now().strftime("%Y-%m-%d")
    p = os.path.join(L.REPORT_DIR, day, "route_map.html")
    if os.path.exists(p):
        return send_file(p)
    return Response("尚無路線地圖。請先傳 Excel 觸發規劃。", mimetype="text/plain; charset=utf-8")


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
    app.run(host="0.0.0.0", port=port)
