# -*- coding: utf-8 -*-
"""
plan_and_push_line.py — 本機跑派車路線 + 回傳 LINE

用途：
  雲端(Render)對外網路卡住、規劃一直逾時時，改在本機用正常的 Google/OSRM 網路
  跑規劃，算完把摘要 push 回你的 LINE（用本機裝好的 line-bot-sdk）。

用法：
  python plan_and_push_line.py                       # 自動找桌面『路線規劃/每日配送.xlsx』
  python plan_and_push_line.py --data 某檔.xlsx
  python plan_and_push_line.py --auto                # 無車號/強制自動分車
  python plan_and_push_line.py --vehicles 2          # 強制 2 台
  python plan_and_push_line.py --start 9.5           # 出發時間(預設 09:30)

回傳 LINE：
  - 設好 LINE_USER_ID（.env 或環境變數）就自動 push 摘要到該使用者。
  - 沒設 LINE_USER_ID：只把摘要印在終端機，並提示如何設定；報表照常產出。
  - 取自己的 LINE_USER_ID：在 LINE 傳「我的id」給 bot（需先部署含該指令的 line_bot），
    或在 .env 加 LINE_USER_ID=你的id。

依賴：logistics_agent、report、line-bot-sdk（本機已裝）
"""
import os
import sys
import time
import json
import argparse
import traceback
from datetime import datetime, timedelta, timezone

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import logistics_agent as L
import report as report_mod


def _today_tw():
    return (datetime.now(timezone.utc) + timedelta(hours=8)).strftime("%Y-%m-%d")


def _hhmm(h):
    hh = int(h) % 24
    mm = int(round((h - int(h)) * 60))
    if mm == 60:
        hh += 1
        mm = 0
    return f"{hh:02d}:{mm:02d}"


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


def build_summary(result, skipped, fuel_cost, mode_note, public_url, day_dir):
    """照 line_bot._format_result 的摘要結構組 LINE 文字（不含報表產出/快取回寫）。"""
    lines = [
        f"📦 路線規劃完成（本機跑・{result.distance_source}）",
        f"🚚 {mode_note}",
        f"車數 {len(result.routes)} 台 ｜ 總實際里程 {result.total_distance_km:.0f} km ｜ 總瓶數 {int(result.total_load)}",
    ]
    if result.fuel_cost_per_km > 0:
        lines.append(f"⛽ 油資單價 {result.fuel_cost_per_km:.1f} 元/km ｜ 預估總油資 {result.total_fuel_cost:.0f} 元")
    for rt in result.routes:
        v = rt["vehicle"]
        ret = _hhmm(rt["end_hour"])
        if rt.get("on_time"):
            tag = "✅準時回倉"
        else:
            tag = f"⚠回倉 {ret}（指定台數模式不限制{_hhmm(L.TARGET_RETURN_HOUR)}）" if "指定" in mode_note else f"⚠超過{_hhmm(L.TARGET_RETURN_HOUR)}({ret})"
        fuel_txt = f" 油資{rt.get('fuel_cost', 0):.0f}元" if result.fuel_cost_per_km > 0 else ""
        lines.append(f"\n【{v.id}】{len(rt['stops'])}站 {rt['distance_km']:.0f}km {ret}回 {tag}{fuel_txt}")
        for si, s in enumerate(rt["stops"][:5]):
            a, lv = rt["etas"][si]
            lines.append(f"  {si + 1}. {s.name} 到{_hhmm(a)} 離{_hhmm(lv)}")
        if len(rt["stops"]) > 5:
            lines.append(f"  …其餘 {len(rt['stops']) - 5} 站請看報表")
    if skipped:
        lines.append(f"\n⚠ 跳過 {len(skipped)} 筆：")
        lines.append("  " + "; ".join(f"{n}({r})" for n, r in skipped[:10]))

    lines.append(f"\n📁 本機報表資料夾：{day_dir}")
    # 只有本機隧道（localhost / trycloudflare 這類）才附線上連結，
    # 因為本機跑的報表在本機磁碟，Render 網址指的是雲端容器裡的（非本次結果）。
    if public_url and _is_local_tunnel(public_url) and _url_alive(public_url):
        lines.append(f"🌐 也可線上看（本機隧道活著）：")
        lines.append(f"   📄 總表：{public_url}/report")
        lines.append(f"   📊 CSV ：{public_url}/report.csv")
        lines.append(f"   🗺️ 地圖（總圖）：{public_url}/route_map")
        lines.append("      ※ 地圖頁內含「📥 下載地圖 PNG」按鈕")
        lines.append("   🚚 各車獨立報表：")
        for i, rt in enumerate(result.routes, 1):
            vid = rt["vehicle"].id
            lines.append(f"     ・{vid} 報表：{public_url}/report/v{i}")
            lines.append(f"     ・{vid} 路線圖：{public_url}/route_map/v{i}")
    return "\n".join(lines)


def _is_local_tunnel(url: str) -> bool:
    """判斷是否本機暴露的隧道網址（localhost / 127 / trycloudflare.com）。"""
    u = (url or "").lower()
    return ("localhost" in u or "127.0.0.1" in u or "trycloudflare.com" in u)


def _url_alive(url: str) -> bool:
    """本機隧道是否活著（指 localhost / trycloudflare 這類本機暴露網址）。"""
    import urllib.request as _ur
    import urllib.error as _ue
    try:
        _ur.urlopen(f"{url.rstrip('/')}/", timeout=4)
        return True
    except Exception:
        return False


def push_to(user_id, text, attempts=3):
    """照 line_bot._push_to 的邏輯用 line-bot-sdk v3 push（帶重試）。"""
    from linebot.v3.messaging import Configuration, ApiClient, MessagingApi, TextMessage, PushMessageRequest
    token = _ENV.get("LINE_CHANNEL_ACCESS_TOKEN") or os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
    if not token:
        print("⚠ 無 LINE token，跳過 push。")
        return False
    if not user_id:
        print("⚠ 無 LINE_USER_ID，跳過 push。")
        return False
    configuration = Configuration(access_token=token)
    api = MessagingApi(ApiClient(configuration))
    last_err = None
    for i in range(1, attempts + 1):
        try:
            api.push_message(PushMessageRequest(
                to=user_id,
                messages=[TextMessage(text=text[:1900])],
            ))
            print(f"✅ 已 push 到 LINE (user {user_id[:6]}…)")
            return True
        except Exception as e:
            last_err = e
            print(f"⚠ push 第 {i} 次失敗: {type(e).__name__}: {str(e)[:120]}")
            if i < attempts:
                time.sleep(0.5 * i)
    if last_err:
        traceback.print_exc()
    return False


def main():
    ap = argparse.ArgumentParser(description="本機跑派車路線並回傳 LINE")
    ap.add_argument("--data", help="每日配送資料 (xlsx/csv)，不給則自動找桌面『路線規劃』")
    ap.add_argument("--auto", action="store_true", help="強制自動分車（忽略 Excel 車號欄）")
    ap.add_argument("--vehicles", type=int, default=None, help="搭配 --auto：強制分 N 台")
    ap.add_argument("--start", type=float, default=L.DEFAULT_START_HOUR, help="出發時間(24h, 預設9.5)")
    ap.add_argument("--straight", action="store_true", help="強制直線距離/估速")
    args = ap.parse_args()

    global _ENV
    _ENV = _load_env()

    # 找資料檔
    data = args.data
    if not data:
        for cand in ["每日配送.xlsx", "每日配送.csv", "stores.xlsx"]:
            p = os.path.join(L.DATA_DIR, cand)
            if os.path.exists(p):
                data = p
                break
    if not data or not os.path.exists(data):
        print(f"⚠ 找不到資料（預設路徑：{L.DATA_DIR}）。請把填好的『每日配送.xlsx』放進去，或加 --data 指定。")
        return

    use_google = not args.straight
    fuel_cost = L._load_fuel_cost()
    try:
        from data_loader import read_fuel_cost as _rf
        _ef = _rf(data)
        if _ef:
            fuel_cost = _ef
    except Exception:
        pass

    # 決定分車模式
    had_no_vehicle = False
    try:
        _v, _sbv, _sk = L.load(data, depot=L.DEPOT)
        had_no_vehicle = all((getattr(v, "id", "").strip() == "未分車") for v in _v) if _v else False
    except Exception:
        had_no_vehicle = False

    if args.auto or had_no_vehicle:
        mode_note = (f"你指定 {args.vehicles} 台（只求最快回倉）" if args.vehicles else f"由 Agent 依『09:30出車/{_hhmm(L.TARGET_RETURN_HOUR)}回倉』自動安排")
        result, skipped = L.plan_auto_assign(args.start, data, use_google, False,
                                              fuel_cost_per_km=fuel_cost, force_vehicles=args.vehicles)
    else:
        mode_note = "照 Excel 車號排序"
        result, skipped = L.plan(args.start, data, use_google, False, fuel_cost_per_km=fuel_cost)

    if result is None:
        print("⚠ 排程失敗：沒有可規劃的車輛/店家。請檢查 Excel 欄位。")
        return

    # 產報表 + 地圖（本機有 Edge/Chrome，可截 PNG）
    day_dir = os.path.join(L.REPORT_DIR, _today_tw())
    os.makedirs(day_dir, exist_ok=True)
    report_mod.build_html_grouped(result, os.path.join(day_dir, "route_report.html"), meta={"start_hour": args.start})
    report_mod.build_csv_grouped(result, os.path.join(day_dir, "route_report.csv"))
    report_mod.build_html_per_vehicle(result, day_dir, meta={"start_hour": args.start})
    try:
        L.build_map(result, day_dir, use_google)
    except Exception as e:
        print(f"⚠ 地圖產出跳過: {e}")

    # 派車單
    dispatch_dir = os.path.join(L.DISPATCH_DIR, _today_tw())
    os.makedirs(dispatch_dir, exist_ok=True)
    report_mod.build_dispatch_grouped(result, dispatch_dir, meta={"start_hour": args.start})

    # 摘要
    public_url = _ENV.get("PUBLIC_URL") or os.environ.get("PUBLIC_URL", "")
    public_url = (public_url or "").rstrip("/")
    summary = build_summary(result, skipped, fuel_cost, mode_note, public_url, day_dir)

    # 終端機印出
    print("=" * 60)
    print(summary)
    print("=" * 60)
    print(f"📁 報表資料夾：{day_dir}")

    # push 回 LINE
    user_id = _ENV.get("LINE_USER_ID") or os.environ.get("LINE_USER_ID", "")
    if user_id:
        ok = push_to(user_id, summary)
        if not ok:
            print("\n⚠ push 失敗，摘要已在上方終端機。請確認 LINE_USER_ID / token 是否正確。")
    else:
        print("\n💡 未設定 LINE_USER_ID，摘要僅印在終端機（未推播）。")
        print("   取得方式：在 LINE 傳『我的id』給 bot，或在 .env 加 LINE_USER_ID=你的id")
        print("   設定後重跑本腳本即會自動 push 到你的 LINE。")

    # 快取回寫：把本次累積的 geo_cache/matrix_cache 自動 push 回 GitHub，
    # 讓雲端 Render 部署也能拿到最新快取，避免重複呼叫 Google 造成費用暴增。
    try:
        from line_bot import _persist_cache_to_git
        _persist_cache_to_git()
    except Exception as e:
        print(f"⚠ 快取回寫跳過（不影響報表）: {e}")


if __name__ == "__main__":
    main()
