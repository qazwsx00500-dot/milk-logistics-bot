"""
sync_from_render.py — 把 Render 雲端「今天」的報表/派車單/地圖抓回本機桌面。

用途：你走 LINE → Render 雲端跑，檔案寫在雲端容器；本機桌面看不到。
      執行本程式即可把雲端當天成果抓下來，存到你 OneDrive 桌面對應資料夾，
      不需要 OneDrive API 授權。

用法：
    python sync_from_render.py
    python sync_from_render.py --date 2026-07-17   # 資料夾用指定日期命名（預設今天）

抓取來源（Render 端點，只會有「今天」那一份）：
    /report       → route_report.html
    /report.csv   → route_report.csv
    /route_map    → route_map.html
    /dispatch     → dispatch.html
    /dispatch.csv → dispatch.csv

存放位置（與本機直跑一致）：
    當日車輛報表/<日期>/ ← route_report.html/.csv, route_map.html
    當日派車單/<日期>/   ← dispatch.html, dispatch.csv
"""

import argparse
import os
import ssl
import sys
import urllib.error
import urllib.request
from datetime import datetime

import logistics_agent as L

BASE = os.environ.get("RENDER_BASE_URL", "https://milk-logistics-bot.onrender.com")

# (端點, 目標資料夾屬性, 檔名, 是否文字類判斷用)
JOBS = [
    ("/report",      "REPORT_DIR",   "route_report.html"),
    ("/report.csv",  "REPORT_DIR",   "route_report.csv"),
    ("/route_map",   "REPORT_DIR",   "route_map.html"),
    ("/dispatch",    "DISPATCH_DIR", "dispatch.html"),
    ("/dispatch.csv","DISPATCH_DIR", "dispatch.csv"),
    ("/workbook",    "DISPATCH_DIR", "整合報表.xlsx"),
    ("/map_png",     "DISPATCH_DIR", "route_map.png"),
]

# 雲端「尚未產生」時回的提示字樣（抓到這個就不是真檔，跳過）
_EMPTY_HINTS = ("尚未產生報表", "尚未產生派車單", "尚無路線地圖")


def _fetch(url, timeout=40):
    ctx = ssl.create_default_context()
    req = urllib.request.Request(url, headers={"User-Agent": "sync-from-render/1.0"})
    with urllib.request.urlopen(req, timeout=timeout, context=ctx) as r:
        return r.status, r.read()


def main():
    ap = argparse.ArgumentParser(description="把 Render 雲端當天報表/派車單抓回本機桌面")
    ap.add_argument("--date", help="資料夾日期(YYYY-MM-DD)，預設今天")
    ap.add_argument("--base", help=f"Render 網址，預設 {BASE}")
    args = ap.parse_args()

    base = (args.base or BASE).rstrip("/")
    day = args.date or datetime.now().strftime("%Y-%m-%d")
    dirs = {"REPORT_DIR": L.REPORT_DIR, "DISPATCH_DIR": L.DISPATCH_DIR}

    print(f"🔄 從 {base} 抓取 {day} 的雲端成果 …")
    ok, empty, fail = 0, 0, 0
    for path, dir_attr, fname in JOBS:
        url = base + path
        try:
            status, body = _fetch(url)
        except urllib.error.HTTPError as e:
            print(f"  ✗ {path} → HTTP {e.code}")
            fail += 1
            continue
        except Exception as e:
            print(f"  ✗ {path} → {type(e).__name__}: {str(e)[:60]}")
            fail += 1
            continue

        # 判斷是否為「尚未產生」提示（短文字）
        head = body[:80].decode("utf-8", errors="ignore")
        if any(h in head for h in _EMPTY_HINTS):
            print(f"  ⚠ {path} → 雲端尚未產生（跳過）")
            empty += 1
            continue

        day_dir = os.path.join(dirs[dir_attr], day)
        os.makedirs(day_dir, exist_ok=True)
        out = os.path.join(day_dir, fname)
        with open(out, "wb") as f:
            f.write(body)
        print(f"  ✓ {path} → {out}  ({len(body):,} bytes)")
        ok += 1

    print(f"\n完成：成功 {ok}、尚未產生 {empty}、失敗 {fail}")
    if ok:
        print(f"📁 報表：{os.path.join(L.REPORT_DIR, day)}")
        print(f"🚚 派車單：{os.path.join(L.DISPATCH_DIR, day)}")
    if empty and not ok:
        print("（雲端這天還沒跑過規劃：先在 LINE 傳一次 Excel，再執行本程式）")
    return 0 if fail == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
