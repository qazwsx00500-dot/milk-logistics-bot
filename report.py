"""
report.py — 路線報表產出

產出：
  - route_report.html : 可列印/手機檢視的路線排程報表 (含每站 ETA、瓶數、下貨時間)
  - route_report.csv  : 給司機/後台用的逐站 CSV

報表欄位（每站）：序號 | 店家 | 地址 | 瓶數 | 下貨(秒) | 預計到達 | 預計離開 | 累計里程
"""

import csv
import json
import os
from datetime import datetime, timedelta


def _hhmm(hour_float):
    """24h 小數 -> HH:MM。"""
    h = int(hour_float) % 24
    m = int(round((hour_float - int(hour_float)) * 60))
    if m == 60:
        h += 1; m = 0
    return f"{h:02d}:{m:02d}"


def _eta_rows(result, depot):
    """把所有路線攤平成逐站報表列。"""
    rows = []
    for ri, rt in enumerate(result.routes, 1):
        v = rt["vehicle"]
        seq = [depot] + rt["stops"] + [depot]
        # 累計里程
        cum_km = 0.0
        prev = depot
        for si, s in enumerate(rt["stops"]):
            # 行車里程 (用該段直線? 這裡用 result 的 distance 反推不精確，改用點對直線估算僅供參考)
            arrive, leave = rt["etas"][si]
            qty = getattr(s, "qty", s.demand)
            svc = int(round(getattr(s, "service_time", 0) or 0))
            rows.append({
                "route_no": ri,
                "vehicle": v.name,
                "seq": si + 1,
                "name": s.name,
                "address": getattr(s, "address", ""),
                "qty": int(qty) if qty == int(qty) else qty,
                "service_sec": svc,
                "arrive": _hhmm(arrive),
                "leave": _hhmm(leave),
            })
    return rows


def build_html(result, depot, out_path, meta=None):
    rows = _eta_rows(result, depot)
    gen = datetime.now().strftime("%Y-%m-%d %H:%M")
    src_map = {"haversine": "直線估算", "osrm": "OSRM 真實道路",
               "fallback": "直線估算(OSRM/Google 失敗降級)", "google": "Google Maps 真實道路"}
    src = src_map.get(result.distance_source, result.distance_source)
    meta = meta or {}

    # 路線摘要卡
    route_cards = ""
    for ri, rt in enumerate(result.routes, 1):
        v = rt["vehicle"]
        end = _hhmm(rt.get("end_hour", 0))
        stops_html = ""
        for si, s in enumerate(rt["stops"]):
            a, lv = rt["etas"][si]
            qty = getattr(s, "qty", s.demand)
            stops_html += (
                f"<tr><td>{si+1}</td><td><b>{s.name}</b><br><span class='addr'>{getattr(s,'address','')}</span></td>"
                f"<td class='num'>{int(qty) if qty==int(qty) else qty}</td>"
                f"<td class='num'>{_hhmm(a)}</td><td class='num'>{_hhmm(lv)}</td></tr>"
            )
        route_cards += f"""
        <div class="card">
          <div class="card-h">路線 {ri} · {v.name}</div>
          <div class="card-meta">
            站數 {len(rt['stops'])} ｜ 里程 {rt['distance_km']:.1f} km ｜
            載重 {rt['load']:.0f} ｜ 回到倉庫 {end}
            {'' if rt['feasible'] else ' <span class="warn">⚠ 超限制</span>'}
          </div>
          <table>
            <thead><tr><th>#</th><th>店家 / 地址</th><th>瓶數</th><th>到達</th><th>離開</th></tr></thead>
            <tbody>{stops_html}</tbody>
          </table>
        </div>"""

    skipped = getattr(result, "skipped", [])
    skip_html = ""
    if skipped:
        items = "".join(f"<li>{name} — {reason}</li>" for name, reason in skipped)
        skip_html = f"<div class='skip'><b>⚠ 未排入 {len(skipped)} 家（地址無法定位或超車隊容量）：</b><ul>{items}</ul></div>"

    html = f"""<!DOCTYPE html>
<html lang="zh-TW"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>鮮奶配送路線報表</title>
<style>
 body{{font-family:-apple-system,"Microsoft JhengHei",sans-serif;margin:0;background:#f4f5f7;color:#222;}}
 .wrap{{max-width:960px;margin:0 auto;padding:18px;}}
 header{{background:#0b6b3a;color:#fff;padding:16px 18px;border-radius:10px;}}
 header h1{{margin:0;font-size:20px;}}
 header .sub{{opacity:.9;font-size:13px;margin-top:4px;}}
 .summary{{display:flex;gap:10px;flex-wrap:wrap;margin:14px 0;}}
 .kpi{{background:#fff;border-radius:10px;padding:10px 14px;flex:1;min-width:120px;box-shadow:0 1px 3px rgba(0,0,0,.08);}}
 .kpi .v{{font-size:20px;font-weight:700;}}
 .kpi .l{{font-size:12px;color:#666;}}
 .card{{background:#fff;border-radius:10px;padding:12px 14px;margin-bottom:12px;box-shadow:0 1px 3px rgba(0,0,0,.08);}}
 .card-h{{font-weight:700;font-size:15px;margin-bottom:4px;}}
 .card-meta{{font-size:12px;color:#555;margin-bottom:8px;}}
 table{{width:100%;border-collapse:collapse;font-size:13px;}}
 th,td{{text-align:left;padding:6px 6px;border-bottom:1px solid #eee;}}
 th{{color:#666;font-weight:600;}}
 td.num,th.num{{text-align:right;font-variant-numeric:tabular-nums;}}
 .addr{{color:#888;font-size:11px;}}
 .warn{{color:#c0392b;font-weight:700;}}
 .ok{{color:#0b6b3a;font-weight:700;}}
 .skip{{background:#fff3cd;border:1px solid #ffe69c;border-radius:10px;padding:10px 14px;font-size:13px;}}
 .skip ul{{margin:6px 0 0;padding-left:18px;}}
 footer{{text-align:center;color:#999;font-size:12px;margin:16px 0;}}
</style></head>
<body><div class="wrap">
<header><h1>🥛 鮮奶配送路線報表</h1>
<div class="sub">出發 {_hhmm(meta.get('start_hour',8.0))} ｜ 距離來源 {src} ｜ 生成 {gen}</div></header>
<div class="summary">
  <div class="kpi"><div class="v">{len(result.routes)}</div><div class="l">出車數</div></div>
  <div class="kpi"><div class="v">{sum(len(r['stops']) for r in result.routes)}</div><div class="l">配送店數</div></div>
  <div class="kpi"><div class="v">{result.total_distance_km:.0f}</div><div class="l">總里程 km</div></div>
  <div class="kpi"><div class="v">{result.total_load:.0f}</div><div class="l">總瓶數</div></div>
</div>
{route_cards}
{skip_html}
<footer>本報表由物流路線規劃 Agent 雛形產出 · 時間含每瓶 10 秒下貨</footer>
</div></body></html>"""
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    return out_path


def build_csv(result, depot, out_path):
    rows = _eta_rows(result, depot)
    with open(out_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["路線", "車輛", "序號", "店家", "地址", "瓶數", "下貨秒數", "預計到達", "預計離開"])
        for r in rows:
            w.writerow([r["route_no"], r["vehicle"], r["seq"], r["name"], r["address"],
                        r["qty"], r["service_sec"], r["arrive"], r["leave"]])
    return out_path


# ---------- 分組版（照車號） ----------

def _hhmm(hour_float):
    h = int(hour_float) % 24
    m = int(round((hour_float - int(hour_float)) * 60))
    if m == 60:
        h += 1; m = 0
    return f"{h:02d}:{m:02d}"


def build_html_grouped(result, out_path, meta=None):
    gen = datetime.now().strftime("%Y-%m-%d %H:%M")
    src_map = {"haversine": "直線估算", "google": "Google Maps 真實道路",
               "osrm": "OSRM 真實道路", "fallback": "直線估算(降級)"}
    src = src_map.get(result.distance_source, result.distance_source)
    meta = meta or {}
    start = meta.get("start_hour", 8.0)

    cards = ""
    for rt in result.routes:
        v = rt["vehicle"]
        ret = _hhmm(rt["end_hour"])
        on_time = rt.get("on_time", True)
        ret_tag = '<span class="ok">✅ 準時回倉</span>' if on_time else f'<span class="warn">⚠ 超過 17:30（{ret}）</span>'
        fuel_txt = f" ｜ 油資 {rt.get('fuel_cost', 0):.0f} 元" if result.fuel_cost_per_km > 0 else ""
        rows_html = ""
        for si, s in enumerate(rt["stops"]):
            a, lv = rt["etas"][si]
            qty = int(s.demand) if s.demand == int(s.demand) else s.demand
            rows_html += (
                f"<tr><td>{si+1}</td><td><b>{s.name}</b><br><span class='addr'>{s.address}</span></td>"
                f"<td class='num'>{qty}</td>"
                f"<td class='num'>{_hhmm(a)}</td><td class='num'>{_hhmm(lv)}</td></tr>"
            )
        cards += f"""
        <div class="card">
          <div class="card-h">{v.id} 路線</div>
          <div class="card-meta">起點 {v.start_addr or '—'} ｜ {len(rt['stops'])} 站 ｜
            實際里程 {rt['distance_km']:.1f} km ｜ 總瓶數 {rt['load']:.0f} ｜
            預計回到起點 {ret}（目標 17:30）{ret_tag}{fuel_txt}</div>
          <table>
            <thead><tr><th>#</th><th>店家 / 地址</th><th>瓶數</th><th>到店</th><th>離店</th></tr></thead>
            <tbody>{rows_html}</tbody>
          </table>
        </div>"""

    skipped = getattr(result, "skipped", [])
    skip_html = ""
    if skipped:
        items = "".join(f"<li>{name} — {reason}</li>" for name, reason in skipped)
        skip_html = f"<div class='skip'><b>⚠ 跳過 {len(skipped)} 筆（地址無法定位）：</b><ul>{items}</ul></div>"

    html = f"""<!DOCTYPE html>
<html lang="zh-TW"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>鮮奶配送路線報表</title>
<style>
 body{{font-family:-apple-system,"Microsoft JhengHei",sans-serif;margin:0;background:#f4f5f7;color:#222;}}
 .wrap{{max-width:960px;margin:0 auto;padding:18px;}}
 header{{background:#0b6b3a;color:#fff;padding:16px 18px;border-radius:10px;}}
 header h1{{margin:0;font-size:20px;}}
 header .sub{{opacity:.9;font-size:13px;margin-top:4px;}}
 .summary{{display:flex;gap:10px;flex-wrap:wrap;margin:14px 0;}}
 .kpi{{background:#fff;border-radius:10px;padding:10px 14px;flex:1;min-width:120px;box-shadow:0 1px 3px rgba(0,0,0,.08);}}
 .kpi .v{{font-size:20px;font-weight:700;}}
 .kpi .l{{font-size:12px;color:#666;}}
 .card{{background:#fff;border-radius:10px;padding:12px 14px;margin-bottom:12px;box-shadow:0 1px 3px rgba(0,0,0,.08);}}
 .card-h{{font-weight:700;font-size:15px;margin-bottom:4px;}}
 .card-meta{{font-size:12px;color:#555;margin-bottom:8px;}}
 table{{width:100%;border-collapse:collapse;font-size:13px;}}
 th,td{{text-align:left;padding:6px 6px;border-bottom:1px solid #eee;}}
 th{{color:#666;font-weight:600;}}
 td.num,th.num{{text-align:right;font-variant-numeric:tabular-nums;}}
 .addr{{color:#888;font-size:11px;}}
 .skip{{background:#fff3cd;border:1px solid #ffe69c;border-radius:10px;padding:10px 14px;font-size:13px;}}
 .skip ul{{margin:6px 0 0;padding-left:18px;}}
 footer{{text-align:center;color:#999;font-size:12px;margin:16px 0;}}
</style></head>
<body><div class="wrap">
<header><h1>🥛 鮮奶配送路線報表</h1>
<div class="sub">出發 {_hhmm(start)} ｜ 距離來源 {src} ｜ 生成 {gen}</div></header>
<div class="summary">
  <div class="kpi"><div class="v">{len(result.routes)}</div><div class="l">出車數</div></div>
  <div class="kpi"><div class="v">{sum(len(r['stops']) for r in result.routes)}</div><div class="l">配送店數</div></div>
  <div class="kpi"><div class="v">{result.total_distance_km:.0f}</div><div class="l">總實際里程 km</div></div>
  <div class="kpi"><div class="v">{result.total_load:.0f}</div><div class="l">總瓶數</div></div>
  {f'<div class="kpi"><div class="v">{result.total_fuel_cost:.0f}</div><div class="l">預估總油資 元</div></div>' if result.fuel_cost_per_km > 0 else ''}
</div>
{cards}
{skip_html}
<footer>本報表由物流路線規劃 Agent 產出 · 下貨時間按每瓶 10 秒計算</footer>
</div></body></html>"""
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    return out_path


def build_csv_grouped(result, out_path):
    with open(out_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["車號", "序號", "店家", "地址", "瓶數", "下貨秒數", "預計到店", "預計離店"])
        for rt in result.routes:
            v = rt["vehicle"]
            for si, s in enumerate(rt["stops"]):
                a, lv = rt["etas"][si]
                qty = int(s.demand) if s.demand == int(s.demand) else s.demand
                svc = int(round(s.service_time))
                w.writerow([v.id, si + 1, s.name, s.address, qty, svc, _hhmm(a), _hhmm(lv)])
    # 總計補一張摘要表(同檔下方另寫會蓋掉，這裡用第二種方式：寫入 summary 區)
    with open(out_path, "a", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow([])
        w.writerow(["=== 路線總計 ==="])
        w.writerow(["出車數", len(result.routes)])
        w.writerow(["總實際里程_km", f"{result.total_distance_km:.1f}"])
        w.writerow(["總瓶數", f"{result.total_load:.0f}"])
        if result.fuel_cost_per_km > 0:
            w.writerow(["油資單價_元每km", f"{result.fuel_cost_per_km:.1f}"])
            w.writerow(["預估總油資_元", f"{result.total_fuel_cost:.0f}"])
        for rt in result.routes:
            ret = _hhmm(rt["end_hour"])
            status = "準時回倉" if rt.get("on_time", True) else f"超過17:30({ret})"
            w.writerow([rt["vehicle"].id, "實際里程_km", f"{rt['distance_km']:.1f}",
                        "回倉", ret, status,
                        f"油資_{rt.get('fuel_cost',0):.0f}元" if result.fuel_cost_per_km > 0 else ""])
    return out_path
