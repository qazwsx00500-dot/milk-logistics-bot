"""
map_capture.py — 用本機 Edge/Chrome headless 把 route_map.html 截成 PNG。

本機有 Edge/Chrome 就不需安裝任何套件；雲端(Render)通常沒有瀏覽器，
capture_map_png 會回 None（呼叫端應優雅降級，不崩）。

用法：
    from map_capture import capture_map_png
    png = capture_map_png("route_map.html", "route_map.png")  # 成功回路徑，失敗回 None
"""

import os
import shutil
import subprocess
import sys

# 常見瀏覽器路徑（Windows 優先；Linux 次之）
_CANDIDATES = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    "/usr/bin/google-chrome", "/usr/bin/chromium", "/usr/bin/chromium-browser",
]


def _find_browser():
    for p in _CANDIDATES:
        if os.path.exists(p):
            return p
    # PATH 裡找
    for name in ("msedge", "google-chrome", "chrome", "chromium", "chromium-browser"):
        found = shutil.which(name)
        if found:
            return found
    return None


def capture_map_png(html_path, png_path, width=1280, height=900, timeout=90):
    """把 html_path 截圖存成 png_path。成功回 png_path，失敗回 None。"""
    if not os.path.exists(html_path):
        return None
    browser = _find_browser()
    if not browser:
        sys.stderr.write("[map_capture] 找不到瀏覽器，跳過路線圖截圖\n")
        return None
    out_abs = os.path.abspath(png_path)
    file_url = "file:///" + os.path.abspath(html_path).replace("\\", "/")
    cmd = [
        browser, "--headless", "--no-sandbox", "--disable-gpu",
        "--hide-scrollbars", f"--window-size={width},{height}",
        f"--screenshot={out_abs}", file_url,
    ]
    try:
        subprocess.run(cmd, check=True, timeout=timeout,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as e:
        sys.stderr.write(f"[map_capture] 截圖失敗: {type(e).__name__}: {e}\n")
        return None
    if os.path.exists(out_abs) and os.path.getsize(out_abs) > 0:
        return out_abs
    return None


if __name__ == "__main__":
    if len(sys.argv) > 1:
        out = capture_map_png(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else "route_map.png")
        print("OK" if out else "SKIP", out)
