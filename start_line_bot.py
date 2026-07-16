"""
start_line_bot.py — 一鍵啟動 LINE 鮮奶物流機器人

功能：
  1. 啟動 cloudflared 隧道 (本機 5000 → 公網 https)
  2. 監控 cloudflared 輸出，抓到新網址後自動寫進 .env 的 PUBLIC_URL
  3. 啟動 line_bot.py (Flask Webhook)
  4. 印出 Webhook URL 讓你複製到 LINE Console

用法：雙擊「啟動LINE機器人.bat」即可（它會呼叫本腳本）。
注意：關閉本視窗 = 全部停止（LINE 就連不上）。
"""

import os
import sys
import re
import time
import subprocess
import threading

HERE = os.path.dirname(os.path.abspath(__file__))


def update_public_url(url: str):
    """把 PUBLIC_URL 寫進 .env（保留其他設定）。"""
    p = os.path.join(HERE, ".env")
    lines = []
    if os.path.exists(p):
        lines = open(p, encoding="utf-8").read().splitlines()
    kept = [l for l in lines if not l.strip().startswith("PUBLIC_URL=")]
    kept.append(f"PUBLIC_URL={url}")
    with open(p, "w", encoding="utf-8") as f:
        f.write("\n".join(kept) + "\n")
    print(f"[OK] 已更新 .env 的 PUBLIC_URL = {url}")


def watch_tunnel(proc):
    """監控 cloudflared 輸出，抓 https://xxx.trycloudflare.com 網址。"""
    pattern = re.compile(r"https://[a-z0-9\-]+\.trycloudflare\.com")
    try:
        for line in iter(proc.stdout.readline, b""):
            text = line.decode("utf-8", "replace")
            m = pattern.search(text)
            if m:
                url = m.group(0).rstrip("/")
                print(f"\n{'='*50}\n🌐 隧道網址: {url}\n📋 Webhook URL (貼到 LINE Console): {url}/callback\n{'='*50}\n")
                update_public_url(url)
                break
    except Exception as e:
        print("監控隧道時出錯:", e)


def main():
    print("=" * 50)
    print("  鮮奶物流 LINE 機器人 啟動中...")
    print("=" * 50)

    # 1) 啟動 cloudflared 隧道
    cf = subprocess.Popen(
        [os.path.join(HERE, "cloudflared.exe"), "tunnel", "--url", "http://localhost:5000"],
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT, cwd=HERE,
    )
    print("✅ cloudflared 隧道已啟動，等待網址生成...")

    # 2) 背景監控抓網址
    watcher = threading.Thread(target=watch_tunnel, args=(cf,), daemon=True)
    watcher.start()

    # 3) 啟動 line_bot（前景，保持視窗活著）
    print("🤖 啟動 line_bot (port 5000)...")
    time.sleep(2)
    lb = subprocess.Popen([sys.executable, "line_bot.py"], cwd=HERE)
    try:
        lb.wait()
    except KeyboardInterrupt:
        pass
    finally:
        cf.terminate()


if __name__ == "__main__":
    main()
