"""
onedrive_auth.py — 第一次授權用：取得 OneDrive refresh_token 寫進 .env。

步驟：
  1) 到 https://portal.azure.com → Microsoft Entra ID → 應用程式註冊 → 新增註冊
       - 支援的帳戶類型：選「任何組織目錄中的帳戶及個人 Microsoft 帳戶」
       - 重新導向 URI：選「Web」，填 http://localhost:8080
       - 註冊後取得「應用程式 (用戶端) 識別碼」與「用戶端密碼」
  2) API 權限 → 新增 «Microsoft Graph» → 委託的權限 → Files.ReadWrite + offline_access → 新增並「授與管理員同意」(個人帳號則使用者本人同意即可)
  3) 把 CLIENT_ID / CLIENT_SECRET 先寫進 .env：
         ONEDRIVE_CLIENT_ID=xxxx
         ONEDRIVE_CLIENT_SECRET=yyyy
         ONEDRIVE_TENANT=consumers
  4) 執行： python onedrive_auth.py
       → 終端機會印出一個網址，用瀏覽器開啟並登入授權
       → 授權後瀏覽器會跳到 localhost:8080（本機小伺服器會自動接住）
       → 終端機印出 refresh_token，自動追加到 .env 的 ONEDRIVE_REFRESH_TOKEN
  5) 之後 line_bot.py 跑完會自動上傳報表到 OneDrive。

安全提示：refresh_token 等同長期存取權，請勿外洩；若曾貼到對話/第三方，請到 Azure 先「刪除用戶端密碼」重發。
"""

import os
import sys
import urllib.parse
import urllib.request
import threading
import http.server
import socketserver
from onedrive_sync import _load_env

HERE = os.path.dirname(os.path.abspath(__file__))
REDIRECT_URI = "http://localhost:8080"
SCOPE = "Files.ReadWrite offline_access"


def build_auth_url(client_id):
    params = {
        "client_id": client_id,
        "response_type": "code",
        "redirect_uri": REDIRECT_URI,
        "scope": SCOPE,
        "response_mode": "query",
    }
    return ("https://login.microsoftonline.com/consumers/oauth2/v2.0/authorize?"
            + urllib.parse.urlencode(params))


def exchange_code(client_id, client_secret, code):
    url = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token"
    body = urllib.parse.urlencode({
        "client_id": client_id,
        "client_secret": client_secret,
        "code": code,
        "redirect_uri": REDIRECT_URI,
        "grant_type": "authorization_code",
        "scope": SCOPE,
    }).encode("utf-8")
    req = urllib.request.Request(
        url, data=body,
        headers={"Content-Type": "application/x-www-form-urlencoded"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.read().decode()


class _Handler(http.server.BaseHTTPRequestHandler):
    def do_GET(self):
        q = urllib.parse.urlparse(self.path).query
        code = urllib.parse.parse_qs(q).get("code", [""])[0]
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        if code:
            self.server.auth_code = code
            self.wfile.write("授權成功！請回到終端機繼續。".encode("utf-8"))
        else:
            self.wfile.write("未收到授權碼，請重試。".encode("utf-8"))

    def log_message(self, *a):
        pass


def main():
    env = _load_env()
    cid = env.get("ONEDRIVE_CLIENT_ID")
    sec = env.get("ONEDRIVE_CLIENT_SECRET")
    if not (cid and sec):
        print("⚠ 請先在 .env 設定 ONEDRIVE_CLIENT_ID 與 ONEDRIVE_CLIENT_SECRET")
        return

    code_holder = {}

    class _S(socketserver.TCPServer):
        allow_reuse_address = True

    httpd = _S(("localhost", 8080), _Handler)
    httpd.auth_code = ""

    def serve():
        httpd.handle_request()

    t = threading.Thread(target=serve, daemon=True)
    t.start()

    auth_url = build_auth_url(cid)
    print("\n請用瀏覽器開啟以下網址並登入授權：\n")
    print(auth_url, "\n")
    print("（授權後瀏覽器會跳到 localhost:8080，本機會自動接住）...")

    t.join(timeout=180)
    code = httpd.auth_code
    if not code:
        print("⚠ 等待授權逾時，未取得 code。")
        return

    tok = exchange_code(cid, sec, code)
    try:
        rt = __import__("json").loads(tok)["refresh_token"]
    except Exception:
        print("⚠ 換發 token 失敗：", tok[:200])
        return

    # 寫入 .env
    env_path = os.path.join(HERE, ".env")
    lines = []
    if os.path.exists(env_path):
        lines = open(env_path, encoding="utf-8").read().splitlines()
    lines = [ln for ln in lines if not ln.startswith("ONEDRIVE_REFRESH_TOKEN=")]
    lines.append(f"ONEDRIVE_REFRESH_TOKEN={rt}")
    lines.append("ONEDRIVE_TENANT=consumers")
    lines.append("ONEDRIVE_REPORT_REL=桌面/當日車輛報表")
    open(env_path, "w", encoding="utf-8").write("\n".join(lines) + "\n")
    print("\n✅ refresh_token 已寫入 .env 的 ONEDRIVE_REFRESH_TOKEN")
    print("   之後 line_bot.py 跑完會自動把報表上傳到 OneDrive → 本機桌面同步出現。")


if __name__ == "__main__":
    main()
