"""
onedrive_sync.py — 把「當日報表」上傳到使用者的 OneDrive，
讓本機 OneDrive 同步桌面自動出現報表檔（桌面/當日車輛報表/YYYY-MM-DD/）。

用途：LINE bot 跑在 Render 等雲端時，報表原本只寫在雲端容器硬碟。
      這支模組在規劃完成後，把報表上傳到你的 OneDrive（個人雲端硬碟），
      你本機「OneDrive 同步桌面」就會自動把檔案拉下來 → 桌面看得到。

需要的 .env 設定：
  ONEDRIVE_CLIENT_ID      = Azure 應用程式 (用戶端) 識別碼
  ONEDRIVE_CLIENT_SECRET  = 用戶端密碼
  ONEDRIVE_REFRESH_TOKEN  = 第一次授權後取得的重新整理權杖（用 onedrive_auth.py 取得）
  ONEDRIVE_TENANT         = consumers（個人 OneDrive）/ 企业或學校填 tenant id 或 common
  ONEDRIVE_REPORT_REL     = 桌面/當日車輛報表   （OneDrive 內的相對路徑，對應本機同步資料夾）

注意：沒有設定上述憑證時，所有函式安全跳過（不影響 LINE 回傳）。
"""

import os
import json
import urllib.parse
import urllib.request
import urllib.error

HERE = os.path.dirname(os.path.abspath(__file__))


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


def get_access_token(env=None):
    """用 refresh_token 換_access_token；沒設定憑證就回 None。"""
    env = env or _load_env()
    cid = env.get("ONEDRIVE_CLIENT_ID")
    sec = env.get("ONEDRIVE_CLIENT_SECRET")
    rt = env.get("ONEDRIVE_REFRESH_TOKEN")
    tenant = env.get("ONEDRIVE_TENANT", "consumers")
    if not (cid and sec and rt):
        return None
    url = f"https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"
    body = urllib.parse.urlencode({
        "client_id": cid,
        "client_secret": sec,
        "refresh_token": rt,
        "grant_type": "refresh_token",
        "scope": "Files.ReadWrite offline_access",
    }).encode("utf-8")
    req = urllib.request.Request(
        url, data=body,
        headers={"Content-Type": "application/x-www-form-urlencoded"})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return json.loads(r.read().decode())
    except urllib.error.HTTPError as e:
        print(f"⚠ OneDrive 換發 token 失敗 ({e.code}): {e.read().decode()[:200]}")
    except Exception as e:
        print(f"⚠ OneDrive 換發 token 例外: {e}")
    return None


def _api(method, url, token, data=None, headers=None):
    h = {"Authorization": f"Bearer {token}"}
    if headers:
        h.update(headers)
    if data is not None and "Content-Type" not in h:
        h["Content-Type"] = "application/json"
    req = urllib.request.Request(url, data=data, headers=h, method=method)
    with urllib.request.urlopen(req, timeout=60) as r:
        ct = r.headers.get("Content-Type", "")
        if ct.startswith("application/json"):
            return json.loads(r.read().decode())
        return r.read()


def _ensure_folder(token, rel_parts):
    """確保 OneDrive 上 rel_parts 路徑的資料夾存在（逐層建立）。"""
    for i in range(1, len(rel_parts) + 1):
        sub = "/".join(rel_parts[:i])
        url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{urllib.parse.quote(sub, safe='/')}:"
        try:
            _api("GET", url, token)
            continue
        except urllib.error.HTTPError as e:
            if e.code != 404:
                raise
        # 不存在 → 在上一層 children 建立
        parent = "/".join(rel_parts[:i - 1])
        if parent:
            parent_url = (f"https://graph.microsoft.com/v1.0/me/drive/root:"
                          f"/{urllib.parse.quote(parent, safe='/')}:/children")
        else:
            parent_url = "https://graph.microsoft.com/v1.0/me/drive/root/children"
        body = json.dumps({
            "name": rel_parts[i - 1],
            "folder": {},
            "@microsoft.graph.conflictBehavior": "replace",
        }).encode("utf-8")
        _api("POST", parent_url, token, data=body)


def upload_file(token, local_path, onedrive_rel):
    """上傳單一檔案到 OneDrive 路徑 onedrive_rel/（檔名取自 local_path）。"""
    fname = os.path.basename(local_path)
    rel = "/".join([p for p in onedrive_rel.split("/") if p] + [fname])
    url = f"https://graph.microsoft.com/v1.0/me/drive/root:/{urllib.parse.quote(rel, safe='/')}:/content"
    with open(local_path, "rb") as f:
        data = f.read()
    return _api("PUT", url, token, data=data,
                headers={"Content-Type": "application/octet-stream"})


def upload_report_dir(day_dir, day, env=None):
    """把 day_dir 下的報表上傳到 OneDrive/{REPORT_REL}/{day}/。回傳是否成功。"""
    env = env or _load_env()
    tok = get_access_token(env)
    if not tok or "access_token" not in tok:
        return False
    token = tok["access_token"]
    rel_base = env.get("ONEDRIVE_REPORT_REL", "桌面/當日車輛報表")
    rel_parts = [p for p in (rel_base.split("/") + [day]) if p]
    try:
        _ensure_folder(token, rel_parts)
    except Exception as e:
        print(f"⚠ OneDrive 建立資料夾失敗: {e}")
        return False
    ok = True
    for fn in ("route_report.html", "route_report.csv", "route_map.html"):
        fp = os.path.join(day_dir, fn)
        if os.path.exists(fp):
            try:
                upload_file(token, fp, "/".join(rel_parts))
                print(f"   ☁ 已上傳 OneDrive: {'/'.join(rel_parts)}/{fn}")
            except Exception as e:
                print(f"⚠ OneDrive 上傳 {fn} 失敗: {e}")
                ok = False
    return ok
