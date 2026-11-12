# -*- coding: utf-8 -*-
"""涉诉画像工作台：本地网页版。

双击「启动工作台.bat」启动，浏览器自动打开 http://127.0.0.1:8777
"""
import base64
import json
import os
import re
import shutil
import subprocess
import sys
import threading
import webbrowser
from datetime import datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import quote, unquote

BASE = Path(__file__).resolve().parent
CORE = BASE / "core"
CACHE = BASE / "缓存"
OUT = BASE / "输出"
LIST = BASE / "名单"
AUTH_FILE = BASE / "凭证.txt"
PKU_FILE = BASE / "法宝凭证.txt"
CFG_DIR = BASE / "配置"
UI_FILE = BASE / "ui.html"

PORT = int(os.environ.get("HANA_WB_PORT", "8777"))
RUNTIME = BASE / "runtime"
NODE_DIR = RUNTIME / "node"
PY_EXE = RUNTIME / "python" / "python.exe"
BUNDLED = (NODE_DIR / "tyc.cmd").exists()
TYC_CFG = Path.home() / ".tyc" / "config.json"

if BUNDLED:
    os.environ["PATH"] = str(NODE_DIR) + os.pathsep + os.environ.get("PATH", "")


def tyc_env():
    e = dict(os.environ)
    if BUNDLED:
        e["PATH"] = str(NODE_DIR) + os.pathsep + e.get("PATH", "")
    return e


def python_exe():
    if PY_EXE.exists():
        return str(PY_EXE)
    return sys.executable


def load_cfg():
    p = CFG_DIR / "config.toml"
    if p.exists():
        try:
            import tomllib
            with p.open("rb") as f:
                return tomllib.load(f)
        except Exception:
            pass
    return {"window": {"start": "2024-01-01", "end": "2026-10-31"},
            "run": {"floor": 0, "group_threshold": 1000}}


def _read_secret(path):
    if path.exists():
        for line in path.read_text(encoding="utf-8-sig").splitlines():
            t = line.strip()
            if t and not t.startswith("#"):
                return t
    return ""


def read_auth():
    return _read_secret(AUTH_FILE) or os.environ.get("QCC_AUTH", "").strip()


def read_pku():
    return _read_secret(PKU_FILE) or os.environ.get("PKULAW_AUTH", "").strip()


def save_secret(path, value):
    value = (value or "").strip()
    if not value:
        if path.exists():
            try:
                path.unlink()
            except OSError:
                pass
        return
    path.write_text(value + "\n", encoding="utf-8")


def mask_token(t):
    if not t:
        return ""
    head = t[:12] if len(t) >= 12 else t[:4]
    return "已填 %s…（%d 字符）" % (head, len(t))


def tyc_logged_in():
    if not TYC_CFG.exists():
        return False, "尚未登录天眼查账号"
    try:
        d = json.loads(TYC_CFG.read_text(encoding="utf-8"))
        ok = bool((d.get("oauth") or {}).get("refreshToken")
                  or (d.get("headers") or {}).get("Authorization"))
        if not ok:
            return False, "凭证文件存在但没有有效登录信息"
        o = d.get("oauth") or {}
        up = o.get("updatedAt")
        when = ""
        if up:
            try:
                when = datetime.fromtimestamp(int(up) / 1000).strftime("%m-%d %H:%M")
            except Exception:
                when = ""
        detail = "凭证已配置"
        if when:
            detail += "，最近刷新 " + when
        return True, detail
    except Exception as e:
        return False, "凭证文件读取失败：%s" % e


def preflight():
    checks = []

    def add(name, ok, detail, fix=""):
        checks.append({"name": name, "ok": bool(ok), "detail": detail, "fix": fix})

    src = "包内便携运行时" if BUNDLED else "本机已装"
    try:
        ver = subprocess.run(["tyc", "--version"], capture_output=True, shell=True,
                             timeout=30, env=tyc_env())
        v = ((ver.stdout or b"") + (ver.stderr or b"")).decode("utf-8", "ignore").strip()
        add("天眼查工具", bool(v), ("%s（%s）" % (v.splitlines()[-1], src)) if v else "未找到 tyc",
            "请确认 runtime\\node 已完整解压" if BUNDLED else "需要 Node 与 tyc-cli")
    except Exception as e:
        add("天眼查工具", False, "无法调用 tyc：%s" % e, "确认 runtime\\node 目录完整")

    ok, detail = tyc_logged_in()
    add("天眼查登录", ok, detail,
        "点「开始授权」，用浏览器打开链接并输入 6 位验证码，再点「我已授权，继续」")

    qcc = read_auth()
    add("企查查凭证（可选）", True,
        "已填写" if qcc else "未填写：出表需要此项（或已有缓存）",
        "" if qcc else "在左侧「企查查」里粘贴形如 Bearer … 的 token 并保存")

    pku = read_pku()
    add("法宝凭证（可选）", True,
        "已填写：可做判赔全量核查" if pku else "未填写：将跳过判赔核查，不影响出表",
        "" if pku else "要核对判赔金额时，在左侧「法宝」里粘贴 Authorization 并保存")

    try:
        import openpyxl
        add("openpyxl", True, openpyxl.__version__)
    except Exception:
        add("openpyxl", False, "未安装", "请重新解压完整绿色包")

    return checks


# ---------- 任务状态 ----------
LOCK = threading.Lock()
JOB = {
    "phase": "idle",          # idle / running / done / error
    "log": [],
    "out": None,
    "error": "",
    "list_path": None,
    "list_name": "",
    "names": [],
    "started_at": None,
    "finished_at": None,
    "stop": False,
    "proc": None,
}


def log_line(text):
    if not text:
        return
    if not text.endswith("\n"):
        text = text + "\n"
    with LOCK:
        JOB["log"].append(text)


def preview_list(path):
    path = Path(path)
    KEYS = ["企业名称", "公司名称", "权利人", "企业", "公司", "客户名称", "名称"]
    if path.suffix.lower() == ".csv":
        import csv
        with path.open(encoding="utf-8-sig", newline="") as f:
            rows = list(csv.reader(f))
    else:
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        rows = list(wb.worksheets[0].iter_rows(values_only=True))
    hdr_i = col = None
    for i, r in enumerate(rows[:8]):
        for j, c in enumerate(r or []):
            if c and str(c).strip() in KEYS:
                hdr_i, col = i, j
                break
        if col is not None:
            break
    if col is None:
        hdr_i, col = 0, 2 if rows and len(rows[0] or []) > 2 else (0 if not rows else 0)
        if rows and len(rows[0] or []) > 2:
            col = 2
        elif rows and len(rows[0] or []) > 1:
            col = 1
        else:
            col = 0
    names, seen = [], set()
    for r in rows[hdr_i + 1:]:
        if r and col < len(r) and r[col]:
            v = str(r[col]).strip()
            if v and v not in KEYS and v not in seen:
                seen.add(v)
                names.append(v)
    return names


def safe_filename(name):
    name = os.path.basename(name or "名单.xlsx")
    name = re.sub(r'[\\/:*?"<>|]', "_", name)
    if not name.lower().endswith((".xlsx", ".xls", ".csv")):
        name += ".xlsx"
    return name or "名单.xlsx"


def run_step(cmd, env, title):
    log_line("\n===== %s =====" % title)
    p = subprocess.Popen(cmd, cwd=str(BASE), env=env,
                         stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    with LOCK:
        JOB["proc"] = p
    for raw in iter(p.stdout.readline, b""):
        log_line(raw.decode("utf-8", "replace"))
    code = p.wait()
    with LOCK:
        JOB["proc"] = None
    if code != 0:
        raise RuntimeError("%s 失败，退出码 %d" % (title, code))


def work(lst, out, do_rel, do_verify):
    cfg = load_cfg()
    run_cfg = cfg.get("run") or {}
    floor = run_cfg.get("floor", 0)
    group_threshold = run_cfg.get("group_threshold", 1000)

    env = dict(os.environ)
    if BUNDLED:
        env["PATH"] = str(NODE_DIR) + os.pathsep + env.get("PATH", "")
    env["PYTHONIOENCODING"] = "utf-8"
    env["TYC_RUN_DIR"] = str(CACHE)
    env["QCC_RUN_DIR"] = str(CACHE)
    env["SCREEN_OVERRIDE"] = str(CFG_DIR / "判赔覆盖.json")
    auth = read_auth()
    if auth:
        env["QCC_AUTH"] = auth
    pku = read_pku()
    if pku:
        env["PKULAW_AUTH"] = pku

    py = python_exe()
    OUT.mkdir(parents=True, exist_ok=True)
    CACHE.mkdir(parents=True, exist_ok=True)
    tmp = OUT / ("_主表_%s.xlsx" % datetime.now().strftime("%H%M%S"))

    with LOCK:
        JOB["phase"] = "running"
        JOB["error"] = ""
        JOB["out"] = None
        JOB["started_at"] = datetime.now().strftime("%H:%M:%S")
        JOB["finished_at"] = None
        JOB["stop"] = False

    try:
        run_step([py, str(CORE / "画像取数.py"), "--list", str(lst), "--out", str(tmp)],
                 env, "第一步：出主表")
        if do_rel:
            run_step([py, str(CORE / "补查关联主体.py"),
                      "--book", str(tmp), "--out", str(out),
                      "--floor", str(floor),
                      "--group-threshold", str(group_threshold)],
                     env, "第二步：补查关联主体")
        else:
            shutil.copy2(tmp, out)
            log_line("已跳过补查，主表直接输出\n")
        if do_verify and read_pku():
            run_step([py, str(CORE / "判赔核查.py"), "--book", str(out)],
                     env, "第三步：判赔全量核查（法宝）")
        elif do_verify:
            log_line("已勾选判赔核查，但无法宝凭证，已跳过。可在左侧填写后再跑。\n")
        try:
            tmp.unlink()
        except OSError:
            pass
        with LOCK:
            JOB["out"] = str(out)
            JOB["phase"] = "done"
            JOB["finished_at"] = datetime.now().strftime("%H:%M:%S")
        log_line("\n完成：%s\n" % out)
    except Exception as e:
        with LOCK:
            JOB["phase"] = "error"
            JOB["error"] = str(e)
            JOB["finished_at"] = datetime.now().strftime("%H:%M:%S")
        log_line("\n出错：%s\n" % e)
    finally:
        with LOCK:
            JOB["proc"] = None


# ---------- HTTP ----------
class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _send(self, code, body, ctype="application/json; charset=utf-8", extra=None):
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def _json(self, obj, code=200):
        self._send(code, json.dumps(obj, ensure_ascii=False))

    def _read_json(self):
        n = int(self.headers.get("Content-Length") or 0)
        try:
            return json.loads(self.rfile.read(n) or b"{}")
        except Exception:
            return {}

    def do_GET(self):
        p = unquote(self.path.split("?")[0])
        if p in ("/", "/index.html"):
            return self._send(200, UI_FILE.read_bytes(), "text/html; charset=utf-8")
        if p == "/api/preflight":
            return self._json({"checks": preflight()})
        if p == "/api/status":
            with LOCK:
                s = {
                    "phase": JOB["phase"],
                    "log": "".join(JOB["log"]),
                    "error": JOB["error"],
                    "has_output": bool(JOB["out"] and Path(JOB["out"]).exists()),
                    "out_name": Path(JOB["out"]).name if JOB["out"] else "",
                    "list_name": JOB["list_name"],
                    "names": JOB["names"][:30],
                    "name_count": len(JOB["names"]),
                    "started_at": JOB["started_at"],
                    "finished_at": JOB["finished_at"],
                }
            return self._json(s)
        if p == "/api/download":
            out = JOB.get("out")
            if not out or not Path(out).exists():
                return self._json({"error": "还没有可下载的结果"}, 404)
            data = Path(out).read_bytes()
            fn = Path(out).name
            return self._send(200, data,
                              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                              {"Content-Disposition":
                               "attachment; filename=result.xlsx; filename*=UTF-8''%s" % quote(fn)})
        if p == "/api/cred":
            qt, pt = read_auth(), read_pku()
            return self._json({"qcc": mask_token(qt), "has_qcc": bool(qt),
                               "pku": mask_token(pt), "has_pku": bool(pt)})
        return self._json({"error": "not found"}, 404)

    def do_POST(self):
        payload = self._read_json()
        p = unquote(self.path.split("?")[0])

        if p == "/api/upload":
            try:
                raw = base64.b64decode(payload.get("data") or "")
                if not raw:
                    return self._json({"error": "空文件"}, 400)
                LIST.mkdir(parents=True, exist_ok=True)
                fn = safe_filename(payload.get("name") or "名单.xlsx")
                dest = LIST / fn
                dest.write_bytes(raw)
                names = preview_list(dest)
                with LOCK:
                    JOB["list_path"] = str(dest)
                    JOB["list_name"] = fn
                    JOB["names"] = names
                return self._json({"name": fn, "path": str(dest),
                                   "count": len(names), "preview": names[:30]})
            except Exception as e:
                return self._json({"error": "名单解析失败：%s" % e}, 400)

        if p == "/api/start":
            with LOCK:
                if JOB["phase"] == "running":
                    return self._json({"error": "任务正在运行"}, 409)
                lst = payload.get("path") or JOB["list_path"]
            if not lst or not Path(lst).exists():
                return self._json({"error": "请先上传名单文件（Excel 或 CSV）"}, 400)
            auth = read_auth()
            has_cache = CACHE.exists() and any(CACHE.rglob("*.json"))
            if not auth and not has_cache:
                return self._json({"error": "还没配置企查查凭证，缓存也是空的。请先在左侧保存 token。"}, 400)
            ok, _ = tyc_logged_in()
            if not ok:
                return self._json({"error": "请先完成天眼查授权（左侧「开始授权」）。"}, 400)
            OUT.mkdir(parents=True, exist_ok=True)
            out_name = (payload.get("out_name") or "").strip()
            if out_name:
                out_name = re.sub(r'[\\/:*?"<>|]', "_", out_name)
                if not out_name.lower().endswith(".xlsx"):
                    out_name += ".xlsx"
                out = OUT / out_name
            else:
                out = OUT / ("客户涉诉画像_%s.xlsx" % datetime.now().strftime("%Y%m%d_%H%M"))
            do_rel = bool(payload.get("rel", True))
            do_verify = bool(payload.get("verify", True))
            with LOCK:
                JOB["log"] = []
                JOB["out"] = None
                JOB["error"] = ""
            threading.Thread(target=work, args=(Path(lst), out, do_rel, do_verify),
                             daemon=True).start()
            return self._json({"started": True, "out": str(out)})

        if p == "/api/cred":
            try:
                saved = {}
                if "qcc" in payload:
                    save_secret(AUTH_FILE, payload.get("qcc") or "")
                    saved["has_qcc"] = bool(read_auth())
                if "pku" in payload:
                    save_secret(PKU_FILE, payload.get("pku") or "")
                    saved["has_pku"] = bool(read_pku())
                saved["saved"] = True
                return self._json(saved)
            except Exception as e:
                return self._json({"error": "保存失败：%s" % e}, 500)

        if p == "/api/login/start":
            try:
                r = subprocess.run(["tyc", "login", "--no-open", "--no-block"],
                                   capture_output=True, shell=True, timeout=120,
                                   cwd=str(BASE), env=tyc_env())
            except Exception as e:
                return self._json({"ok": False, "raw": "启动授权失败：%s" % e})
            out = ((r.stdout or b"") + (r.stderr or b"")).decode("utf-8", "ignore")
            m_url = re.search(r"(https?://[^\s\u4e00-\u9fff]+)", out)
            m_code = re.search(r"(\d{6})", out)
            return self._json({"ok": bool(m_url), "url": m_url.group(1) if m_url else "",
                               "code": m_code.group(1) if m_code else "", "raw": out[:800]})

        if p == "/api/login/resume":
            try:
                r = subprocess.run(["tyc", "login", "--resume"],
                                   capture_output=True, shell=True, timeout=180,
                                   cwd=str(BASE), env=tyc_env())
            except Exception as e:
                return self._json({"ok": False, "raw": "确认授权失败：%s" % e})
            out = ((r.stdout or b"") + (r.stderr or b"")).decode("utf-8", "ignore")
            ok = ("登录成功" in out) or ("authorization_pending" not in out and "成功" in out)
            return self._json({"ok": ok, "raw": out[-800:]})

        return self._json({"error": "not found"}, 404)


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    CACHE.mkdir(parents=True, exist_ok=True)
    LIST.mkdir(parents=True, exist_ok=True)
    if not UI_FILE.exists():
        print("缺少 ui.html，请确认压缩包解压完整。")
        sys.exit(1)
    srv = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    url = "http://127.0.0.1:%d" % PORT
    print("涉诉画像工作台已启动：%s" % url)
    print("关掉这个窗口即停止工作台。")
    if not os.environ.get("HANA_WB_NOOPEN"):
        try:
            webbrowser.open(url)
        except Exception:
            pass
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        print("已停止。")


if __name__ == "__main__":
    main()
