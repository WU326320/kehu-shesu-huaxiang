# -*- coding: utf-8 -*-
"""涉诉画像工作台：选名单、点开始、出表。"""
import json
import os
import queue
import re
import subprocess
import sys
import threading
import webbrowser
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent
RUNTIME = BASE / "runtime"
NODE_DIR = RUNTIME / "node"
PY_DIR = Path(sys.executable).resolve().parent
_tcl, _tk = PY_DIR / "tcl" / "tcl8.6", PY_DIR / "tcl" / "tk8.6"
if _tcl.exists():
    os.environ.setdefault("TCL_LIBRARY", str(_tcl))
if _tk.exists():
    os.environ.setdefault("TK_LIBRARY", str(_tk))
if (NODE_DIR / "tyc.cmd").exists():
    os.environ["PATH"] = str(NODE_DIR) + os.pathsep + os.environ.get("PATH", "")

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

CORE = BASE / "core"
CACHE = BASE / "缓存"
OUT = BASE / "输出"
LIST = BASE / "名单"
AUTH_FILE = BASE / "凭证.txt"
PKU_FILE = BASE / "法宝凭证.txt"
TYC_CFG = Path.home() / ".tyc" / "config.json"
BUNDLED = (NODE_DIR / "tyc.cmd").exists()


def load_cfg():
    p = BASE / "配置" / "config.toml"
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


def tyc_env():
    e = dict(os.environ)
    if BUNDLED:
        e["PATH"] = str(NODE_DIR) + os.pathsep + e.get("PATH", "")
    return e


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
        detail = "已登录"
        if when:
            detail += "，最近刷新 " + when
        return True, detail
    except Exception as e:
        return False, "凭证文件读取失败：%s" % e



def _help_box(parent, lines):
    """可展开的短教程：普通文字 + 可点击的网址。"""
    box = ttk.Frame(parent)
    for item in lines:
        if isinstance(item, tuple) and item[0] == "link":
            text, url = item[1], item[2]
            lbl = tk.Label(box, text=text, fg="#0563C1", cursor="hand2",
                           anchor="w", justify="left")
            lbl.pack(anchor="w", padx=2)
            lbl.bind("<Button-1>", lambda e, u=url: webbrowser.open(u))
        else:
            ttk.Label(box, text=item, wraplength=720, justify="left").pack(
                anchor="w", padx=2, pady=1)
    return box


QCC_HELP = [
    "这是企查查「智能体数据平台」的 API Key，用来查出表所需的裁判文书。",
    "1. 用浏览器打开下面这个网站，用自己的企查查账号登录：",
    ("link", "https://agent.qcc.com/", "https://agent.qcc.com/"),
    "2. 点右上角头像，在菜单里选「获取我的 API KEY」。也可以直接打开密钥页：",
    ("link", "https://agent.qcc.com/profile/api-key", "https://agent.qcc.com/profile/api-key"),
    "3. 复制 API Key，粘贴到上面的输入框。如果复制出来没有「Bearer」，请在前面加上「Bearer 」（后面有一个空格）再保存。",
    "4. 点「保存」。凭证只写在这台电脑的「凭证.txt」里，不要发给别人或发到群里。",
]

PKU_HELP = [
    "这是北大法宝 MCP 平台的 Access Token，用来逐笔核对判赔金额。可不填。",
    "1. 用浏览器打开下面这个网站，注册或登录：",
    ("link", "https://mcp.pkulaw.com/", "https://mcp.pkulaw.com/"),
    "2. 按页面提示进入控制台，获取 Access Token（官网步骤是「注册账号，控制台获取 Token」）。",
    "3. 把 Token 粘贴到上面的输入框。格式是「Bearer 」加 Token，中间一个空格。如果复制出来已经带 Bearer，直接粘贴即可。",
    "4. 点「保存」。凭证只写在这台电脑的「法宝凭证.txt」里。不填也能出表，只是不做判赔核查。",
]


class App:
    def __init__(self, root):
        self.root = root
        self.cfg = load_cfg()
        self.q = queue.Queue()
        self.proc = None
        self.running = False
        self._tyc_code = ""

        root.title("涉诉画像工作台 · 2026-10-19")
        root.geometry("940x720")
        root.minsize(760, 520)

        frm = ttk.Frame(root, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="名单文件（Excel 或 CSV）").grid(row=0, column=0, sticky="w")
        self.var_list = tk.StringVar()
        ttk.Entry(frm, textvariable=self.var_list, width=70).grid(row=0, column=1, padx=6, sticky="ew")
        ttk.Button(frm, text="浏览", command=self.pick_list).grid(row=0, column=2)

        ttk.Label(frm, text="输出文件").grid(row=1, column=0, sticky="w", pady=6)
        self.var_out = tk.StringVar(value=str(OUT / ("客户涉诉画像_%s.xlsx" % datetime.now().strftime("%Y%m%d"))))
        ttk.Entry(frm, textvariable=self.var_out, width=70).grid(row=1, column=1, padx=6, sticky="ew")
        ttk.Button(frm, text="另存为", command=self.pick_out).grid(row=1, column=2)

        ttk.Label(frm, text="天眼查").grid(row=2, column=0, sticky="w")
        tycf = ttk.Frame(frm)
        tycf.grid(row=2, column=1, columnspan=2, sticky="w", pady=4)
        self.var_tyc = tk.StringVar(value="检查中…")
        ttk.Label(tycf, textvariable=self.var_tyc).pack(side="left")
        ttk.Button(tycf, text="开始授权", command=self.tyc_start).pack(side="left", padx=8)
        ttk.Button(tycf, text="我已授权，继续", command=self.tyc_resume).pack(side="left")

        ttk.Label(frm, text="企查查").grid(row=3, column=0, sticky="nw", pady=(6, 0))
        self.var_qcc = tk.StringVar()
        ttk.Entry(frm, textvariable=self.var_qcc, width=70, show="*").grid(row=3, column=1, padx=6, sticky="ew", pady=(4, 0))
        qcc_btns = ttk.Frame(frm)
        qcc_btns.grid(row=3, column=2, pady=(4, 0), sticky="e")
        ttk.Button(qcc_btns, text="保存", command=lambda: self.save_secret("qcc")).pack(side="left")
        self.btn_qcc_help = ttk.Button(qcc_btns, text="说明 ▸", width=8, command=self.toggle_qcc_help)
        self.btn_qcc_help.pack(side="left", padx=(6, 0))
        self.var_qcc_st = tk.StringVar(value="粘贴接口凭证，一般以 Bearer 开头。出表需要填写。不清楚从哪复制，点右侧「说明」。")
        ttk.Label(frm, textvariable=self.var_qcc_st, foreground="#555").grid(row=4, column=1, columnspan=2, sticky="w", padx=6, pady=(2, 4))
        self.frm_qcc_help = _help_box(frm, QCC_HELP)
        self._qcc_help_open = False

        ttk.Label(frm, text="法宝").grid(row=6, column=0, sticky="nw", pady=(6, 0))
        self.var_pku = tk.StringVar()
        ttk.Entry(frm, textvariable=self.var_pku, width=70, show="*").grid(row=6, column=1, padx=6, sticky="ew", pady=(4, 0))
        pku_btns = ttk.Frame(frm)
        pku_btns.grid(row=6, column=2, pady=(4, 0), sticky="e")
        ttk.Button(pku_btns, text="保存", command=lambda: self.save_secret("pku")).pack(side="left")
        self.btn_pku_help = ttk.Button(pku_btns, text="说明 ▸", width=8, command=self.toggle_pku_help)
        self.btn_pku_help.pack(side="left", padx=(6, 0))
        self.var_pku_st = tk.StringVar(value="粘贴 Authorization。可不填，不填也能出表，只是不做判赔核查。不清楚从哪复制，点右侧「说明」。")
        ttk.Label(frm, textvariable=self.var_pku_st, foreground="#555").grid(row=7, column=1, columnspan=2, sticky="w", padx=6, pady=(2, 4))
        self.frm_pku_help = _help_box(frm, PKU_HELP)
        self._pku_help_open = False

        chk = ttk.Frame(frm)
        chk.grid(row=9, column=1, sticky="w", pady=6)
        self.var_rel = tk.BooleanVar(value=True)
        ttk.Checkbutton(chk, text="补查关联主体（穿透控制企业，生成附表）",
                        variable=self.var_rel).pack(side="left")
        self.var_verify = tk.BooleanVar(value=True)
        ttk.Checkbutton(chk, text="判赔全量核查（法宝，逐笔）",
                        variable=self.var_verify).pack(side="left", padx=12)

        btns = ttk.Frame(frm)
        btns.grid(row=10, column=1, sticky="w", pady=8)
        self.btn_start = ttk.Button(btns, text="开始", command=self.start)
        self.btn_start.pack(side="left")
        ttk.Button(btns, text="打开输出目录", command=self.open_out).pack(side="left", padx=8)
        ttk.Button(btns, text="检查环境", command=self.check_env).pack(side="left")
        ttk.Button(btns, text="口径", command=self.show_rules).pack(side="left", padx=8)

        self.log = tk.Text(frm, height=22, wrap="word")
        self.log.grid(row=11, column=0, columnspan=3, sticky="nsew", pady=8)
        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(11, weight=1)

        self.refresh_status()
        self.root.after(200, self.poll)


    def _toggle_help(self, kind):
        if kind == "qcc":
            open_now = self._qcc_help_open
            frame, btn, row = self.frm_qcc_help, self.btn_qcc_help, 5
        else:
            open_now = self._pku_help_open
            frame, btn, row = self.frm_pku_help, self.btn_pku_help, 8
        if open_now:
            frame.grid_remove()
            btn.config(text="说明 ▸")
            opened = False
        else:
            frame.grid(row=row, column=1, columnspan=2, sticky="ew", padx=6, pady=(0, 6))
            btn.config(text="收起 ▾")
            opened = True
        if kind == "qcc":
            self._qcc_help_open = opened
        else:
            self._pku_help_open = opened

    def toggle_qcc_help(self):
        self._toggle_help("qcc")

    def toggle_pku_help(self):
        self._toggle_help("pku")

    def refresh_status(self):
        ok, detail = tyc_logged_in()
        self.var_tyc.set(detail if ok else ("未登录（" + detail + "）" if detail != "尚未登录天眼查账号" else "尚未登录"))
        qcc_hint = "粘贴接口凭证，一般以 Bearer 开头。出表需要填写。不清楚从哪复制，点右侧「说明」。"
        pku_hint = "粘贴 Authorization。可不填，不填也能出表，只是不做判赔核查。不清楚从哪复制，点右侧「说明」。"
        self.var_qcc_st.set("已保存到本机「凭证.txt」。要更换请重新粘贴后点保存。" if _read_secret(AUTH_FILE) or os.environ.get("QCC_AUTH", "").strip() else qcc_hint)
        self.var_pku_st.set("已保存到本机「法宝凭证.txt」。要更换请重新粘贴后点保存。" if _read_secret(PKU_FILE) or os.environ.get("PKULAW_AUTH", "").strip() else pku_hint)

    def save_secret(self, kind):
        path = AUTH_FILE if kind == "qcc" else PKU_FILE
        value = (self.var_qcc.get() if kind == "qcc" else self.var_pku.get()).strip()
        try:
            if not value:
                if path.exists():
                    path.unlink()
            else:
                path.write_text(value + "\n", encoding="utf-8")
            if kind == "qcc":
                self.var_qcc.set("")
            else:
                self.var_pku.set("")
            self.refresh_status()
        except OSError as e:
            messagebox.showerror("保存失败", str(e))

    def tyc_start(self):
        self.write("正在获取天眼查授权链接…\n")
        self.root.update_idletasks()
        try:
            r = subprocess.run(["tyc", "login", "--no-open", "--no-block"],
                               capture_output=True, shell=True, timeout=60,
                               cwd=str(BASE), env=tyc_env())
        except Exception as e:
            messagebox.showerror("授权失败", "获取授权链接失败：%s" % e)
            return
        out = ((r.stdout or b"") + (r.stderr or b"")).decode("utf-8", "ignore")
        m_url = re.search(r"https?://\S+", out)
        m_code = re.search(r"(\d{6})", out)
        url = m_url.group(0).rstrip(").,，。\"'") if m_url else ""
        code = m_code.group(1) if m_code else ""
        self._tyc_code = code
        if url:
            try:
                webbrowser.open(url)
            except Exception:
                pass
        msg = []
        if url:
            msg.append("已在浏览器打开授权页：\n%s" % url)
        else:
            msg.append("没有解析到授权链接。")
        if code:
            msg.append("请输入验证码：%s" % code)
        msg.append("在页面里确认授权后，回到这里点「我已授权，继续」。")
        if not url:
            msg.append("\n命令输出：\n%s" % (out[-800:] or "(空)"))
        self.write(("验证码：%s\n" % code) if code else "")
        if url:
            self.write("授权页：%s\n" % url)
        messagebox.showinfo("天眼查授权", "\n\n".join(msg))

    def tyc_resume(self):
        try:
            r = subprocess.run(["tyc", "login", "--resume"],
                               capture_output=True, shell=True, timeout=180,
                               cwd=str(BASE), env=tyc_env())
        except Exception as e:
            messagebox.showerror("确认失败", "确认授权失败：%s" % e)
            return
        out = ((r.stdout or b"") + (r.stderr or b"")).decode("utf-8", "ignore")
        ok = ("登录成功" in out) or ("authorization_pending" not in out and "成功" in out)
        self.refresh_status()
        if ok or tyc_logged_in()[0]:
            messagebox.showinfo("天眼查授权", "授权成功")
            self.write("天眼查授权成功\n")
        else:
            messagebox.showerror("还没成功", "请确认已在浏览器完成授权后再试。\n\n%s" % out[-800:])

    def pick_list(self):
        p = filedialog.askopenfilename(filetypes=[("表格", "*.xlsx *.xls *.csv"), ("全部", "*.*")])
        if p:
            self.var_list.set(p)

    def pick_out(self):
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if p:
            self.var_out.set(p)

    def open_out(self):
        OUT.mkdir(parents=True, exist_ok=True)
        os.startfile(OUT)  # noqa: S606

    def write(self, text):
        self.log.insert("end", text)
        self.log.see("end")

    def check_env(self):
        msgs = []
        msgs.append("Python: %s" % sys.version.split()[0])
        try:
            import openpyxl
            msgs.append("openpyxl: %s" % openpyxl.__version__)
        except Exception:
            msgs.append("openpyxl: 未安装（请重新解压完整绿色包）")
        import shutil as sh
        tyc = sh.which("tyc", path=tyc_env().get("PATH"))
        msgs.append("tyc: %s" % (tyc or "未找到（绿色包需完整解压 runtime\\node）"))
        ok, detail = tyc_logged_in()
        msgs.append("天眼查登录: %s" % (detail if ok else "未登录 — 请点「开始授权」"))
        msgs.append("企查查凭证: %s" % ("已填写" if self.read_auth() else "未填写（出表需要，或已有缓存）"))
        msgs.append("法宝凭证: %s" % ("已填写" if self.read_pku() else "未填写（核查会跳过）"))
        msgs.append("缓存: %d 个文件" % (len(list(CACHE.rglob("*.json"))) if CACHE.exists() else 0))
        messagebox.showinfo("环境检查", "\n".join(msgs))

    def show_rules(self):
        txt = (
            "时间：裁判日期 2024-01-01 至 2026-10-31\n"
            "案件：专利 / 商标 / 著作权 / 不正当竞争 / 其他知产\n"
            "执行类：案由含「执行」且不含「之诉」的，不计件不计判赔\n"
            "成件：按案号；同案的「之一」「之二」算一件，不同案号各算一件\n"
            "二审：正文点到原审案号且当事人有交集才并入\n"
            "判赔：主文赔偿项之和；连带范围不另计；裁定书的案件金额不用\n"
            "等级：A>100 / B 51-100 / C 21-50 / D 6-20，基数=自身+关联\n"
            "法宝：逐笔全量核查；查不到不改变金额，只标注来源"
        )
        messagebox.showinfo("口径（2026-10-19）", txt)

    def read_auth(self):
        return _read_secret(AUTH_FILE) or os.environ.get("QCC_AUTH", "").strip()

    def read_pku(self):
        return _read_secret(PKU_FILE) or os.environ.get("PKULAW_AUTH", "").strip()

    def start(self):
        if self.running:
            return
        lst = Path(self.var_list.get().strip())
        if not lst.exists():
            messagebox.showerror("缺少名单", "请先选择名单文件")
            return
        auth = self.read_auth()
        has_cache = CACHE.exists() and any(CACHE.rglob("*.json"))
        if not auth and not has_cache:
            messagebox.showerror("还没配置", "还没有企查查凭证，缓存也是空的。\n请先在「企查查」栏粘贴 token 并保存。")
            return
        ok, _ = tyc_logged_in()
        if not ok:
            messagebox.showerror("还没授权", "请先完成天眼查授权：点「开始授权」，在浏览器输入验证码，再点「我已授权，继续」。")
            return

        out = Path(self.var_out.get().strip())
        if not out.name:
            messagebox.showerror("输出文件", "请填写输出文件路径")
            return
        if out.suffix.lower() != ".xlsx":
            out = out.with_suffix(".xlsx")
            self.var_out.set(str(out))
        out.parent.mkdir(parents=True, exist_ok=True)
        self.running = True
        self.btn_start.config(state="disabled")
        self.log.delete("1.0", "end")
        threading.Thread(target=self.work, args=(lst, out), daemon=True).start()

    def work(self, lst, out):
        env = tyc_env()
        env["PYTHONIOENCODING"] = "utf-8"
        env["TYC_RUN_DIR"] = str(CACHE)
        env["QCC_RUN_DIR"] = str(CACHE)
        env["SCREEN_OVERRIDE"] = str(BASE / "配置" / "判赔覆盖.json")
        auth = self.read_auth()
        if auth:
            env["QCC_AUTH"] = auth
        pku = self.read_pku()
        if pku:
            env["PKULAW_AUTH"] = pku

        tmp = OUT / ("_主表_%s.xlsx" % datetime.now().strftime("%H%M%S"))
        try:
            OUT.mkdir(parents=True, exist_ok=True)
            CACHE.mkdir(parents=True, exist_ok=True)
            self.run_step([sys.executable, str(CORE / "画像取数.py"),
                           "--list", str(lst), "--out", str(tmp)], env, "第一步：出主表")
            if self.var_rel.get():
                self.run_step([sys.executable, str(CORE / "补查关联主体.py"),
                               "--book", str(tmp), "--out", str(out),
                               "--floor", str(self.cfg["run"]["floor"]),
                               "--group-threshold", str(self.cfg["run"]["group_threshold"])],
                              env, "第二步：补查关联主体")
            else:
                shutil_copy(tmp, out)
                self.q.put(("log", "已跳过补查，主表直接输出\n"))
            if self.var_verify.get() and self.read_pku():
                self.run_step([sys.executable, str(CORE / "判赔核查.py"),
                               "--book", str(out)], env, "第三步：判赔全量核查（法宝）")
            elif self.var_verify.get():
                self.q.put(("log", "已勾选判赔核查，但无法宝凭证，已跳过。可在「法宝」栏填写后再跑。\n"))
            try:
                tmp.unlink()
            except OSError:
                pass
            self.q.put(("done", str(out)))
        except Exception as e:
            self.q.put(("error", str(e)))
        finally:
            self.q.put(("stop", ""))

    def run_step(self, cmd, env, title):
        self.q.put(("log", "\n===== %s =====\n" % title))
        p = subprocess.Popen(cmd, cwd=str(BASE), env=env,
                             stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        self.proc = p
        for raw in iter(p.stdout.readline, b""):
            line = raw.decode("utf-8", "replace")
            self.q.put(("log", line))
        code = p.wait()
        if code != 0:
            raise RuntimeError("%s 失败，退出码 %d" % (title, code))

    def poll(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "log":
                    self.write(payload)
                elif kind == "done":
                    self.write("\n完成：%s\n" % payload)
                    messagebox.showinfo("完成", "已生成：\n%s" % payload)
                elif kind == "error":
                    self.write("\n出错：%s\n" % payload)
                    messagebox.showerror("出错", payload)
                elif kind == "stop":
                    self.running = False
                    self.btn_start.config(state="normal")
        except queue.Empty:
            pass
        self.root.after(200, self.poll)


def shutil_copy(a, b):
    import shutil
    shutil.copy2(a, b)


def main():
    root = tk.Tk()
    try:
        root.iconbitmap(default="")
    except Exception:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        import traceback
        err = traceback.format_exc()
        try:
            (BASE / "启动失败.txt").write_text(err, encoding="utf-8")
        except Exception:
            pass
        try:
            r = tk.Tk()
            r.withdraw()
            messagebox.showerror("启动失败", err)
        except Exception:
            pass
        raise
