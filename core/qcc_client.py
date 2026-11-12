#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""企查查 MCP 客户端（streamable-http，JSON-RPC over POST）

凭证优先：环境变量 QCC_AUTH，然后包内 凭证.txt。
仅当环境变量 QCC_MCP_CFG 已设置时才读 MCP 配置。
所有返回体缓存到 缓存/qcc_raw/，重复调用不再走网络。
"""

import csv
import json
import os
import re
import sys
import urllib.request
from datetime import datetime
from pathlib import Path

PKG = Path(__file__).resolve().parent.parent
RUN = Path(os.environ.get("QCC_RUN_DIR") or str(PKG / "缓存"))
CACHE = RUN / "qcc_raw"
QCC_LEDGER = RUN / "qcc_ledger.csv"          # 每次调用记一行（时间/主体/工具/扣分）
QCC_BASELINE = RUN / "qcc_baseline.json"      # 用户提供的余额基准
COST_PER_CALL = 3                            # 实测：get_judicial_documents 每次 3 分
DOMAIN = "risk"
BASE = "https://agent.qcc.com/mcp/%s/stream" % DOMAIN


def set_baseline(points, note=""):
    """按控制台看到的余额设基准。剩余 = 基准 − 之后累计扣分。"""
    RUN.mkdir(parents=True, exist_ok=True)
    QCC_BASELINE.write_text(json.dumps(
        {"points": float(points), "at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "note": note},
        ensure_ascii=False, indent=1), encoding="utf-8")


def spent_since_baseline():
    if not QCC_BASELINE.exists() or not QCC_LEDGER.exists():
        return 0.0
    base = json.loads(QCC_BASELINE.read_text(encoding="utf-8"))
    t0 = base.get("at", "")
    total = 0.0
    with QCC_LEDGER.open(encoding="utf-8-sig", newline="") as f:
        for r in csv.DictReader(f):
            if (r.get("时间") or "") >= t0:
                try:
                    total += float(r.get("扣分") or 0)
                except ValueError:
                    pass
    return total


def remaining():
    """本地估算的剩余积分；无基准时返回 None。"""
    if not QCC_BASELINE.exists():
        return None
    base = json.loads(QCC_BASELINE.read_text(encoding="utf-8"))
    return float(base.get("points") or 0) - spent_since_baseline()


def _log(company, tool, cost):
    RUN.mkdir(parents=True, exist_ok=True)
    new = not QCC_LEDGER.exists()
    with QCC_LEDGER.open("a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["时间", "主体", "工具", "扣分"])
        w.writerow([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), company, tool, cost])


AUTH_FILE = PKG / "凭证.txt"


def _auth():
    """企查查凭证。优先：环境变量 QCC_AUTH，然后包内 凭证.txt。
    仅当环境变量 QCC_MCP_CFG 已设置时才读 MCP 配置。
    """
    v = os.environ.get("QCC_AUTH")
    if v and v.strip():
        return v.strip()
    if AUTH_FILE.exists():
        for line in AUTH_FILE.read_text(encoding="utf-8-sig").splitlines():
            tok = line.strip()
            if tok and not tok.startswith("#"):
                return tok
    cfg_path = (os.environ.get("QCC_MCP_CFG") or "").strip()
    if cfg_path:
        cfg = json.loads(Path(cfg_path).read_text(encoding="utf-8"))
        for c in cfg["global"]["mcp"]["connectors"]:
            if c.get("id") == "qcc-company":
                auth = (c.get("headers") or {}).get("Authorization", "")
                if auth:
                    return auth
    raise RuntimeError("找不到企查查凭证：请设置环境变量 QCC_AUTH，或在包内 凭证.txt 填写")


def safe_name(name):
    """缓存文件名的清洗规则，与 judicial_documents 保持一致。"""
    return re.sub(r"[^\w\u4e00-\u9fff-]", "_", str(name))[:44]


def credit_error(text):
    """判断错误是不是「余额/积分不足」。是则应整轮停下，别再逐家撞墙。"""
    t = str(text or "")
    return any(k in t for k in ("300008", "积分不足", "余额不足", "欠费", "quota", "insufficient"))


def has_cache(company):
    """该主体是否已有可用缓存（命中缓存不扣分，余额检查应当放行）。"""
    for f in CACHE.glob(safe_name(company) + "__*.json"):
        try:
            if "_err" not in json.loads(f.read_text(encoding="utf-8")):
                return True
        except Exception:
            continue
    return False


def pool_from_cache(company):
    """把某主体所有缓存切片拼成 {案号: 文书}，同案号取信息更全的一份。

    只读本地缓存，不联网、不计费。缺缓存时返回空字典。
    """
    pool = {}
    for f in CACHE.glob(safe_name(company) + "__*.json"):
        try:
            d = json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            continue
        if "_err" in d:
            continue
        for x in d.get("裁判文书") or []:
            k = x.get("案号") or x.get("文书ID")
            if k:
                pool[k] = richer(pool.get(k), x)
    return pool


def richer(a, b):
    """同一案号可能出现在多个切片里，内容未必相同。

    实测（2024）沪0116民初14618号在缓存里有两份：判决书（裁判结果 153 字，判赔 50 万）
    与「本案转为普通程序」的裁定书（9 字，无金额）。按后到覆盖会拿到裁定，判赔丢失。
    规则改为取裁判结果更长的一份；等长时保留先到的一份，保证可复现。
    """
    if a is None:
        return b
    ka = (len(str(a.get("裁判结果") or "")), len(str(a.get("当事人") or "")))
    kb = (len(str(b.get("裁判结果") or "")), len(str(b.get("当事人") or "")))
    return a if ka >= kb else b


class Qcc:
    def __init__(self, base=BASE):
        self.base = base
        self.sid = None
        self.calls = 0

    def _rpc(self, method, params=None, _id=1):
        body = {"jsonrpc": "2.0", "id": _id, "method": method}
        if params is not None:
            body["params"] = params
        req = urllib.request.Request(self.base, data=json.dumps(body).encode(), method="POST")
        req.add_header("Authorization", _auth())
        req.add_header("Content-Type", "application/json")
        req.add_header("Accept", "application/json, text/event-stream")
        if self.sid:
            req.add_header("Mcp-Session-Id", self.sid)
        with urllib.request.urlopen(req, timeout=90) as r:
            raw = r.read().decode("utf-8", "ignore")
            sid = r.headers.get("Mcp-Session-Id")
            if sid:
                self.sid = sid
        for line in raw.splitlines():
            if line.startswith("data:"):
                try:
                    return json.loads(line[5:].strip())
                except Exception:
                    pass
        try:
            return json.loads(raw)
        except Exception:
            return {"_raw": raw[:400]}

    def handshake(self):
        self._rpc("initialize", {"protocolVersion": "2024-11-05", "capabilities": {},
                                 "clientInfo": {"name": "hana", "version": "1"}})
        self._rpc("notifications/initialized", None)

    def call(self, tool, args):
        self.calls += 1
        res = self._rpc("tools/call", {"name": tool, "arguments": args}, 100 + self.calls)
        try:
            out = json.loads("\n".join(p.get("text", "") for p in res["result"]["content"]))
        except Exception:
            return {"_err": json.dumps(res, ensure_ascii=False)[:300]}
        # 台账写失败不能连累已经成功的调用：2026-10-19 因缓存目录不可写，
        # PermissionError 被上面的 except 捕获，把取回的数据当成 _err 丢了，白烧 3 分。
        try:
            _log(str(args.get("searchKey") or ""), tool, COST_PER_CALL)
        except Exception as e:
            print("[warn] 台账写入失败，本次扣分未记账：%s" % e, file=sys.stderr)
        return out

    # ---------- 司法文书 ----------
    def judicial_documents(self, company, year=None, use_cache=True):
        CACHE.mkdir(parents=True, exist_ok=True)
        tag = "y%d" % year if year else "all"
        f = CACHE / ("%s__%s.json" % (re.sub(r"[^\w\u4e00-\u9fff-]", "_", company)[:44], tag))
        if use_cache and f.exists():
            try:
                d = json.loads(f.read_text(encoding="utf-8"))
                if "_err" not in d:
                    return d
            except Exception:
                pass
        args = {"searchKey": company}
        if year:
            args["year"] = year
        d = self.call("get_judicial_documents", args)
        f.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")
        return d


def truncation(company):
    """检查缓存里 2024-2026 各年份是否被接口上限截断。

    企查查对高量主体每次只返回最近 30 条，摘要里的「累计共有 N 条」大于实收条数即为缺。
    返回 [(年, 累计, 实收)]，空列表表示取全。
    """
    out = []
    safe = re.sub(r"[^\w\u4e00-\u9fff-]", "_", company)[:44]
    for y in (2024, 2025, 2026):
        f = CACHE / ("%s__y%d.json" % (safe, y))
        if not f.exists():
            continue
        try:
            d = json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            continue
        n = len(d.get("裁判文书") or [])
        m = re.search(r"累计共有(\d+)条", str(d.get("摘要") or ""))
        if m and int(m.group(1)) > n:
            out.append((y, int(m.group(1)), n))
    return out


def _cache_file(company, year=None):
    tag = "y%d" % year if year else "all"
    return CACHE / ("%s__%s.json" % (re.sub(r"[^\w\u4e00-\u9fff-]", "_", company)[:44], tag))


def _read_cache(company, year=None):
    """只读缓存，不建连接。离线复现时零凭证也能跑。"""
    f = _cache_file(company, year)
    if not f.exists():
        return None
    try:
        d = json.loads(f.read_text(encoding="utf-8"))
    except Exception:
        return None
    return None if "_err" in d else d


def documents(company, years=(2024, 2025)):
    """取该企业的司法文书。

    先发一枪默认调用。默认调用是「实时检索 <窗口起点> 至今，封顶 30 条」，
    窗口起点每家不同。只有摘要明说「包含 30 条」或本页返回满 30 条，才认为被截断，
    才按年补调；否则一枪即全（含 0 条的空结果，也要一枪了事）。

    补调年份只给 (2024, 2025)：因「发布日期 ≥ 裁判日期」，裁判日期落在 2024-2026
    的案子其发布年份必也在 2024-2026，而默认窗口已覆盖 2026，故补 2026 多余。
    """
    holder = {"q": None}

    def q():
        if holder["q"] is None:
            holder["q"] = Qcc()
            holder["q"].handshake()
        return holder["q"]

    calls = 0
    pool, total, notes = {}, None, []

    d = _read_cache(company)
    if d is None:
        d = q().judicial_documents(company, None)
        calls += 1
    if "_err" in d:
        notes.append("默认调用失败:%s" % str(d.get("_err"))[:120])
        return [], None, calls, notes
    s = str(d.get("摘要") or "")
    m = re.search(r"累计共有(\d+)条", s)
    if m:
        total = int(m.group(1))
    got = d.get("裁判文书") or []
    for x in got:
        k = x.get("案号") or x.get("文书ID")
        pool[k] = richer(pool.get(k), x)
    # 截断与否只能看摘要/条数：摘要明写「包含 30 条」，或本页返回满 30 条。
    # ⚠ 旧判据 complete = ("完整呈现" in s) or (total and len(got) >= total)
    #   在「查无记录」（0 条、摘要为空）时恒为 False，会为每家空主体白烧两次补调
    #   （实测 B站 10 个关联主体里 7 家是空的，每家 3 次调用）。
    truncated = ("包含30条" in s.replace(" ", "")) or (len(got) >= 30)
    complete = (not truncated) or (total is not None and len(got) >= total)

    if not complete:
        for y in years:
            d2 = _read_cache(company, y)
            if d2 is None:
                d2 = q().judicial_documents(company, y)
                calls += 1
            if "_err" in d2:
                notes.append("year=%d 失败:%s" % (y, str(d2.get("_err"))[:120]))
                continue
            for x in (d2.get("裁判文书") or []):
                k = x.get("案号") or x.get("文书ID")
                pool[k] = richer(pool.get(k), x)
    return list(pool.values()), total, calls, notes
