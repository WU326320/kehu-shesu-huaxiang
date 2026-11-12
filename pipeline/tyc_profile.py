#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
天眼查取数内核（tyc CLI 封装）

本模块只提供底层调用能力，**不是可执行脚本**。被 `tools\画像取数.py` 通过
`load_core()` 动态加载，对外暴露：run_tyc / log_call / fetch_registration /
fetch_ip / NO_CALLS。

职责边界：
- 负责调用 tyc CLI、落盘原始返回体、记台账、处理额度与假空重试
- 不负责案件口径（筛时间、筛案由、归并、判赔）——在 tools\画像取数.py
- 不负责成表——在 tools\画像取数.py 与 tools\补查关联主体.py

关键约束：
- 列表接口每页硬上限 20 条，pageSize 传 100 无效
- 「假空」（_empty 但措辞笃定）**只对 tag 以 ip 开头的查询重试**；其余接口的
  _empty 是真空（未上榜、未上市），重试纯属浪费调用
- 原始返回体落盘到 RUN_DIR/raw，配合 TYC_REUSE=1 可零花费重算

历史：本文件原属 2026-10-10 的 9 列工作台流程；旧流程（简介 / 联系人 / 主要IP /
类型 / 侵权情况 / 天眼查取案件）已于 2026-10-18 整体移除。
"""

import csv
import hashlib
import json
import os
import subprocess
import sys
from datetime import datetime
from pathlib import Path

# tyc 落盘对中文路径不友好，原始体与台账一律走 ASCII 路径；
# 默认存在本包内的 state 目录，可用环境变量 TYC_RUN_DIR 覆盖
RUN_DIR = Path(os.environ.get("TYC_RUN_DIR") or (Path(__file__).resolve().parent.parent / "state"))
RAW_DIR = RUN_DIR / "raw"
LEDGER = RUN_DIR / "ledger.csv"

REUSE = os.environ.get("TYC_REUSE") == "1"   # True 时优先读缓存，缺缓存才调用（--derive-only 也置 True）
NO_CALLS = 0
QUOTA_HIT = False       # 额度耗尽的硬信号（HTTP 402）


# ---------- tyc 调用 ----------
def log_call(company, cmd, ok, note=""):
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    new = not LEDGER.exists()
    with LEDGER.open("a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["时间", "企业", "命令", "成功", "备注"])
        w.writerow([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), company, cmd, "Y" if ok else "N", note])


def parse_json_loose(s: str):
    if not s:
        return None
    i, j = s.find("{"), s.rfind("}")
    if i >= 0 and j > i:
        try:
            return json.loads(s[i:j + 1])
        except Exception:
            return None
    return None


def run_tyc(args, company, tag, retries=1, empty_retry=None):
    """执行一次 tyc 调用。优先读 --output-file，失败退回 stdout。REUSE 时只读缓存。

    empty_retry：命中「假空」（_empty）时的重试次数。假空实测只出现在知产接口，
    其他接口的 _empty 是真空（如未上榜、未上市），重试纯属浪费调用，故默认只对
    tag 以 ip 开头的查询重试。
    """
    global NO_CALLS, QUOTA_HIT
    RAW_DIR.mkdir(parents=True, exist_ok=True)
    h = hashlib.md5(company.encode("utf-8")).hexdigest()[:10]
    out_file = RAW_DIR / f"{h}__{tag}.json"

    if REUSE and out_file.exists():
        try:
            return True, json.loads(out_file.read_text(encoding="utf-8-sig")), ""
        except Exception:
            pass

    if out_file.exists():
        out_file.unlink()
    cmd = ["tyc"] + args + ["--output-file", str(out_file)]
    stdout, stderr = "", ""
    # 实测：「已全量扫描…未发现任何记录」会偶发假空（同一主体几分钟后又能查到）。
    # 空结果不计费，重试是免费的，所以假空必重试一次。
    if empty_retry is None:
        empty_retry = 2 if tag.startswith("ip") else 0
    for _ in range(retries + 2):
        try:
            p = subprocess.run(cmd, capture_output=True, shell=True, timeout=180)
        except subprocess.TimeoutExpired:
            continue
        NO_CALLS += 1
        stderr = p.stderr.decode("utf-8", "ignore").strip()
        stdout = p.stdout.decode("utf-8", "ignore").strip()
        blob = stdout + "\n" + stderr
        if ("quota_exceeded" in blob or "HTTP 402" in blob
                or "次数已用完" in blob or "额度已用完" in blob or "配额" in blob):
            QUOTA_HIT = True
            log_call(company, " ".join(args), False, "quota_exceeded")
            return False, None, "quota_exceeded: " + blob[:300]
        data = None
        if out_file.exists():
            try:
                data = json.loads(out_file.read_text(encoding="utf-8-sig"))
            except Exception:
                data = None
        if data is None:
            data = parse_json_loose(stdout)
        if data is not None:
            if isinstance(data, dict) and data.get("_empty") and empty_retry > 0:
                empty_retry -= 1
                log_call(company, " ".join(args), False, "假空重试")
                continue
            ded = 0 if (isinstance(data, dict) and data.get("_empty")) else 1
            log_call(company, " ".join(args), True,
                     f"{'file' if out_file.exists() else 'stdout'}|扣{ded}")
            return True, data, ""
    log_call(company, " ".join(args), False,
             "异常: " + (stderr or stdout)[:150].replace("\n", " "))
    return False, None, (stderr or stdout)[:300] or "empty"


def handle_quota(err):
    if err and "quota_exceeded" in err:
        print("\n!! 额度耗尽，立即停止。报错体：")
        print(err)
        sys.exit(2)


# ---------- 各维度取数 ----------
def fetch_registration(name):
    ok, d, err = run_tyc(["company", "registration-info", name], name, "registration")
    handle_quota(err)
    if not ok:
        return None
    b = (d.get("sources") or {}).get("base") or {}
    return {
        "全名": b.get("name") or name,
        "统一社会信用代码": b.get("creditCode"),
        "法定代表人": b.get("legalPersonName"),
        "成立日期": b.get("estiblishTime"),
        "登记状态": b.get("regStatus"),
        "行业": (b.get("industryAll") or {}).get("categoryBig") or b.get("industry") or (b.get("industryAll") or {}).get("categoryMiddle"),
        "经营范围": b.get("businessScope"),
        "官网": b.get("websiteList"),
        "省": b.get("city"),
    }


def fetch_ip(name, page_size=20):
    """知识产权四项的 total。

    page_size 默认 20（旧流程取“主要IP”需要明细）；新流程只要 total，传 1 即可：
    实测 pageSize=20 时“已全量扫描未发现记录”的假空率异常高，pageSize=1 未复现。
    """
    out = {}
    for tag, args in [
        ("商标", ["intellectual_property", "trademark-info", name, "--pageNum", "1", "--pageSize", str(page_size)]),
        ("专利", ["intellectual_property", "patent-info", name, "--pageNum", "1", "--pageSize", str(page_size)]),
        ("软件著作权", ["intellectual_property", "software-copyright-info", name, "--pageNum", "1", "--pageSize", str(page_size)]),
        ("作品著作权", ["intellectual_property", "copyright-work-info", name, "--pageNum", "1", "--pageSize", str(page_size)]),
    ]:
        ok, d, err = run_tyc(args, name, "ip" + tag)
        handle_quota(err)
        out[tag] = d.get("total") if (ok and isinstance(d, dict)) else None
        out[tag + "_items"] = (d.get("items") if (ok and isinstance(d, dict)) else None) or []
    return out


if __name__ == "__main__":
    main()
