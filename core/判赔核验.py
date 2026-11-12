#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""判赔核验：按案号取法宝判决主文，与企查查解析值对账。

为什么按案号而不是按当事人：法宝把当事人匿名化（北京光线影业 -> 北京某某影业），
按公司名查命中率极低；按案号查返回 Total=1 且 CaseFlag 精确匹配。

用法：
    python 判赔核验.py                          # 跑内置 13 件
    python 判赔核验.py --cases 清单.json --out 结果.csv [--quiet]
清单格式：[{"parent":..., "sub":..., "rep":"（2024）...号", "amt":63800}, ...]
"""
import argparse
import csv
import io
import json
import os
import re
import sys
import time
import urllib.request
from pathlib import Path

PKG = Path(__file__).resolve().parent.parent
PKU_FILE = PKG / "法宝凭证.txt"
URL = "https://apim-gateway.pkulaw.com/mcp-case"
WINDOW = {"startLastInstanceDate": "2024.1.1", "endLastInstanceDate": "2026.10.31"}

CASES = [
    ("北京光线传媒", "北京光线影业", "（2024）沪0105民初36770号", 63800),
    ("北京光线传媒", "北京光线影业", "（2024）鲁1728民初1226号", 25000),
    ("北京光线传媒", "北京光线影业", "（2025）赣1104民初1113号", 20000),
    ("北京光线传媒", "北京光线影业", "（2023）沪0118民初35450号", 10000),
    ("上海鹰角网络科技", "上海木鸢网络科技", "（2025）渝0192民初766号", 20000),
    ("上海鹰角网络科技", "上海木鸢网络科技", "（2025）粤0704民初5878号", 20000),
    ("上海鹰角网络科技", "上海木鸢网络科技", "（2025）粤0704民初5879号", 8000),
    ("上海鹰角网络科技", "上海木鸢网络科技", "（2024）沪0105民初30625号", 8000),
    ("上海鹰角网络科技", "上海木鸢网络科技", "（2024）沪0105民初37198号", 3500),
    ("上海鹰角网络科技", "上海木鸢网络科技", "（2024）沪0105民初37203号", 3000),
    ("上海鹰角网络科技", "上海木鸢网络科技", "（2024）沪0105民初19710号", 3000),
    ("上海鹰角网络科技", "上海木鸢网络科技", "（2024）沪0105民初19868号", 1500),
    ("上海哔哩哔哩科技", "上海萌电文化科技", "（2024）内0102知民初362号", 30000),
]

AMT = re.compile(r"([0-9][0-9,]*(?:\.[0-9]+)?)\s*(万元|元)")
CN_AMT = re.compile(r"([零一二三四五六七八九十百千万亿两]+)\s*(万元|元)")
CN_D = {"零": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4, "五": 5,
        "六": 6, "七": 7, "八": 8, "九": 9}
CN_U = {"十": 10, "百": 100, "千": 1000}


def cn2num(s):
    """中文数字转阿拉伯数字，支持到亿。无法解析返回 None。"""
    total, section, num = 0, 0, 0
    for ch in s:
        if ch in CN_D:
            num = CN_D[ch]
        elif ch in CN_U:
            section += (num or 1) * CN_U[ch]
            num = 0
        elif ch == "万":
            total += (section + num) * 10000
            section, num = 0, 0
        elif ch == "亿":
            total = (total + section + num) * 100000000
            section, num = 0, 0
        else:
            return None
    return total + section + num


def auth():
    """凭证优先：环境变量 PKULAW_AUTH，然后包内 法宝凭证.txt。"""
    v = os.environ.get("PKULAW_AUTH")
    if v and v.strip():
        return v.strip()
    if PKU_FILE.exists():
        for line in PKU_FILE.read_text(encoding="utf-8-sig").splitlines():
            t = line.strip()
            if t and not t.startswith("#"):
                return t
    raise RuntimeError("找不到法宝凭证：请设置环境变量 PKULAW_AUTH，或在包内 法宝凭证.txt 填写")


def rpc(method, params, sid=None, _id=1):
    body = {"jsonrpc": "2.0", "id": _id, "method": method}
    if params is not None:
        body["params"] = params
    req = urllib.request.Request(URL, data=json.dumps(body).encode(), method="POST")
    req.add_header("Authorization", auth())
    req.add_header("Content-Type", "application/json")
    req.add_header("Accept", "application/json, text/event-stream")
    if sid:
        req.add_header("Mcp-Session-Id", sid)
    with urllib.request.urlopen(req, timeout=90) as r:
        raw = r.read().decode("utf-8", "ignore")
        newsid = r.headers.get("Mcp-Session-Id") or sid
    if method == "notifications/initialized":
        return None, newsid
    for line in raw.splitlines():
        if line.startswith("data:"):
            try:
                return json.loads(line[5:].strip()), newsid
            except Exception:
                pass
    try:
        return json.loads(raw), newsid
    except Exception:
        return {"_raw": raw[:400]}, newsid


def norm(s):
    return re.sub(r"[\s（）()]", "", str(s or ""))


def fetch(case_no):
    sid = None
    _, sid = rpc("initialize", {"protocolVersion": "2024-11-05", "capabilities": {},
                                "clientInfo": {"name": "hana-verify", "version": "1"}})
    rpc("notifications/initialized", None, sid)
    args = {"fulltext": case_no}
    args.update(WINDOW)
    res, _ = rpc("tools/call", {"name": "get_case_list", "arguments": args}, sid, 2)
    try:
        payload = json.loads("\n".join(p.get("text", "") for p in res["result"]["content"]))
    except Exception:
        return {"_err": json.dumps(res, ensure_ascii=False)[:300]}
    rows = payload.get("Data") or []
    hit = [x for x in rows if norm(x.get("CaseFlag")) == norm(case_no)]
    return hit[0] if hit else {"_miss": True, "n": len(rows)}


def award(result):
    """从判决主文提取赔偿总额。返回 (金额 or None, 摘录)。

    规则：只取含「赔偿/赔付/给付」的判项，剔受理费等；有「共计」取共计值，
    否则把该判项内的金额相加。阿拉伯数字优先，中文数字兜底。
    """
    t = str(result or "").replace("\t", "")
    segs = [s.strip() for s in re.split(r"[；;\n]", t)]
    picked, total = [], 0.0
    for seg in segs:
        if not re.search(r"赔偿|赔付|给付", seg):
            continue
        if re.search(r"受理费|诉讼费|保全费|执行费|鉴定费", seg):
            continue
        # 「在 X 元范围内承担连带责任」「其中 X 元」是责任范围，不是另加一笔。
        # 实测（2024）沪0105民初22613号：40,000 赔偿 + 30,000 连带范围，
        # 旧逻辑相加得 70,000，与判决主文不符。
        skip = set()
        for mm in re.finditer(r"在([0-9][0-9,]*(?:\.[0-9]+)?)元范围内", seg):
            skip.add(mm.start(1))
        for mm in re.finditer(r"其中([0-9][0-9,]*(?:\.[0-9]+)?)元", seg):
            skip.add(mm.start(1))
        vals = []
        for m in AMT.finditer(seg):
            if m.start(1) in skip:
                continue
            v = float(m.group(1).replace(",", ""))
            vals.append(v * 10000 if m.group(2) == "万元" else v)
        chinese = False
        if not vals:
            for m in CN_AMT.finditer(seg):
                v = cn2num(m.group(1))
                if v:
                    vals.append(v * 10000 if m.group(2) == "万元" else v)
                    chinese = True
        if not vals:
            continue
        picked.append(seg[:200])
        key = "共计" if "共计" in seg else ("合计" if "合计" in seg else "")
        if key:
            after = seg.split(key, 1)[1]
            # 「共计」后面可能是阿拉伯数字，也可能是中文数字。按命中哪个正则来解析，
            # 不能按前面的分项写法猜，否则「赔偿10000元，共计三万元」会 float("三万") 抛错。
            m_ar = AMT.search(after)
            m_cn = None if m_ar else CN_AMT.search(after)
            if m_ar:
                v = float(m_ar.group(1).replace(",", ""))
                total += v * 10000 if m_ar.group(2) == "万元" else v
                continue
            if m_cn:
                v = cn2num(m_cn.group(1))
                if v:
                    total += v * 10000 if m_cn.group(2) == "万元" else v
                    continue
        total += sum(vals)
    return (total if picked else None), "；".join(picked)[:300]


def load_cases(path):
    rows = json.load(io.open(path, encoding="utf-8"))
    return [(x.get("parent", ""), x.get("sub", ""), x.get("rep") or (x.get("codes") or [""])[0],
             x.get("amt")) for x in rows]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--cases", default="")
    ap.add_argument("--out", default="")
    ap.add_argument("--quiet", action="store_true")
    a = ap.parse_args()
    cases = load_cases(a.cases) if a.cases else CASES
    rows, ok, bad, miss = [], 0, 0, 0
    for i, (grp, sub, no, qcc) in enumerate(cases, 1):
        d = fetch(no)
        if "_err" in d or "_miss" in d:
            miss += 1
            rows.append([grp, sub, no, qcc, "", "未取到", "", "", ""])
            if not a.quiet:
                print("[%d/%d] %s 未取到" % (i, len(cases), no), file=sys.stderr)
            time.sleep(0.4)
            continue
        got, seg = award(d.get("RefereeResult"))
        if got is None:
            verdict = "未解析"
            miss += 1
        elif qcc is not None and abs(got - float(qcc)) < 1:
            verdict = "一致"
            ok += 1
        else:
            verdict = "不一致"
            bad += 1
        rows.append([grp, sub, no, qcc, ("" if got is None else ("%.0f" % got)),
                     verdict, seg, d.get("Court"), d.get("LastInstanceDate")])
        if not a.quiet:
            print("[%d/%d] %s %s %s vs %s" % (i, len(cases), no, verdict, qcc, got),
                  file=sys.stderr)
        time.sleep(0.4)

    print("\n合计 %d｜一致 %d｜不一致 %d｜未取到/未解析 %d" % (len(rows), ok, bad, miss))
    for r in rows:
        if r[5] in ("不一致", "未解析"):
            print("  %s %s：企查查 %s，法宝 %s｜%s" % (r[2], r[5], r[3], r[4] or "-", r[6][:120]))
    if a.out:
        with io.open(a.out, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(["主体组", "关联主体", "案号", "企查查判赔", "法宝判赔", "判定",
                        "法宝主文摘录", "法院", "审结日期"])
            w.writerows(rows)
        print("已写出:", a.out)


if __name__ == "__main__":
    main()
