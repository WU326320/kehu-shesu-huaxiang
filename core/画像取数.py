#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""客户涉诉画像 · 新流程取数

    等级 | 企业名称 | 知识产权 | 案件情况

天眼查：工商登记（取现名）+ 商标/专利/软著/作品著作权
企查查：司法文书（逐案明细，近两年知产，含原被告与判赔）

用法：
    python 画像取数.py --list "名单.xlsx" --out "输出.xlsx" [--limit N] [--sheet Sheet1]
    python 画像取数.py --company "某某有限公司"        # 单家试跑，不写表
"""

import argparse
import importlib.util
import json
import os
import re
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

BASE = Path(__file__).resolve().parent
PKG = BASE.parent
os.environ.setdefault("TYC_RUN_DIR", str(PKG / "缓存"))
sys.path.insert(0, str(BASE))

from qcc_client import (documents as qcc_documents, truncation as qcc_truncation,
                        credit_error)   # noqa: E402

SINCE, UNTIL = "2024-01-01", "2026-10-31"

# 经判决原文核对后的判赔覆盖表：{案号: 金额}。企查查解析错、但法宝原文可查的，
# 写在这里，重跑时优先采用，避免人工更正在下一次跑表时被冲掉。
OVERRIDE_FILE = Path(os.environ.get("SCREEN_OVERRIDE")
                    or str(PKG / "配置" / "判赔覆盖.json"))
OVERRIDE = {}
if OVERRIDE_FILE.exists():
    try:
        OVERRIDE = {str(k): float(v) for k, v in
                    json.loads(OVERRIDE_FILE.read_text(encoding="utf-8")).items()}
    except Exception:
        OVERRIDE = {}

# 类目按优先级从高到低，命中即止。
# 「知识产权权属、侵权纠纷」是企查查的泛化案由，不含任何具体类别词，
# 必须靠最后的「知识产权」兜住；否则整类案子会被静默丢弃（实测 10 家缓存里有 16 条）。
IP_RULES = [("专利", ["专利", "发明", "实用新型", "外观设计"]),
            ("商标", ["商标"]),
            ("著作权", ["著作权", "作品", "信息网络传播", "复制权", "发行权", "软件"]),
            ("不正当竞争", ["不正当竞争", "不正当竞", "仿冒", "混淆", "商业秘密", "擅自使用"]),
            ("其他知产", ["知识产权", "特许经营", "技术合同", "垄断", "植物新品种",
                        "集成电路布图", "商业诋毁"])]

AMT_ANY = re.compile(r"([0-9][0-9,]*(?:\.[0-9]+)?)\s*(万元|元)")


def load_core():
    for p in (BASE.parent / "workbench" / "pipeline" / "tyc_profile.py",
              BASE.parent / "pipeline" / "tyc_profile.py"):
        if p.exists():
            break
    spec = importlib.util.spec_from_file_location("tyc_core", p)
    m = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(m)
    return m


def ip_kind(reason):
    r = reason or ""
    # 执行类文书不是新的纠纷，是既有判决的履行程序，计入件数会重复。
    # 实测：上海吼美 13 条「…案件执行」被归到「其他知产/不正当竞争」，
    # 其中 2 条进了件数、1 条把「案件金额」当成判赔（6100）。
    # 只有「执行异议之诉」这类带「之诉」的才是独立诉讼，保留。
    if "执行" in r and "之诉" not in r:
        return None
    for name, kws in IP_RULES:
        if any(k in r for k in kws):
            return name
    return None


def sides(doc):
    """归一当事人：甲侧（原告/上诉人）、乙侧（被告/被上诉人）

    企查查有约 2.6% 的文书把双方混在泛化键「当事人」里（实测 29/1112），
    这类文书若已有被告侧，则把泛化列表里不在被告侧的人归到原告侧，
    否则明细会显示成「— 诉 唐*」，丢掉原告。泛化键为空或被告侧也空时不猜。
    """
    p = doc.get("当事人") or {}
    a, b, generic = [], [], []
    for role, names in p.items():
        rr = str(role)
        for n in (names or []):
            if "被上诉" in rr:
                b.append(n)
            elif "上诉" in rr or "原告" in rr:
                a.append(n)
            elif "被告" in rr:
                b.append(n)
            elif rr.strip() == "当事人":
                generic.append(n)
    if not a and generic and b:
        a = [n for n in generic if n not in b]
    return a, b


def to_yuan(num, unit):
    v = float(str(num).replace(",", ""))
    return v * 10000 if unit == "万元" else v


def damage(doc):
    """判赔金额：判决主文中「赔偿/赔付」项之和（剔除诉讼费）。返回 (元 或 None, 来源)

    规则（2026-10-19 修订）：
    1. 遍历**全部**含赔偿/赔付/给付的判项并求和，不能拿到第一笔就返回
       （实测（2024）粤1973民初2699号：三名被告分别赔偿 9000+150000+4000，
       旧逻辑只取到 9000）。
    2. 「…X元及Y元，两项共计Z元」只取 Z，分项与总计同时加会翻倍。
    3. 「在对X元范围内承担连带责任」「其中X元」是责任范围，不另计。
    4. 同一判项分列「经济损失 X 元和维权合理费用 Y 元」仍要相加。
    案件受理费/诉讼费不是赔给原告的，不计入。
    """
    txt = str(doc.get("裁判结果") or "")
    total, hit = 0.0, False
    for seg in re.split(r"[；\n。]", txt):
        if not ("赔偿" in seg or "赔付" in seg or "给付" in seg):
            continue
        # 「在对X元范围内承担连带责任」「其中X元」是责任范围，不是另加一笔
        skip = set()
        for mm in re.finditer(r"在([0-9][0-9,]*(?:\.[0-9]+)?)元范围内", seg):
            skip.add(mm.start(1))
        for mm in re.finditer(r"其中([0-9][0-9,]*(?:\.[0-9]+)?)元", seg):
            skip.add(mm.start(1))
        vals = []
        for m in AMT_ANY.finditer(seg):
            pre = seg[max(0, m.start() - 24):m.start()]
            # 只有费用词紧贴金额（中间没有数字）时才认定这笔是费用本身。
            # 旧写法「24 字内出现受理费就跳过」会把「案件受理费500元，被告赔偿原告1000元」
            # 里的 1000 一并剔掉，实测已复现。
            if re.search(r"(受理费|诉讼费|保全费|鉴定费|执行费)[^0-9]{0,8}$", pre):
                continue
            if m.start(1) in skip:
                continue
            vals.append(to_yuan(m.group(1), m.group(2)))
        if not vals:
            continue
        hit = True
        # 「…X元及Y元，两项共计Z元」只取 Z；把分项与总计一起加会翻倍
        # （实测（2023）京0107民初10391号：2068600 被算成 4237200）
        if "共计" in seg or "合计" in seg:
            k = "共计" if "共计" in seg else "合计"
            after = seg.split(k, 1)[1]
            m = AMT_ANY.search(after)
            if m:
                total += to_yuan(m.group(1), m.group(2))
                continue
        total += sum(vals)
    if hit and total:
        return total, "文本"
    # 「案件金额」对裁定书来说往往是保全标的额（冻结/查封金额），不是判赔。
    # 实测（2023）沪0114民初27791号之二：裁定书案件金额 250,000 会被当成判赔。
    # 因此只对判决书、调解书启用字段兜底。
    title = str(doc.get("文书标题") or "")
    v = str(doc.get("案件金额") or "").strip()
    if v and ("判决" in title or "调解" in title):
        try:
            return float(v.replace(",", "")), "字段回退"
        except ValueError:
            pass
    # 不再做「取正文第一个金额」的兜底：诉讼请求额常出现在判决书正文里，
    # 会把诉请当判赔。解析不出就回 None，宁缺不错。
    return None, None


def fmt_money(v):
    if v is None:
        return ""
    return "{:,.0f}".format(v)


ZHI = re.compile(r"^(.*号)之[一二三四五六七八九十]+$")
CASE_REF = re.compile(r"（\d{4}）[^\s，、]{0,20}?民初\d+号")


def base_code(code):
    """（2024）沪0104民初12909号之一 → （2024）沪0104民初12909号

    法院编号里「之一/之二」是同一案件里的另一份文书（多为裁定），不是新案。
    """
    s = str(code or "").strip()
    m = ZHI.match(s)
    return m.group(1) if m else s


def name_set(doc):
    """当事人名字集合，用于判断二审引用的原审是不是同一批人。"""
    out = set()
    for names in (doc.get("当事人") or {}).values():
        for n in (names or []):
            out.add(str(n).split("（")[0].strip())
    return out


def merge_key(doc):
    """归并键：基础案号优先。

    2026-10-19 定：**按案号成件，不按当事人集合成件**。当事人被匿名化成
    「刘**」，不同的人会同名，按名字并组会把不同案件缝成一件（实测吼美 6 处、
    米哈游 1 处，少计 12 件并漏两笔判赔）。反方向，「之一」是同一案件的另一份
    文书，不并又会多计。两者都要靠案号判断，不能靠人名。

    二审不并到一审，除非它自己的正文点到在池的一审案号（见 build_cases）。
    案号缺失时退回「类别 + 当事人集合」。
    """
    code = base_code(doc.get("案号"))
    if code:
        return ("案号", code)
    a, b = sides(doc)
    names = frozenset(str(n).split("（")[0] for n in a + b)
    return ("当事人", ip_kind(doc.get("案由")), names)


def build_cases(docs):
    """筛时间 → 筛知产 → 归并 → 排序 → 文本

    规则：解析不出当事人的行，只在有判赔时保留；两边都空的直接跳过。
    """
    kept = []
    for d in docs:
        day = str(d.get("裁判日期") or "")[:10]     # 带时分秒的日期会让字符串比较出错
        if not (SINCE <= day <= UNTIL):
            continue
        if not ip_kind(d.get("案由")):
            continue
        kept.append(d)
    # 先把在池的一审案号收集起来，供二审正文引用匹配
    first_codes = {}
    for d in kept:
        c = base_code(d.get("案号"))
        if c and "民终" not in c:
            first_codes[c] = d

    groups = {}
    for d in kept:
        code = base_code(d.get("案号"))
        key = None
        if code and "民终" in code:
            txt = str(d.get("裁判结果") or "") + str(d.get("文书标题") or "")
            mine = name_set(d)
            for ref in CASE_REF.findall(txt):
                rc = base_code(ref)
                # 正文里引用的案号未必是本案原审（可能只是援引在先判决），
                # 要求当事人有交集才算同一纠纷。
                if rc in first_codes and (mine & name_set(first_codes[rc])):
                    key = ("案号", rc)
                    break
        if key is None:
            key = merge_key(d)
        groups.setdefault(key, []).append(d)

    rows, skipped = [], 0
    for g in groups.values():
        rep = max(g, key=lambda x: str(x.get("裁判日期") or ""))
        codes = "、".join(sorted({str(x.get("案号")) for x in g if x.get("案号")})) or "案号待核"
        a, b = sides(rep)
        # 判赔取组内**有金额**的最新一份：之一/之二是同案裁定，按日期会盖过判决书
        with_amt = [(x, damage(x)) for x in g]
        with_amt = [(x, r) for x, r in with_amt if r[0] is not None]
        if with_amt:
            dmg_doc, (amt, src) = max(with_amt, key=lambda t: str(t[0].get("裁判日期") or ""))
        else:
            amt, src = damage(rep)
        # 人工更正：优先取代表文书案号上的覆盖值；没有则按案号排序取第一条，
        # 保证结果可复现（同一组里有多个覆盖值时，不依赖列表顺序）。
        ovs = [(str(x.get("案号") or ""), OVERRIDE[str(x.get("案号"))])
               for x in g if OVERRIDE.get(str(x.get("案号"))) is not None]
        if ovs:
            rep_code = str(rep.get("案号") or "")
            amt = next((v for c, v in ovs if c == rep_code), sorted(ovs)[0][1])
            src = "覆盖表"
        if not a and not b and amt is None:
            skipped += 1          # 没当事人也没金额，这一行没有信息
            continue
        line = "%s %s 诉 %s %s" % (codes,
                                   "、".join(a) or "—",
                                   "、".join(b) or "—",
                                   "判赔 %s 元" % fmt_money(amt) if amt else "未判赔")
        rows.append((amt if amt is not None else -1, line))
    rows.sort(key=lambda x: -x[0])
    return [r[1] for r in rows], len(groups), len(kept), skipped


def grade(n):
    if n > 100:
        return "A"
    if n > 50:
        return "B"
    if n > 20:
        return "C"
    if n > 5:
        return "D"
    return ""


def run_one(core, name, use_cache=True):
    reg = core.fetch_registration(name)
    real = (reg or {}).get("全名") or name
    ip = core.fetch_ip(real, page_size=1)      # 只要 total，pageSize 越小越稳
    if not ip or any(v is None for v in ip.values()):
        print("     !! 知识产权存量未取全：%s（单元格写「未返回」）" % real, file=sys.stderr)
    docs, total, qcalls, notes = qcc_documents(real)
    lines, n_merged, n_kept, n_skip = build_cases(docs)
    ip_text = "\n".join("%s：%s 件" % (k, ip.get(k) if ip.get(k) is not None else "未返回")
                        for k in ["商标", "专利", "软件著作权", "作品著作权"])
    if lines:
        case_text = "近两年知产案件 %d 件\n" % n_merged + "\n".join(lines)
    else:
        case_text = "近两年无知产案件"
    trunc = qcc_truncation(real)
    if trunc:
        # 接口每次只回最近 30 条；取不全时必须在表内标明，件数只是下限
        case_text += "\n（企查查接口上限：%s，件数为下限，需人工下载剩余裁判文书）" % "、".join(
            "%d 年累计 %d 条仅取 %d 条" % t for t in trunc)
    return {"等级": grade(n_merged), "企业名称": real, "知识产权": ip_text, "案件情况": case_text,
            "_meta": {"输入名": name, "登记文书总数": total, "本次文书": len(docs),
                      "近两年知产记录": n_kept, "归并件数": n_merged, "无信息跳过": n_skip,
                      "企查查调用": qcalls, "备注": notes}}


HEADERS = ["等级", "企业名称", "知识产权", "案件情况"]


def write_book(rows, out):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "涉诉画像"
    ws.append(HEADERS)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="D9D9D9")
        c.alignment = Alignment(horizontal="center", vertical="center")
    for r in rows:
        ws.append([r.get(h) for h in HEADERS])
    for col, w in zip("ABCD", [6, 30, 24, 90]):
        ws.column_dimensions[col].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out


def read_list(path, sheet=None):
    path = Path(path)
    if path.suffix.lower() == ".csv":
        import csv
        with path.open(encoding="utf-8-sig", newline="") as f:
            rows = list(csv.reader(f))
    else:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb[sheet] if sheet else wb.worksheets[0]
        rows = list(ws.iter_rows(values_only=True))
    hdr_i = col = None
    KEYS = ["企业名称", "公司名称", "权利人", "企业", "公司", "客户名称", "名称"]
    for i, r in enumerate(rows[:8]):
        for j, c in enumerate(r or []):
            if c and str(c).strip() in KEYS:
                hdr_i, col = i, j
                break
        if col is not None:
            break
    if col is None:
        hdr_i, col = 0, 2 if len(rows[0]) > 2 else 1
    names, seen = [], set()
    for r in rows[hdr_i + 1:]:
        if col < len(r or []) and r[col]:
            v = str(r[col]).strip()
            if v and v not in KEYS and v not in seen:
                seen.add(v)
                names.append(v)
    return names


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--list")
    ap.add_argument("--sheet")
    ap.add_argument("--out")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--company")
    a = ap.parse_args()
    core = load_core()

    if a.company:
        r = run_one(core, a.company)
        meta = r.pop("_meta")
        print(json.dumps(r, ensure_ascii=False, indent=1))
        print("---", json.dumps(meta, ensure_ascii=False))
        print("天眼查调用:", core.NO_CALLS, "| 企查查调用:", meta["企查查调用"])
        return

    if not a.list:
        ap.error("需要 --list 或 --company")
    names = read_list(Path(a.list), a.sheet)
    if a.limit:
        names = names[:a.limit]
    print("待跑 %d 家" % len(names), file=sys.stderr)

    rows = []
    for i, n in enumerate(names, 1):
        print("[%d/%d] %s" % (i, len(names), n), file=sys.stderr)
        try:
            r = run_one(core, n)
            if any(credit_error(x) for x in (r["_meta"].get("备注") or [])):
                # 余额耗尽：不写这家（否则会写成「近两年无知产案件」），整轮停
                print("     !! 企查查余额耗尽，整轮停止。已完成 %d 家，未写入：%s"
                      % (len(rows), n), file=sys.stderr)
                break
            rows.append(r)
            print("     等级=%s 归并=%s 件 | 天眼查累计=%d 企查查=%d"
                  % (r["等级"] or "（空）", r["_meta"]["归并件数"],
                     core.NO_CALLS, r["_meta"]["企查查调用"]), file=sys.stderr)
        except Exception as e:
            print("     !! 失败：%s" % e, file=sys.stderr)
    out = Path(a.out) if a.out else (PKG / "输出" / "客户涉诉画像.xlsx")
    print("已写出:", write_book(rows, out))
    print("天眼查调用 %d 次" % core.NO_CALLS)


if __name__ == "__main__":
    main()
