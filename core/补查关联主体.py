#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""第二步：网络评估 + 关联主体补查

输入：第一步产出的表（等级 / 企业名称 / 知识产权 / 案件情况）
流程：
    ① 全部公司（不设案件数门槛；2026-10-16 前为「近两年案件数 ≤ 20」）
    ② 网络评估：是否有「企业榜」上榜（只看 type=企业榜，不计品牌/产品榜）
    ③ 取 controlled-companies（向下穿透实际控制企业，读 item.percent）
    ④ 够格 = 有企业榜 或 有持股过半（≥50%）的控制企业；不够格的跳过
    ⑤ 逐个关联主体查企查查司法文书
    ⑥ 只有查出近两年知产案件的关联主体才进附表（没案子的不列）
    ⑦ 输出：主表（原样）+ 附表（关联主体及其案件，按主体组标同底色）

用法：
    python 补查关联主体.py --book "第一步结果.xlsx" --out "补查结果.xlsx"
    python 补查关联主体.py --probe "某某有限公司"      # 单家看返回结构，不写表
"""

import argparse
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

from qcc_client import (  # noqa: E402
    documents as qcc_documents,
    remaining as qcc_remaining,
    set_baseline,
    truncation as qcc_truncation,
    has_cache,
    credit_error,
)
from 画像取数 import build_cases, load_core, grade          # noqa: E402

# 主体组循环底色（浅色，互不混淆）
GROUP_FILLS = ["E8F0F6", "F6EFE6", "E9F2EA", "F3E9F1", "F1F0E2", "E6EEF2"]
MAX_CASES_FOR_EVAL = None   # None = 不设案件数门槛，全部公司都做关联补查
                            # （2026-10-16 前为 20，那个门槛没有论证过，且恰好把米哈游挡在门外）
GROUP_ENTITY_THRESHOLD = 40  # 控制企业数 ≥ 此值视为「集团性质」，本次跳过、末尾集中待确认


def case_count(case_text):
    """从「案件情况」里取**自身**件数。

    优先认汇总行里的「自身 X 件」：等级已改为「自身 + 关联主体」合计，
    重跑时若只看开头那个数字，会把合计当成自身，越算越大。
    老格式（「近两年知产案件 N 件」且无拆解）仍兼容。
    """
    t = str(case_text or "")
    m = re.search(r"自身\s*(\d+)\s*件", t)
    if m:
        return int(m.group(1))
    m = re.search(r"近两年知产案件\s*(\d+)\s*件", t)
    if m:
        return int(m.group(1))
    return 0


def signals(core, real):
    """网络评估：是否有「企业本身行业地位」的榜单。

    只看 type=「企业榜」的条目（排的是企业本身），不计「品牌产品榜」。
    返回 (企业榜条目, 全部条目, 原始体)。
    """
    ok, d, _ = core.run_tyc(["company", "companies-by-ranking", real,
                             "--pageNum", "1", "--pageSize", "20"], real, "ranking")
    if not ok or not isinstance(d, dict):
        # 不要静默返回空——空会与「真的没上榜」混淆，把公司误判为不够格
        raise RuntimeError("companies-by-ranking 调用失败")
    allc, corp = [], []
    if ok and isinstance(d, dict):
        for it in (d.get("items") or []):
            if not isinstance(it, dict):
                continue
            allc.append(it)
            if str(it.get("type") or "").strip() == "企业榜":
                corp.append(it)
    return corp, allc, (d if ok else None)


def controlled(core, real):
    """向下穿透实际控制的企业。返回 [(名称, 持股原文, 持股数值 or None)]

    持股比例直接读 item.percent（顶层字段，即穿透后的实际控制比例），
    不从 chainList 里刨，因为多层链条上各段百分比之和/积不等于综合持股。

    ⚠️ 列表接口每页**硬上限 20 条**，pageSize 传 100 无效，必须翻页。
    实测不翻页时：光线 78 条只拿 20、荃银 49 条只拿 20、华熙 49 条只拿 20。
    """
    best, total, page = {}, None, 1
    while page <= 12:                       # 硬上限 240 家，足够
        ok, d, _ = core.run_tyc(["company", "controlled-companies", real,
                                 "--pageNum", str(page), "--pageSize", "20"],
                                real, "controlled_p%d" % page)
        if not ok or not isinstance(d, dict):
            # 第一页就失败 → 报错；后续页失败 → 保留已拿到的
            # （不要静默返回 []，那与「真的没有控制企业」无法区分）
            if page == 1:
                raise RuntimeError("controlled-companies 调用失败")
            break
        if total is None:
            total = d.get("total")
        items = d.get("items") or []
        for it in items:
            if not isinstance(it, dict):
                continue
            comp = str(it.get("name") or "").strip()
            if not comp or comp == real:
                continue
            pct_raw = str(it.get("percent") or "")
            m = re.search(r"\d+(?:\.\d+)?", pct_raw)
            pct = float(m.group(0)) if m else None
            if comp not in best or (pct is not None and (best[comp][1] is None or pct > best[comp][1])):
                best[comp] = (pct_raw, pct)
        if len(items) < 20 or (total is not None and len(best) >= total):
            break
        page += 1
    # 按持股从高到低排（无比例的置底），余额不够时先覆盖控制力最强的
    return sorted(((c, v[0], v[1]) for c, v in best.items()),
                  key=lambda x: (x[2] is None, -(x[2] or 0)))


def qualified(corp, rels):
    """够格判据：有企业榜上榜 或 有持股过半（≥50%）的控制企业。"""
    over = [r for r in rels if r[2] is not None and r[2] >= 50.0]
    return bool(corp) or bool(over), over


def probe(core, name):
    reg = core.fetch_registration(name)
    real = (reg or {}).get("全名") or name
    print("主体:", real)
    corp, allc, raw1 = signals(core, real)
    print("\n-- companies-by-ranking：共 %d 条，其中「企业榜」%d 条 --" % (len(allc), len(corp)))
    for it in corp[:12]:
        print("    %s | 名次 %s | %s" % (it.get("name"), it.get("precedence"), it.get("source")))
    if not corp:
        others = [(it.get("name"), it.get("type")) for it in allc[:5]]
        print("    （无企业榜；非企业榜条目示例：%s）" % others)
    rels = controlled(core, real)
    over = [r for r in rels if r[2] is not None and r[2] >= 50.0]
    print("\n-- controlled-companies：共 %d 家，其中持股过半 %d 家 --" % (len(rels), len(over)))
    for c, p_raw, p in rels[:15]:
        print("    %s  %s%s" % (c, p_raw, "   ←过半" if (p is not None and p >= 50.0) else ""))
    ok, _over = qualified(corp, rels)
    print("\n结论：企业榜 %s · 持股过半 %s → %s" % (
        "有" if corp else "无", "有" if over else "无", "够格" if ok else "不够格"))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--book")
    ap.add_argument("--out")
    ap.add_argument("--probe")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--set-balance", type=float, default=None,
                    help="按控制台看到的余额设基准（分）")
    ap.add_argument("--floor", type=float, default=0.0,
                    help="本地估算余量低于此值就停（默认 0 = 不停，等接口自己报余额不足）")
    ap.add_argument("--dry-run", action="store_true",
                    help="只做网络评估（天眼查），打印关联主体数量与预算，不调用企查查")
    ap.add_argument("--group-threshold", type=int, default=GROUP_ENTITY_THRESHOLD,
                    help="控制企业数 ≥ 此值视为集团性质，本次跳过、末尾集中待确认（默认 %d）"
                         % GROUP_ENTITY_THRESHOLD)
    a = ap.parse_args()

    if a.set_balance is not None:
        set_baseline(a.set_balance, "命令行设定")
        print("已设企查查余额基准：%.0f 分" % a.set_balance)
        return

    core = load_core()
    rem = qcc_remaining()
    print("企查查本地估算余额：%s" % ("%.0f 分" % rem if rem is not None else "未设基准"), file=sys.stderr)

    if a.probe:
        probe(core, a.probe)
        print("\n天眼查调用:", core.NO_CALLS)
        return

    if not a.book:
        ap.error("需要 --book 或 --probe")
    wb = openpyxl.load_workbook(a.book)
    ws = wb[wb.sheetnames[0]]
    rows = []
    for r in range(2, ws.max_row + 1):
        nm = ws.cell(row=r, column=2).value
        if not nm:
            continue
        n = case_count(ws.cell(row=r, column=4).value)
        if MAX_CASES_FOR_EVAL is None or n <= MAX_CASES_FOR_EVAL:
            rows.append((r, nm, n))
    if a.limit:
        rows = rows[:a.limit]
    print("待评估 %d 家（门槛：%s）" % (len(rows),
          "无" if MAX_CASES_FOR_EVAL is None else "近两年案件数 ≤ %d" % MAX_CASES_FOR_EVAL),
          file=sys.stderr)

    # ---------- 预演：只打网络评估（花天眼查），不碰企查查 ----------
    if a.dry_run:
        md = ["| 企业 | 现有件数 | 企业榜 | 控制企业 | 持股过半 | 够格 | 预计企查查分 |",
              "|---|---|---|---|---|---|---|"]
        total_rel, big = 0, []
        for i, (r, nm, n) in enumerate(rows, 1):
            try:
                corp, allc, _ = signals(core, nm)
                rels = controlled(core, nm)
                yes, over = qualified(corp, rels)
            except Exception as e:
                print("[%d/%d] %-28s !! 失败：%s" % (i, len(rows), nm, e), file=sys.stderr)
                continue
            total_rel += len(rels)
            mark = "**集团？**" if len(rels) >= a.group_threshold else ""
            if len(rels) >= a.group_threshold:
                big.append((nm, len(rels)))
            md.append("| %s | %d | %d | %d | %d | %s | %d | %s"
                      % (nm, n, len(corp), len(rels), len(over),
                         "是" if yes else "否", len(rels) * 3, mark))
            print("[%d/%d] %-28s 控制企业 %2d 家 → 预计 %3d 分 %s"
                  % (i, len(rows), nm, len(rels), len(rels) * 3, mark), file=sys.stderr)
        print("\n【预演】未调用企查查。关联主体合计 %d 家，全量预计 %d 分。"
              % (total_rel, total_rel * 3), file=sys.stderr)
        if big:
            print("疑似集团性质（控制企业 ≥ %d 家）%d 家："
                  % (a.group_threshold, len(big)), file=sys.stderr)
            for nm, c in big:
                print("   %s：%d 家，约 %d 分" % (nm, c, c * 3), file=sys.stderr)
        print("天眼查调用 %d 次" % core.NO_CALLS)
        if a.out:
            Path(a.out).write_text("\n".join(md) + "\n", encoding="utf-8")
            print("已写出预演表:", a.out)
        return

    groups, skipped, failed, failed_tyc, stopped = [], [], [], [], []
    abort = False
    for i, (r, nm, n) in enumerate(rows, 1):
        if abort:
            break
        print("[%d/%d] %s（现有 %d 件）" % (i, len(rows), nm, n), file=sys.stderr)
        try:
            corp, allc, _ = signals(core, nm)
            rels = controlled(core, nm)
            yes, over = qualified(corp, rels)
            print("     企业榜 %d 条 | 持股过半 %d 家 → %s"
                  % (len(corp), len(over), "够格，查关联" if yes else "不够格，跳过"), file=sys.stderr)
            if not yes:
                continue
            if len(rels) >= a.group_threshold:
                skipped.append({"母公司": nm, "关联": len(rels), "预计分": len(rels) * 3})
                print("     ⚠ 控制企业 %d 家（≥%d），疑似集团性质 → 本次跳过，末尾待确认"
                      % (len(rels), a.group_threshold), file=sys.stderr)
                continue
            hits = []
            print("     关联主体 %d 个%s" % (len(rels),
                  "（预计企查查 %.0f 分）" % (len(rels) * 3)), file=sys.stderr)
            for rel, pct_raw, pct in rels:
                r = qcc_remaining()
                if a.floor and r is not None and r < a.floor and not has_cache(rel):
                    # 已缓存的主体不花钱，不该被余额拦住；未缓存的才停，并留痕
                    print("     !! 企查查余额估算剩 %.0f 分，低于底线 %.0f，停止补查" % (r, a.floor),
                          file=sys.stderr)
                    stopped.append((nm, rel, "余额 %.0f 分低于底线 %.0f，未查" % (r, a.floor)))
                    break
                docs, _, _, _notes = qcc_documents(rel)
                if _notes:
                    # 不能吞掉：余额不足等失败会让结果静默缺失，看起来像“这家没案子”
                    failed.append((nm, rel, "；".join(_notes)))
                    print("     !! 未查成：%s（%s）" % (rel, "；".join(_notes)), file=sys.stderr)
                    if any(credit_error(x) for x in _notes):
                        print("     !! 企查查余额耗尽，整轮停止", file=sys.stderr)
                        abort = True
                        break
                    continue
                lines, n_merged, n_kept, n_skip = build_cases(docs)
                trunc = qcc_truncation(rel)
                if trunc:
                    print("     !! 接口上限：%s（%s），件数为下限，需人工下载剩余文书"
                          % (rel, "、".join("%d年累计%d条仅取%d条" % t for t in trunc)),
                          file=sys.stderr)
                # 被截断的主体即使这批没捞到案子也要进附表，否则缺的那部分永远看不见
                if lines or trunc:
                    hits.append({"主体": rel, "持股": pct_raw, "件数": n_merged,
                                 "明细": lines, "截断": trunc})
            if hits:
                groups.append({"母公司": nm,
                               "上榜": [str(it.get("name")) for it in corp[:3]],
                               "企业榜条数": len(corp),
                               "关联": hits})
        except Exception as e:
            failed_tyc.append((nm, str(e)))
            print("     !! 未查成（天眼查）：%s" % e, file=sys.stderr)

    print("\n【结果】%d 家够格，其中 %d 家关联主体名下查到案件"
          % (len([g for g in groups]), len(groups)), file=sys.stderr)
    for g in groups:
        print("  %s：%d 个关联主体有案件" % (g["母公司"], len(g["关联"])), file=sys.stderr)
    if skipped:
        print("\n【待确认】以下 %d 家控制企业数量过大，疑似集团性质，本次已跳过："
              % len(skipped), file=sys.stderr)
        for s in skipped:
            print("   %s：控制企业 %d 家，全查约 %d 分"
                  % (s["母公司"], s["关联"], s["预计分"]), file=sys.stderr)
        print("   下一步：单独确认是否展开，或收窄到「只查持股过半」。", file=sys.stderr)
    if failed:
        print("\n【未查成】%d 个关联主体调用失败，结果未计入附表：" % len(failed), file=sys.stderr)
        for g, rel, why in failed:
            print("   %s → %s：%s" % (g, rel, why[:70]), file=sys.stderr)
        print("   下一步：充值后重跑，缓存里的 _err 会自动重试。", file=sys.stderr)
    if stopped:
        print("\n【因余额停止】%d 个关联主体未查，结果未计入附表：" % len(stopped), file=sys.stderr)
        for g, rel, why in stopped:
            print("   %s → %s：%s" % (g, rel, why), file=sys.stderr)
        print("   下一步：充值后重跑；已缓存的主体不会重复扣分。", file=sys.stderr)
    if failed_tyc:
        print("\n【未查成·天眼查】%d 家主体评估失败，已保留其原有行，不按缺失重算：" % len(failed_tyc),
              file=sys.stderr)
        for g, why in failed_tyc:
            print("   %s：%s" % (g, why[:70]), file=sys.stderr)
    if stopped:
        print("\n【因余额停止】%d 个关联主体未查，结果未计入附表：" % len(stopped), file=sys.stderr)
        for g, rel, why in stopped:
            print("   %s → %s：%s" % (g, rel, why), file=sys.stderr)
        print("   下一步：充值后重跑；已缓存的主体不会重复扣分。", file=sys.stderr)
    if failed_tyc:
        print("\n【未查成·天眼查】%d 家主体评估失败，已保留其原有行，不按缺失重算：" % len(failed_tyc),
              file=sys.stderr)
        for g, why in failed_tyc:
            print("   %s：%s" % (g, why[:70]), file=sys.stderr)

    print("天眼查调用 %d 次" % core.NO_CALLS)
    r2 = qcc_remaining()
    if r2 is not None:
        print("企查查本地估算余额：剩 %.0f 分" % r2, file=sys.stderr)

    if a.out:
        out = Path(a.out)
        # 等级按「自身 + 关联主体」合计重算，汇总行写清合计与拆解
        rel_total = {}
        for g in groups:
            rel_total[g["母公司"]] = sum(h["件数"] for h in g["关联"])
        skip_names = {n for n, _ in failed_tyc}
        for r in range(2, ws.max_row + 1):
            nm = ws.cell(r, 2).value
            if not nm:
                continue
            if nm in skip_names:
                continue      # 天眼查失败：保留原行，不按「关联 0」重算等级
            own = case_count(ws.cell(r, 4).value)
            rel = rel_total.get(nm, 0)
            total = own + rel
            ws.cell(r, 1).value = grade(total)
            lines = str(ws.cell(r, 4).value or "").split("\n")
            if total == 0:
                lines = ["近两年无知产案件"]
            elif rel:
                lines[0] = ("近两年知产案件合计 %d 件（自身 %d 件 + 关联主体名下 %d 件，明细见附表）"
                            % (total, own, rel))
            else:
                lines[0] = "近两年知产案件 %d 件" % total
            ws.cell(r, 4).value = "\n".join(lines)
        wb.save(out)   # 主表（等级已按合计重算）
        if groups:
            # 重跑时先把旧附表删掉，否则会叠出「关联主体明细1」
            for _name in [n for n in wb.sheetnames if n.startswith("关联主体明细")]:
                del wb[_name]
            ws2 = wb.create_sheet("关联主体明细")
            ws2.append(["主体组", "关联主体", "持股", "案件明细"])
            for c in ws2[1]:
                c.font = Font(bold=True)
                c.alignment = Alignment(horizontal="center", vertical="center")
            gi = 0
            for g in groups:
                for rel in g["关联"]:
                    det = "\n".join(rel["明细"])
                    miss = rel["件数"] - len(rel["明细"])
                    if miss > 0:
                        # 件数按归并组算，明细只列得出当事人的；差出来的几条必须说明，否则看表的人会以为漏了
                        det += "\n（另有 %d 件当事人与金额均无法解析，计入件数但未列明细）" % miss
                    if rel.get("截断"):
                        det += "\n（企查查接口上限：%s，需人工下载剩余裁判文书）" % "、".join(
                            "%d 年累计 %d 条仅取 %d 条" % t for t in rel["截断"])
                    ws2.append([g["母公司"], rel["主体"], rel["持股"], det])
                    fill = PatternFill("solid", fgColor=GROUP_FILLS[gi % len(GROUP_FILLS)])
                    for c in ws2[ws2.max_row]:
                        c.fill = fill
                        c.alignment = Alignment(wrap_text=True, vertical="top")
                gi += 1
            for col, w in zip("ABCD", [26, 30, 10, 96]):
                ws2.column_dimensions[col].width = w
            ws2.freeze_panes = "A2"
            wb.save(out)
        print("已写出:", out)


if __name__ == "__main__":
    main()
