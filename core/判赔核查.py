# -*- coding: utf-8 -*-
"""判赔核查：把工作簿里的每一笔判赔拿去北大法宝核一遍，结果写成新工作表。

全量核查：有一笔核一笔，不抽样。法宝查不到不改变表内金额，只如实标注。

用法（工作台会自动调用）：
    python 判赔核查.py --book "结果.xlsx"

输出：
    1. 在同一个工作簿里新增「判赔核查」表
    2. 同目录下 判赔核查_时间.csv
    3. 控制台打印汇总与不一致项
"""

import argparse
import csv
import importlib.util
import re
import sys
import time
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

BASE = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE))

spec = importlib.util.spec_from_file_location("verify", BASE / "判赔核验.py")
V = importlib.util.module_from_spec(spec)
spec.loader.exec_module(V)

spec2 = importlib.util.spec_from_file_location("portrait", BASE / "画像取数.py")
P = importlib.util.module_from_spec(spec2)
spec2.loader.exec_module(P)

CODE = re.compile(r"（\d{4}）[^\s，、]+?号(?:之[一二三四五六七八九十]+)?|晋\d+知民初\d+号")
AMT = re.compile(r"判赔\s*([\d,]+)")


def collect(book):
    """从工作簿里取出所有「有金额的判赔行」。同一案号只取一次。"""
    wb = openpyxl.load_workbook(book)
    cases, seen = [], set()
    for ws in wb.worksheets:
        if ws.title == "判赔核查":
            continue
        for r in ws.iter_rows(min_row=2):
            name = str(r[1].value or "")
            for line in str(r[3].value or "").split("\n"):
                m = AMT.search(line)
                if not m:
                    continue
                codes = CODE.findall(line)
                if not codes:
                    continue
                code = codes[0]
                key = V.norm(P.base_code(code))
                if key in seen:
                    continue
                seen.add(key)
                cases.append({"主体": name, "案号": code,
                              "金额": int(m.group(1).replace(",", ""))})
    return cases


def run(cases):
    rows = []
    for i, c in enumerate(cases, 1):
        d = V.fetch(c["案号"])
        if "_err" in d:
            rows.append([c["主体"], c["案号"], c["金额"], "", "接口失败", "", "", ""])
            print("[%d/%d] %s 接口失败" % (i, len(cases), c["案号"]), file=sys.stderr)
        elif "_miss" in d:
            rows.append([c["主体"], c["案号"], c["金额"], "", "未取到", "", "", ""])
            print("[%d/%d] %s 未取到" % (i, len(cases), c["案号"]), file=sys.stderr)
        else:
            got, seg = V.award(d.get("RefereeResult"))
            if got is None:
                verdict = "未解析"
            elif abs(got - c["金额"]) < 1:
                verdict = "一致"
            else:
                verdict = "不一致"
            rows.append([c["主体"], c["案号"], c["金额"],
                         "" if got is None else int(got), verdict, seg[:300],
                         d.get("Court") or "", d.get("LastInstanceDate") or ""])
            print("[%d/%d] %s %s 表 %s / 法宝 %s"
                  % (i, len(cases), c["案号"], verdict, c["金额"],
                     "" if got is None else int(got)), file=sys.stderr)
        time.sleep(0.4)
    return rows


def write_sheet(book, rows):
    wb = openpyxl.load_workbook(book)
    for name in [n for n in wb.sheetnames if n == "判赔核查"]:
        del wb[name]
    ws = wb.create_sheet("判赔核查")
    head = ["主体", "案号", "表内判赔", "法宝判赔", "判定", "法宝主文摘录", "法院", "审结日期"]
    ws.append(head)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
    red = PatternFill("solid", fgColor="F6DEDC")
    for row in rows:
        ws.append(row)
        if row[4] == "不一致":
            for cell in ws[ws.max_row]:
                cell.fill = red
    for col, w in zip("ABCDEFGH", [26, 34, 12, 12, 10, 60, 20, 12]):
        ws.column_dimensions[col].width = w
    for r in ws.iter_rows(min_row=2):
        for c in r:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    wb.save(book)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--book", required=True)
    a = ap.parse_args()

    cases = collect(a.book)
    if not cases:
        print("没有需要核查的判赔，跳过。")
        return

    try:
        V.auth()
    except Exception as e:
        print("法宝凭证缺失，跳过核查：%s" % e, file=sys.stderr)
        return

    print("待核查 %d 笔，逐笔调法宝…" % len(cases), file=sys.stderr)
    rows = run(cases)
    write_sheet(a.book, rows)

    csv_path = Path(a.book).with_name("判赔核查_%s.csv" % datetime.now().strftime("%Y%m%d_%H%M"))
    with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["主体", "案号", "表内判赔", "法宝判赔", "判定", "法宝主文摘录", "法院", "审结日期"])
        w.writerows(rows)

    from collections import Counter
    c = Counter(r[4] for r in rows)
    print("\n合计 %d 笔｜一致 %d｜不一致 %d｜未解析 %d｜未取到 %d｜接口失败 %d"
          % (len(rows), c["一致"], c["不一致"], c["未解析"], c["未取到"], c["接口失败"]))
    bad = [r for r in rows if r[4] == "不一致"]
    if bad:
        print("不一致项（已在新表标红）：")
        for r in bad:
            print("  %s 表 %s / 法宝 %s" % (r[1], r[2], r[3]))
    print("明细：%s" % csv_path.name)


if __name__ == "__main__":
    main()
