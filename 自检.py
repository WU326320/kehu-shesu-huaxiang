# -*- coding: utf-8 -*-
"""环境自检：优先使用包内便携 Python 与 tyc，不强制要求系统安装。"""
import importlib.util
import json
import os
import shutil
import socket
import subprocess
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent
TOKEN_FILE = BASE / "凭证.txt"
PKU_FILE = BASE / "法宝凭证.txt"
CACHE = BASE / "缓存"
OUT = BASE / "输出"
NODE_DIR = BASE / "runtime" / "node"
PY_EXE = BASE / "runtime" / "python" / "python.exe"

if (NODE_DIR / "tyc.cmd").exists():
    os.environ["PATH"] = str(NODE_DIR) + os.pathsep + os.environ.get("PATH", "")


def tyc_env():
    e = dict(os.environ)
    if (NODE_DIR / "tyc.cmd").exists():
        e["PATH"] = str(NODE_DIR) + os.pathsep + e.get("PATH", "")
    return e


def line(ok, name, detail=""):
    print(("[OK]   " if ok else "[FAIL] ") + name + ("  " + detail if detail else ""))
    return ok


def line_soft(ok, name, detail=""):
    """可选依赖：缺了不影响出表，只提示。"""
    print(("[OK]   " if ok else "[WARN] ") + name + ("  " + detail if detail else ""))
    return True


def main():
    results = []

    results.append(line(sys.version_info >= (3, 11), "Python 3.11+",
                        sys.version.split()[0]
                        + ("（包内便携运行时）" if Path(sys.executable).resolve() == PY_EXE.resolve()
                           else ("；包内已自带 runtime\\python，建议双击 启动工作台.bat" if PY_EXE.exists() else ""))))

    spec = importlib.util.find_spec("openpyxl")
    ver = ""
    if spec:
        import openpyxl
        ver = openpyxl.__version__
    results.append(line(spec is not None, "openpyxl",
                        ver or "未安装。请用 runtime\\python\\python.exe 运行本脚本，或重新解压完整绿色包"))

    bundled_tyc = NODE_DIR / "tyc.cmd"
    tyc = str(bundled_tyc) if bundled_tyc.exists() else shutil.which("tyc")
    results.append(line(tyc is not None, "tyc 命令",
                        (tyc + ("（包内）" if bundled_tyc.exists() else ""))
                        or "未找到，请确认 runtime\\node 已完整解压"))

    if tyc:
        try:
            v = subprocess.run("tyc -V", shell=True, capture_output=True, timeout=30, env=tyc_env())
            out = ((v.stdout or b"") + (v.stderr or b"")).decode("utf-8", "ignore").strip().splitlines()
            results.append(line(v.returncode == 0, "tyc 可执行", out[-1] if out else ""))
        except Exception as e:
            results.append(line(False, "tyc 可执行", str(e)[:80]))

    cfg = Path.home() / ".tyc" / "config.json"
    logged = False
    if cfg.exists():
        try:
            d = json.loads(cfg.read_text(encoding="utf-8"))
            logged = bool((d.get("oauth") or {}).get("refreshToken")
                          or (d.get("headers") or {}).get("Authorization"))
        except Exception:
            logged = cfg.stat().st_size > 0
    results.append(line(logged, "天眼查登录态",
                        "已登录" if logged else "未登录，请在工作台点「开始授权」完成登录"))

    token = ""
    if TOKEN_FILE.exists():
        for raw in TOKEN_FILE.read_text(encoding="utf-8-sig").splitlines():
            t = raw.strip()
            if t and not t.startswith("#"):
                token = t
                break
    results.append(line(token.startswith("Bearer "), "企查查凭证",
                        "已填写" if token.startswith("Bearer ") else "在工作台「企查查」口子填写，或写入 凭证.txt"))

    pku = ""
    if PKU_FILE.exists():
        for raw in PKU_FILE.read_text(encoding="utf-8-sig").splitlines():
            t = raw.strip()
            if t and not t.startswith("#"):
                pku = t
                break
    results.append(line_soft(bool(pku), "法宝凭证",
                             "已填写" if pku else "缺失，判赔核查会跳过（不影响出表）"))
    chk = BASE / "core" / "判赔核查.py"
    results.append(line(chk.exists(), "判赔核查脚本", str(chk.name)))

    try:
        socket.create_connection(("agent.qcc.com", 443), timeout=8).close()
        results.append(line(True, "企查查网络", "agent.qcc.com:443 可达"))
    except Exception as e:
        results.append(line(False, "企查查网络", str(e)[:80]))

    for d in (CACHE, OUT):
        try:
            d.mkdir(parents=True, exist_ok=True)
            probe = d / ".write_test"
            probe.write_text("x", encoding="utf-8")
            probe.unlink()
            results.append(line(True, "可写目录", str(d)))
        except Exception as e:
            results.append(line(False, "可写目录", str(d) + " " + str(e)[:60]))

    results.extend(check_rules())

    ok = all(results)
    print("")
    print("结论：" + ("全部通过，可以运行 启动工作台.bat" if ok else "还有 %d 项未通过" % results.count(False)))
    return 0 if ok else 1


def check_rules():
    """口径自检：不联网，确认 core 里的脚本就是当前规则。"""
    import importlib.util

    out = []
    core = BASE / "core"
    sys.path.insert(0, str(core))
    try:
        spec = importlib.util.spec_from_file_location("portrait", core / "画像取数.py")
        m = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(m)
    except Exception as e:
        return [line(False, "核心脚本加载", str(e)[:80])]

    try:
        ok1 = m.ip_kind("知识产权权属、侵权纠纷案件执行") is None
        out.append(line(ok1, "执行类不计件", "案由含执行且无之诉 → 不计"))
    except Exception as e:
        out.append(line(False, "执行类不计件", str(e)[:80]))

    try:
        ok2 = m.base_code("（2024）沪0104民初12909号之一") == "（2024）沪0104民初12909号"
        out.append(line(ok2, "之一并入基础案号", ""))
    except Exception as e:
        out.append(line(False, "之一并入基础案号", str(e)[:80]))

    try:
        amt, _ = m.damage({"裁判结果": "被告赔偿原告40,000元；被告某某公司在30,000元范围内承担连带责任"})
        out.append(line(amt == 40000, "连带范围不另计", "40,000" if amt == 40000 else str(amt)))
    except Exception as e:
        out.append(line(False, "连带范围不另计", str(e)[:80]))

    try:
        rows, groups, _, _ = m.build_cases([
            {"案号": "（2024）沪0104民初1号", "案由": "侵害商标权纠纷", "裁判日期": "2024-05-01",
             "文书标题": "民事判决书", "裁判结果": "被告赔偿原告1000元",
             "当事人": {"原告": ["甲公司"], "被告": ["乙店"]}},
            {"案号": "（2024）沪0104民初1号之一", "案由": "侵害商标权纠纷", "裁判日期": "2024-06-01",
             "文书标题": "民事裁定书", "裁判结果": "本案转为普通程序",
             "当事人": {"申请人": ["甲公司"], "被申请人": ["乙店"]}},
        ])
        out.append(line(groups == 1, "同案多文书计一件", "1 件 / %d 行" % len(rows)))
    except Exception as e:
        out.append(line(False, "同案多文书计一件", str(e)[:80]))

    return out


if __name__ == "__main__":
    sys.exit(main())
