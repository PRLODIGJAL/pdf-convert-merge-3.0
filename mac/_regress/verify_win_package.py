#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Windows 交付包验收：静态检查 + 把包拷到临时目录真跑一遍环境流程。

跑法：.venv/bin/python -B _regress/verify_win_package.py
"""
from __future__ import annotations

import ast
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

DEV = Path(__file__).resolve().parent.parent
WIN = Path("/Users/langzi/Desktop/PDF转换拼接3.0_Windows版")
BODY = ("seq_core.py", "seq_app.py", "setup_env.py")
# 启动器的名字不写死：Mac 侧有什么 .command 就拉黑什么 —— 以后再改名不用回来改。
# LEGACY 是历史上用过的名字，改名后万一有残留也要报出来。
LEGACY_MAC_LAUNCHERS = ("run_macos.command",)
FORBIDDEN = ("_regress", "selftest.py", "__pycache__", ".venv", ".DS_Store") \
            + tuple(sorted({p.name for p in DEV.glob("*.command")} | set(LEGACY_MAC_LAUNCHERS)))

# ⚠️ 2026-10-10 起这里**不再** import seq_core 去核说明文件 —— 交付包里已经
#    没有说明书了（见下面「不再放使用说明」那段）。

PASS: list[str] = []
FAIL: list[str] = []


def check(name: str, cond: bool, detail: str = "") -> None:
    (PASS if cond else FAIL).append(name)
    print(f"  {'OK ' if cond else 'NG '} {name}" + (f"  <{detail}>" if detail and not cond else ""))


def md5(p: Path) -> str:
    import hashlib

    return hashlib.md5(p.read_bytes()).hexdigest()


def third_party_imports(py: Path) -> set[str]:
    """程序实际 import 的**第三方**库（用来核对 setup_env 的依赖清单）。

    同目录下自己的模块（seq_core 之类）不算第三方 —— 漏掉这一步就会要求
    setup_env 去 pip install 一个根本不存在的包。
    """
    tree = ast.parse(py.read_text(encoding="utf-8"))
    std = set(sys.stdlib_module_names)
    local = {p.stem for p in DEV.glob("*.py")}
    out: set[str] = set()
    for node in ast.walk(tree):
        names: list[str] = []
        if isinstance(node, ast.Import):
            names = [a.name.split(".")[0] for a in node.names]
        elif isinstance(node, ast.ImportFrom) and node.module:
            names = [node.module.split(".")[0]]
        for top in names:
            if top not in std and top not in local:
                out.add(top)
    return out


def main() -> int:
    print("【静态】交付包本身")
    check("交付包存在", WIN.is_dir(), str(WIN))

    print("\n  · 程序本体逐字节一致")
    for name in BODY:
        check(f"{name} 与开发目录一致",
              (WIN / name).is_file() and md5(WIN / name) == md5(DEV / name))

    print("\n  · .bat 的编码与结构")
    bats = sorted(WIN.glob("*.bat"))
    check("包里有 .bat 启动器", len(bats) >= 1, str([b.name for b in bats]))
    for b in bats:
        raw = b.read_bytes()
        check(f"{b.name} 是 CRLF 行尾",
              raw.count(b"\r\n") > 0 and raw.count(b"\n") == raw.count(b"\r\n"))
        check(f"{b.name} 没有 BOM", not raw.startswith(b"\xef\xbb\xbf"))
        bad = [ch for ch in raw.decode("latin-1") if ord(ch) > 127]
        check(f"{b.name} 命令行纯 ASCII", not bad, str(sorted(set(bad))[:5]))
        lines = raw.decode("ascii").splitlines()
        labels = {m.group(1).lower() for ln in lines
                  if (m := re.match(r"^\s*:([A-Za-z0-9_]+)", ln))}
        gotos = {m.group(1).lower() for ln in lines
                 if (m := re.search(r"goto\s+(\w+)", ln, re.I))}
        check(f"{b.name} 的 goto 标签都存在", not (gotos - labels), str(sorted(gotos - labels)))
    check("没有把 Mac 的启动器混进来", not list(WIN.glob("*.command")))

    # ── 2026-10-10：用户只会开 CMD、而且被占位符坑过一次 ──────────────────
    #    我上一条给的排查命令里写了 `<Windows包目录>` 当占位符，他**原样粘进去**了，
    #    于是 `cd /d <Windows包目录>` 报「命令语法不正确」，后面每条都
    #    「系统找不到指定的路径」。根子不在他，在于**让用户手打路径这件事本身就是错的**。
    #    ⇒ 交付包必须自带一个双击就跑的体检入口，里面连一个需要用户填的路径都没有。
    doc = WIN / "运行前的环境准备.bat"
    check("⭐ 交付包里有「运行前的环境准备.bat」（排问题的统一入口，双击就跑）", doc.is_file())
    if doc.is_file():
        dtext = doc.read_text(encoding="ascii")
        check("⭐⭐ 环境准备.bat 靠 `%~dp0` 定位自己（工具放哪个盘都对，**不需要用户填任何路径**）",
              'pushd "%~dp0"' in dtext)
        check("⭐⭐ 环境准备.bat 里**不许出现尖括号占位符**"
              "（`<Windows包目录>` 那次就是被它坑的 —— 用户会原样粘进 CMD）",
              "<Windows" not in dtext and "<包目录>" not in dtext
              and "<路径>" not in dtext and "<目录>" not in dtext)
        check("⭐ 环境准备.bat 先在缺环境时自己把环境备好（`setup_env.py --setup`，只装环境不开界面）",
              "setup_env.py" in dtext and "--setup" in dtext)
        check("⭐ 环境准备.bat 跑的是「只看不动」的四个入口（不转文件、不碰原件）",
              "--show-rules" in dtext and "--show-cad" in dtext
              and "--show-bartender" in dtext)
        check("⭐ 环境准备.bat 支持**把一份文件拖到它图标上**试转（`%~1` + `--check-convert`）",
              "%~1" in dtext and "--check-convert" in dtext and "%~f1" in dtext)
        check("⭐ 环境准备.bat 用 `py -3` → `python` 两级找解释器（跟启动器同一套，别另起一套）",
              "py -3" in dtext and "python -c" in dtext)

    # ── 2026-10-10 用户拍板：交付包**不再放使用说明** ──────────────────────
    #    他的原话：「以后没必要出使用说明，不管哪个版本，节省点 token」。
    #    这份 txt 每改一次界面文案都得跟着手工重写一遍，而界面本身每一项都有
    #    悬浮提示、他自己从来不看 —— 它是纯开销，而且**会腐烂**（界面改了说明
    #    没改，读到的还是旧话）。我这边要对照的东西一律写进记忆 / skill，不进包。
    #    ⇒ 这里改成**反向哨兵**：包里有它反而要红，免得哪天又被补回来。
    print("\n  · 不再放使用说明（反向哨兵）")
    for n in ("使用说明.txt", "说明书.txt", "使用说明.md", "使用说明.doc",
              "README.txt", "README.md"):
        check(f"包里没有 {n}（2026-10-10 起不再出说明书）", not (WIN / n).exists())

    print("\n  · 不该出现的文件")
    for n in FORBIDDEN:
        check(f"没有混进 {n}", not (WIN / n).exists())

    print("\n  · setup_env.py 与程序的一致性")
    se = WIN / "setup_env.py"
    check("setup_env.py 在包里", se.is_file())
    if se.is_file():
        src = se.read_text(encoding="utf-8")
        used: set[str] = set()
        for name in ("seq_core.py", "seq_app.py"):
            used |= third_party_imports(DEV / name)
        need = src.lower()
        for lib in sorted(used):
            check(f"依赖清单含 {lib}", lib.lower() in need, f"程序实际用到 {sorted(used)}")
        m = re.search(r"BASE_NEED\s*=\s*\(([^)]*)\)", src)
        check("能读出依赖清单", bool(m))
        if m:
            check("pypdf 版本下限不低于 3.16（用到 Transformation 接口）",
                  "3.16" in m.group(1), m.group(1))
        # ── v3.0 新增「Office 转 PDF」后，依赖清单必须按平台分 ────────────
        # Windows 的转换引擎走 win32com（pywin32），macOS 走 AppleScript 用不上它；
        # 平白在 Mac 上装一个 10MB 的包没必要，而 Windows 上漏装则「界面能开、
        # 一点转换就报错」—— 两种错法都得在这里挡住。
        m2 = re.search(r"WIN_NEED\s*=\s*\(([^)]*)\)", src)
        check("⭐ Windows 那份依赖单列了一张表", bool(m2), str(m2))
        if m2:
            check("⭐ Windows 那份里有 pywin32（Office 转 PDF 靠它）",
                  "pywin32" in m2.group(1), m2.group(1))
        check("⭐ 装什么由 need_list() 按平台算（不写死一张单子）",
              "def need_list" in src and "os.name" in src)
        check("⭐ 自检 import 的东西也按平台拼（防「装完仍报不可用」）",
              "def check_imports" in src and "PLATFORM_IMPORTS" in src)
        # ── 2026-10-09：界面库（Tk）**能不能真的弹出来**要单独探一次 ────────
        # 踩过的坑：venv 不会复制 Tcl/Tk 的库目录，而 Tcl 只在 venv 里找它，
        # 于是「依赖全齐、import 全通、点开就是打不开窗口」——而且双击启动器时
        # 窗口一闪都不闪，用户只会以为是自己电脑的问题。
        check("⭐ 环境自检会真试一次「界面能不能起来」（只 import tkinter 不算数）",
              "def probe_tk" in src and "TK_PROBE" in src)
        check("  ↳ 用的必须是真 Tk()（Tcl() 不走 Tk 那套库查找，会假通过）",
              "tkinter.Tk()" in src)
        check("⭐ 界面库缺失时会自动补 Tcl/Tk 目录链接",
              "def fix_tk_library" in src and "symlink" in src)
        check("  ↳ 补链接问的是 venv 自己的 base_prefix（不拿 setup_env 的顶替）",
              "sys.base_prefix" in src)
        check("入口文件名指向包内真实文件", "seq_app.py" in src and (WIN / "seq_app.py").is_file())

    print("\n【功能】把包拷到临时目录真跑一遍")
    tmp = Path(tempfile.gettempdir()) / "pdfseq_win_check"
    if tmp.exists():
        shutil.rmtree(tmp)
    tmp.mkdir(parents=True)
    work = tmp / "包"
    shutil.copytree(WIN, work, ignore=shutil.ignore_patterns(".venv", ".DS_Store"))

    py = sys.executable
    r = subprocess.run([py, "-B", "setup_env.py", "--check"], cwd=work,
                       capture_output=True, text=True)
    check("全新目录自检：退出码 1", r.returncode == 1, str(r.returncode))
    check("并说明是「没装过」而不是含糊报错",
          "不存在" in (r.stdout or "") and "不可用" in (r.stdout or ""), (r.stdout or "")[-120:])

    r = subprocess.run([py, "-B", "setup_env.py", "--setup"], cwd=work,
                       capture_output=True, text=True)
    setup_out = (r.stdout or "") + (r.stderr or "")
    check("--setup 首次安装成功", r.returncode == 0, setup_out[-300:])
    check("确实建了 .venv", (work / ".venv").is_dir())
    check("⭐ 依赖是「按平台算出来的」那句话打出来了（看得见装什么）",
          "安装依赖" in setup_out and "pypdf" in setup_out, setup_out[-300:])
    check("⭐ 在 macOS 上不会去装 pywin32（那是 Windows 才要的）",
          "pywin32" not in setup_out, setup_out[-300:])
    vpy = work / ".venv" / "bin" / "python"
    check("环境里的解释器存在", vpy.exists())

    r = subprocess.run([py, "-B", "setup_env.py", "--check"], cwd=work,
                       capture_output=True, text=True)
    check("装完再自检：可用", r.returncode == 0 and "可用" in (r.stdout or ""),
          (r.stdout or "")[-200:])
    check("自检报出依赖版本", "pypdf" in (r.stdout or ""), (r.stdout or "")[-120:])

    r = subprocess.run([py, "-B", "setup_env.py", "--check"], cwd=work,
                       capture_output=True, text=True)
    check("第二次仍是「可用」快路径", r.returncode == 0 and "可用" in (r.stdout or ""))

    # 解释器断链（换机 / 原 Python 卸载）能被识别
    if vpy.exists():
        bak = vpy.with_suffix(".bak")
        shutil.copy2(vpy, bak)
        vpy.unlink()
        r = subprocess.run([py, "-B", "setup_env.py", "--check"], cwd=work,
                           capture_output=True, text=True)
        check("解释器断链时自检能识别出不可用", r.returncode == 1, (r.stdout or "")[-200:])
        shutil.copy2(bak, vpy)

    # 用户自己放了个普通 .venv 文件夹 —— 绝不能误删
    tmp2 = tmp / "包2"
    shutil.copytree(WIN, tmp2, ignore=shutil.ignore_patterns(".venv", ".DS_Store"))
    (tmp2 / ".venv").mkdir()
    (tmp2 / ".venv" / "这是我的东西.txt").write_text("别删", encoding="utf-8")
    r = subprocess.run([py, "-B", "setup_env.py", "--setup"], cwd=tmp2,
                       capture_output=True, text=True)
    check("拒绝删除不像 venv 的 .venv 目录", "不像虚拟环境" in (r.stdout or ""),
          (r.stdout or "")[-200:])
    check("用户文件原封不动", (tmp2 / ".venv" / "这是我的东西.txt").is_file())

    shutil.rmtree(tmp, ignore_errors=True)
    print(f"\n通过 {len(PASS)} 项，失败 {len(FAIL)} 项")
    for n in FAIL:
        print("  失败：" + n)
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
