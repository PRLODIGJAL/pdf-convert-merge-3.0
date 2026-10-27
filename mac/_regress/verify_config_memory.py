#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""「记住上次选择」回归 —— 真建界面对象，读回填了什么、落盘了什么。

⚠️ CONFIG_PATH 一律重定向到 /tmp：v1 差点用测试路径覆盖掉用户真实记录。
跑法：.venv/bin/python -B _regress/verify_config_memory.py
"""
from __future__ import annotations

import json
import os
import shutil
import sys
import tempfile
import tkinter as tk
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import seq_app as A  # noqa: E402
import seq_core as C  # noqa: E402

REAL_CONFIG_PATH = A.CONFIG_PATH       # 重定向之前先留一份，用来验名字
TMP = Path(tempfile.gettempdir()) / "pdfseq_cfg"
if TMP.exists():
    shutil.rmtree(TMP)
TMP.mkdir(parents=True)
A.CONFIG_PATH = TMP / "config.json"

PASS: list[str] = []
FAIL: list[str] = []


def check(name: str, cond: bool, detail: str = "") -> None:
    (PASS if cond else FAIL).append(name)
    print(f"  {'OK ' if cond else 'NG '} {name}" + (f"  <{detail}>" if detail and not cond else ""))


def new_app(root: tk.Tk) -> A.App:
    app = A.App(root)
    root.update()
    return app


def main() -> int:
    print("[1] 存与取")
    root = tk.Tk()
    root.withdraw()
    app = new_app(root)
    src = TMP / "原件"
    src.mkdir()
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from selftest import build_pdf  # noqa: E402
    build_pdf(TMP / "扫描件.pdf", [(595, 842, "SCAN-1")])
    app._add_sources([str(src)])
    app.vars["scan"].set(str(TMP / "扫描件.pdf"))
    app.vars["reverse"].set(True)
    app.vars["scale"].set(False)
    app.vars["suffix"].set("_已签字首页")
    app.vars["conv_suffix"].set("_已转")          # ⭐ 两个模式各记各的后缀
    root.update()
    app._save_config()
    check("记录文件写出来了", A.CONFIG_PATH.is_file())
    data = json.loads(A.CONFIG_PATH.read_text(encoding="utf-8"))
    # ⭐ v3.0：路径类的东西一个都不记（2026-10-08 用户要求）——
    #    上一轮的 PDF 早删了，记着只会指着一个不存在的东西。
    check("⭐ 原件来源清单**不落盘**（加了也不写进配置）",
          "sources" not in data, str(sorted(data)))
    check("⭐ 扫描件路径**不落盘**", "scan" not in data, str(sorted(data)))
    check("记录里不再有 src / out / recursive 这几个老键",
          not ({"src", "out", "recursive"} & set(data)), str(sorted(data)))
    check("该记的照记：两份后缀 / 缩放 / 页序",
          data.get("reverse") is True and data.get("scale") is False
          and data.get("suffix") == "_已签字首页"
          and data.get("conv_suffix") == "_已转", str(data))
    check("记了配置版本号", data.get("_version") == A.CONFIG_VERSION, str(data.get("_version")))
    check("记录了窗口大小", isinstance(data.get("window"), list) and len(data["window"]) == 2,
          str(data.get("window")))
    # 这条盯的是「以后谁往 self.vars 里加了个带路径的新输入框，记得顺手想想要不要落盘」
    check("落盘的键不多不少，正好是白名单 + 版本号 + 窗口（白名单说了算）",
          set(data) == set(A.CONFIG_PERSIST_VARS) | {"_version", "window"},
          str(sorted(data)))
    app.on_close()
    try:
        root.destroy()
    except tk.TclError:
        pass

    print("\n[2] 下次打开：只填回「跟这一批无关」的选项")
    root2 = tk.Tk()
    root2.withdraw()
    app2 = new_app(root2)
    app2.on_mode("swap")            # [2]/[2a] 测的是替换模式的后缀恢复，显式切回 swap 起步
    check("⭐ 原件来源**不**恢复（清单是空的，重新挑一次）",
          app2.sources == [], str(app2.sources))
    check("清单控件里也是空的", app2.lst_sources.size() == 0, str(app2.lst_sources.size()))
    check("⭐ 扫描件**不**恢复", app2.vars["scan"].get() == "", app2.vars["scan"].get())
    check("倒序勾回来了", app2.vars["reverse"].get() is True)
    check("缩放关的状态也回来了", app2.vars["scale"].get() is False)
    check("替换模式那份后缀回来了", app2.vars["suffix"].get() == "_已签字首页")
    check("⭐ 转换模式那份后缀也回来了（两份各记各的）",
          app2.vars["conv_suffix"].get() == "_已转", app2.vars["conv_suffix"].get())
    check("状态栏**不再**提「已填好上次的路径」（已经没有路径可填）",
          "已自动填好" not in app2.status.cget("text")
          and "已经不存在" not in app2.status.cget("text"), app2.status.cget("text"))
    check("白名单就是「两份后缀 / 缩放 / 页序」这四项（表驱动，改口径只动这张表）",
          tuple(A.CONFIG_PERSIST_VARS) == ("suffix", "conv_suffix", "scale", "reverse"),
          str(A.CONFIG_PERSIST_VARS))
    check("老的 CONFIG_KEEP_KEYS 已经拆掉（免得两套口径并存）",
          not hasattr(A, "CONFIG_KEEP_KEYS"))
    # ⚠️ 这条是 2026-10-08 修掉的启动 bug：trace 挂在 _load_config 之后，
    #    恢复出来的 suffix 不会触发刷新 → 提示格一直是空的（看着像没填）。
    check("恢复出来的后缀也显示在右边那格（不是空壳）",
          "_已签字首页" in app2.lbl_suffix.cget("text"), app2.lbl_suffix.cget("text"))
    check("填回后缀之后「开始替换」才可能可用（后缀是必填闸门）",
          app2._suffix_ready() is True and str(app2.btn_run["state"]) == "disabled",
          f"ready={app2._suffix_ready()} state={app2.btn_run['state']} plan={app2.plan}")

    print("\n[2a] ⭐ 两个模式的后缀**不共享**（2026-10-08 用户要求）")
    # 用户原话：「转换那边的后缀和 PDF 拼接那边的后缀不要共享记忆 —— 转换那边大概率
    # 不用加后缀，PDF 拼接这边又必须加，这两个记忆不能共享。」
    check("后缀变量表就是两条：swap→suffix、convert→conv_suffix",
          dict(A.SUFFIX_VAR_BY_MODE) == {"swap": "suffix", "convert": "conv_suffix"},
          str(A.SUFFIX_VAR_BY_MODE))
    check("启动落在替换模式，输入框显示的是替换那份",
          app2.ent_suffix.get() == "_已签字首页", repr(app2.ent_suffix.get()))
    app2.on_mode("convert")
    root2.update()
    check("⭐ 切到转换模式，那一格立刻换成**转换那份**（不是替换那份）",
          app2.ent_suffix.get() == "_已转", repr(app2.ent_suffix.get()))
    check("转换模式的悬浮说明换成「可以不填」那一份",
          "可以不填" in app2._tip_suffix.text
          and "可不填" in app2._tip_suffix_need.text, app2._tip_suffix.text[:40])
    app2.vars["conv_suffix"].set("")
    root2.update()
    check("⭐ 把转换那份清空，替换那份**一点也不受影响**",
          app2.vars["suffix"].get() == "_已签字首页", repr(app2.vars["suffix"].get()))
    check("转换模式后缀空着 → 它自己那一侧闸门是开的（转换不要求填后缀）",
          app2._suffix_ready() is False
          and str(app2.btn_convert["state"]) in ("normal", "disabled"),
          f"ready={app2._suffix_ready()}")
    app2.on_mode("swap")
    root2.update()
    check("⭐ 切回替换模式，又是替换那份，且「必填」照样达标",
          app2.ent_suffix.get() == "_已签字首页" and app2._suffix_ready() is True,
          f"{app2.ent_suffix.get()!r} / {app2._suffix_ready()}")
    check("替换模式的悬浮说明换回「必填」那一份（不是转换那份）",
          "必填" in app2._tip_suffix.text and "可以不填" not in app2._tip_suffix.text,
          app2._tip_suffix.text[:40])
    app2.vars["conv_suffix"].set("_已转")          # 还原，别影响后面几节

    print("\n[2b] 后缀为空时恢复配置 → 闸门照样关着")
    A.CONFIG_PATH.write_text(json.dumps(
        {"sources": [str(src)], "scan": str(TMP / "扫描件.pdf"), "suffix": ""},
        ensure_ascii=False), encoding="utf-8")
    app2b = new_app(root2)
    app2b.on_mode("swap")           # [2b] 测的是替换模式的后缀为空 → 闸门关着
    check("空后缀恢复回来也是空的", app2b.vars["suffix"].get() == "")
    check("后缀没填 → 不认为能跑", app2b._suffix_ready() is False)
    check("后缀提示写成红字「不填不让跑」",
          "不填不让跑" in app2b.lbl_suffix.cget("text"), app2b.lbl_suffix.cget("text"))
    check("同一份配置里的 sources / scan 照样不恢复",
          app2b.sources == [] and app2b.vars["scan"].get() == "",
          f"{app2b.sources} / {app2b.vars['scan'].get()!r}")

    print("\n[2c] 老配置里的路径键（src / sources / scan）一律忽略")
    A.CONFIG_PATH.write_text(json.dumps(
        {"src": str(src), "sources": [str(src)], "scan": str(TMP / "扫描件.pdf"),
         "suffix": "-x"}, ensure_ascii=False), encoding="utf-8")
    app2c = new_app(root2)
    check("老 src 不再迁成清单（v3.0 起不记路径）",
          app2c.sources == [], str(app2c.sources))
    check("老 sources 也不恢复", app2c.sources == [], str(app2c.sources))
    check("老 scan 也不恢复", app2c.vars["scan"].get() == "", app2c.vars["scan"].get())
    check("同一份配置里该记的照样记（suffix）", app2c.vars["suffix"].get() == "-x",
          app2c.vars["suffix"].get())
    app2c._save_config()
    saved = json.loads(A.CONFIG_PATH.read_text(encoding="utf-8"))
    check("再存一次，老键被整份重写冲掉（不用专门写迁移代码）",
          not ({"src", "sources", "scan"} & set(saved)), str(sorted(saved)))

    print("\n[3] 配置里留着已经不存在的路径 → 不恢复、也不报错")
    (TMP / "原件").rmdir()
    A.CONFIG_PATH.write_text(json.dumps(
        {"sources": [str(TMP / "原件")], "scan": str(TMP / "扫描件.pdf"),
         "suffix": "-y"}, ensure_ascii=False), encoding="utf-8")
    app3 = new_app(root2)
    check("不恢复那条失效路径",
          app3.sources == [] and app3.vars["scan"].get() == "",
          f"{app3.sources} / {app3.vars['scan'].get()!r}")
    check("也**不**红字抱怨「已经不存在」（没有路径就不会有这种坑）",
          "已经不存在" not in app3.status.cget("text"), app3.status.cget("text"))
    check("状态栏停在默认那句「就绪…」", app3.status.cget("text").startswith("就绪"),
          app3.status.cget("text"))

    print("\n[4] 记录文件坏了 / 不是 dict")
    A.CONFIG_PATH.write_text("{ 这不是 json", encoding="utf-8")
    app4 = new_app(root2)
    check("读不出来只提示，不崩", "读不出来" in app4.status.cget("text")
          and app4.sources == [], app4.status.cget("text"))
    A.CONFIG_PATH.write_text("[1,2,3]", encoding="utf-8")
    app5 = new_app(root2)
    check("不是 dict 也照样能开", app5.sources == [])

    print("\n[5] 不认识的键 / 类型对不上")
    A.CONFIG_PATH.write_text(json.dumps(
        {"sources": [str(TMP)], "reverse": "不是布尔", "未来才会有的键": 123},
        ensure_ascii=False), encoding="utf-8")
    app6 = new_app(root2)
    check("不认识的键忽略，版本升级不炸", "未来才会有的键" not in A.CONFIG_PERSIST_VARS)
    check("类型对不上的那项跳过，其余键不受牵连",
          app6.vars["reverse"].get() is False and app6.sources == [],
          f"{app6.vars['reverse'].get()!r} / {app6.sources}")

    print("\n[6] 原子写失败 → 退一步直接写，且不留 .tmp")
    real_replace = os.replace
    calls = {"n": 0}

    def boom(a, b):
        calls["n"] += 1
        raise OSError("模拟安全软件拦住重命名")

    os.replace = boom
    try:
        app6.vars["suffix"].set("-回退")
        root2.update()
        app6._save_config()
    finally:
        os.replace = real_replace
    check("退路把配置存下来了",
          A.CONFIG_PATH.is_file() and "-回退" in A.CONFIG_PATH.read_text(encoding="utf-8"),
          A.CONFIG_PATH.read_text(encoding="utf-8") if A.CONFIG_PATH.is_file() else "没写")
    check("没有留下 .tmp 垃圾",
          not list(TMP.glob("config.json.tmp")), str([p.name for p in TMP.glob("*")]))

    print("\n[7] 记忆口径：选项照记，路径一个字都不进配置")
    check("已删掉「忘记记住的文件夹」入口", not hasattr(A.App, "on_forget"))
    check("已删掉「加文件夹」入口（用户要求：一次多选就够）",
          not hasattr(A.App, "on_add_folder")
          and not hasattr(A.App, "_pick_many_folders"))
    app6._add_sources([str(TMP / "回退目录")])       # 清单里塞一条路径
    root2.update()
    app6._save_config()
    body = A.CONFIG_PATH.read_text(encoding="utf-8")
    check("清单里的路径没溜进配置文件（后缀记了、路径没记）",
          "回退目录" not in body and "-回退" in body, body)
    check("记录文件仍在", A.CONFIG_PATH.exists())

    print("\n[8] 和 v1 各记各的（3.0 与 2.0 暂时共用一份，是否拆开待用户拍板）")
    check("v2/v3 的记录文件名和 v1 不同（两边各记各的）",
          REAL_CONFIG_PATH.name == ".pdf_batch_cover_swap.json"
          and "firstpage" not in REAL_CONFIG_PATH.name, REAL_CONFIG_PATH.name)
    check("3.0 没有偷偷换过记录文件名（共用这一份是当前口径，不是漏改）",
          REAL_CONFIG_PATH.name == ".pdf_batch_cover_swap.json",
          REAL_CONFIG_PATH.name)
    check("配置文件的 schema 版本号升到了 3（v3.0 起不再记任何路径）",
          A.CONFIG_VERSION == 3, str(A.CONFIG_VERSION))

    print("\n[9] 窗口大小只恢复合理范围")
    check("太小的尺寸不恢复（免得开出一个看不见的窗）",
          A.fit_geometry([10, 10], 1920, 1080) is None, str(A.fit_geometry([10, 10], 1920, 1080)))
    check("比屏幕还大也不恢复", A.fit_geometry([5000, 3000], 1920, 1080) is None)
    check("正常尺寸照常恢复", A.fit_geometry([1100, 700], 1920, 1080) == "1100x700")
    check("不是两个数就不认", A.fit_geometry([1100], 1920, 1080) is None
          and A.fit_geometry("1100x700", 1920, 1080) is None)
    check("类型对不上不认", A.fit_geometry([1100, "abc"], 1920, 1080) is None)

    print("\n[10] 初始窗口：够大且不超出屏幕")
    check("大屏上用默认尺寸",
          A.default_geometry(1920, 1080) == f"{A.GEOM_DEFAULT[0]}x{A.GEOM_DEFAULT[1]}",
          A.default_geometry(1920, 1080))
    check("矮屏笔记本上要收一收高度（别被任务栏切掉）",
          A.default_geometry(1366, 768) == "1080x700", A.default_geometry(1366, 768))
    for sw, sh in ((1920, 1080), (1440, 900), (1366, 768), (1280, 800)):
        w, h = (int(x) for x in A.default_geometry(sw, sh).split("x"))
        check(f"屏幕 {sw}×{sh} 上开出的窗口不超出屏幕", w <= sw and h <= sh,
              f"{w}x{h}")

    for r in (root2,):
        try:
            r.destroy()
        except tk.TclError:
            pass
    shutil.rmtree(TMP, ignore_errors=True)
    print(f"\n通过 {len(PASS)} 项，失败 {len(FAIL)} 项")
    for n in FAIL:
        print("  失败：" + n)
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
