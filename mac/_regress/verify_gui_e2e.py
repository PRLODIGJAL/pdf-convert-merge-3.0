#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""界面驱动端到端 —— 真建界面对象、真点按钮，验证 v3.0 的顺序配对流程。

⚠️ 必须先把 CONFIG_PATH 指到 /tmp 再建界面，否则会把用户真实的记录覆盖掉。
跑法：.venv/bin/python -B _regress/verify_gui_e2e.py

v3.0 换掉的输入输出模型（这一版回归的重点）：
    · 没有「输出文件夹」了 —— 产物写在**每份原件旁边**；
    · 因此后缀**必填**，空后缀时「开始替换」直接禁用、core 也拒绝；
    · 覆盖不再有勾选框，改成运行时弹窗（是 / 否 / 取消）让人点一次头（只报份数、不列文件名）；
    · ⭐ 两个模式的「输出文件名后缀」**各记各的**（见 SUFFIX_VAR_BY_MODE）；
    · ⭐ 选完扫描件**不再把队列作废**，而是自动把配对补上（见 [2c] 那一节）。
"""
from __future__ import annotations

import shutil
import sys
import tempfile
import time
import tkinter as tk
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

TMP = Path(tempfile.gettempdir()) / "pdfseq_gui_e2e"
if TMP.exists():
    shutil.rmtree(TMP)
TMP.mkdir(parents=True)

import seq_app as A  # noqa: E402
import seq_core as C  # noqa: E402

A.CONFIG_PATH = TMP / "config.json"          # 别碰用户真实记录
sys.path.insert(0, str(ROOT / "_regress"))
from selftest import build_pdf, page_texts, page_sizes  # noqa: E402

SUF = "_已签字首页"          # 本脚本统一的产物后缀
PASS: list[str] = []
FAIL: list[str] = []
ASK_YES = True               # askokcancel / askyesno 的答案
ASK_OVERWRITE: bool | None = True   # askyesnocancel 的答案（覆盖确认窗）
WARNED: list[str] = []
ASKED: list[str] = []        # 覆盖确认窗被问到时，把文案记下来


class _DropEvent:
    """够用的假 `<Drop>` 事件 —— `App._on_drop` 只读 `.data` 这一个字段。"""

    def __init__(self, data: str) -> None:
        self.data = data


def check(name: str, cond: bool, detail: str = "") -> None:
    (PASS if cond else FAIL).append(name)
    print(f"  {'OK ' if cond else 'NG '} {name}" + (f"  <{detail}>" if detail and not cond else ""))


def pump(root: tk.Tk, done, timeout: float = 8.0) -> bool:
    """把 Tk 事件循环泵起来，等后台线程回调完成。"""
    end = time.time() + timeout
    while time.time() < end:
        root.update()
        if done():
            return True
        time.sleep(0.02)
    return False


# ── 假的消息框：不阻塞、并记录被调用的内容 ──────────────────────────────
def fake_askokcancel(title, msg, **kw):
    return ASK_YES


def fake_askyesnocancel(title, msg, **kw):
    ASKED.append(str(msg))
    return ASK_OVERWRITE


def fake_showinfo(title, msg, **kw):
    return None


def fake_showwarning(title, msg, **kw):
    WARNED.append(str(msg))
    return None


A.messagebox.askokcancel = fake_askokcancel
A.messagebox.askyesnocancel = fake_askyesnocancel
A.messagebox.showinfo = fake_showinfo
A.messagebox.showwarning = fake_showwarning


def prod(src: Path, name: str, suffix: str = SUF) -> Path:
    """产物路径 —— v3.0 起就写在原件旁边（和 selftest 里的同名工具一个口径）。"""
    p = Path(name)
    return src / f"{p.stem}{suffix}{p.suffix}"


# 表格列名 → 下标。⚠️ 别写死数字下标：加一列（比如 v3.0 的「所在文件夹」）
# 就会让后面所有下标整体挪位，而写死的数字是静默错位的。
COL: dict[str, int] = {name: i for i, (name, *_rest) in enumerate(A.COLUMNS)}
# 转换模式那套列（两套列共用一棵 Treeview，所以映射也要各记一份）
CCOL: dict[str, int] = {name: i for i, (name, *_rest) in enumerate(A.CONVERT_COLUMNS)}


def set_sources(app, paths: list) -> None:
    """把「原件来源清单」设成给定的这几个（走界面自己的那条路）。"""
    app.sources = [str(p) for p in paths]
    app._refresh_source_list()
    app._invalidate("来源变了")


def set_conv_sources(app, paths: list) -> None:
    """转换模式下换来源 —— 走 `_sources_changed`，队列会跟着重列一遍。

    ⚠️ 跟 `set_sources` 的区别就在这儿：替换模式来源一变只是**作废**，
       要用户再点「预览」；转换模式没有预览那一步，来源一变就得自己列好队。
    """
    app.sources = []
    app._refresh_source_list()
    app._sources_changed("来源变了")
    app._add_sources([str(p) for p in paths])


# ── 假转换引擎：回归**不依赖** Office / LibreOffice ──────────────────────
# 真引擎要拉起 Word / Excel，在 macOS 上会被 TCC 拦（-10004 权限违例）；
# 在没有 Office 的机器上更是压根没有。所以整条「列队 → 执行 → 回填」用假引擎验。
FAKE_ENGINE_KEY = "fake_regress"
FAKE_FAIL_MARK = "BAD"          # 名字里带这个的会被假引擎判失败（验失败分支）


def fake_conv_run(src: Path, dst: Path, fmt) -> str:
    """假引擎：正常情况写一份真 PDF 出去，带 FAKE_FAIL_MARK 的故意报错。"""
    if FAKE_FAIL_MARK in src.name:
        return "假引擎故意失败（回归用）"
    build_pdf(dst, [(595.0, 842.0, f"CONV-{src.stem}"),
                    (595.0, 842.0, f"CONV-{src.stem}-P2")])
    return ""


FAKE_ENGINE = C.ConvEngine(
    FAKE_ENGINE_KEY, "假引擎（回归用）", "只在回归脚本里出现，不碰真的 Office。",
    lambda: True, fake_conv_run,
)


def make_corpus(n_ori: int, n_scan: int, size=(595.0, 842.0)) -> tuple[Path, Path]:
    """建一小批原件 + 一份扫描件。

    ⚠️ v3.0 不再返回「输出目录」—— 产物就落在 `src`（原件目录）里。
    """
    d = TMP / f"corpus_{n_ori}_{n_scan}"
    if d.exists():
        shutil.rmtree(d)
    src = d / "原件"
    src.mkdir(parents=True)
    for i in range(1, n_ori + 1):
        build_pdf(src / f"{i:03d}.pdf",
                  [(size[0], size[1], f"ORI{i}-P1"), (size[0], size[1], f"ORI{i}-P2")])
    scan = d / "扫描件_签字版.pdf"
    build_pdf(scan, [(size[0], size[1], f"SCAN-{k}") for k in range(1, n_scan + 1)])
    return src, scan


def products_of(d: Path, suffix: str = SUF) -> list[str]:
    """那个目录里属于这一轮的产物名（按名称排）。"""
    return sorted(p.name for p in d.glob(f"*{suffix}.pdf"))


def main() -> int:
    global ASK_OVERWRITE          # 覆盖三态：是 / 否 / 取消 三种回答都要走一遍
    root = tk.Tk()
    root.withdraw()
    app = A.App(root)
    app.on_mode("swap")            # [1]~[12] 测的是「首页替换」流程，显式切回 swap 起步

    print("[1] 默认布局")
    labels = [w.label.cget("text") for w in app.rows.values()]
    check("输入框区只剩「扫描件 PDF」一行（原件来源改成清单、输出文件夹也早撤了）",
          labels == ["扫描件 PDF"], str(labels))
    check("默认不倒序", app.vars["reverse"].get() is False)
    check("默认开缩放", app.vars["scale"].get() is True)
    check("vars 里没有 src（原件来源是清单，不是一行输入框）",
          "src" not in app.vars, str(list(app.vars)))
    check("vars 里没有 out（没有输出文件夹这回事了）", "out" not in app.vars, str(list(app.vars)))
    check("vars 里没有 recursive（「连同子文件夹」那个开关删了）",
          "recursive" not in app.vars, str(list(app.vars)))
    check("vars 里没有 overwrite（覆盖改成弹窗，不是勾选框）",
          "overwrite" not in app.vars, str(list(app.vars)))
    check("选项里只剩「扫描页缩放」一项",
          [k for k, _l, _t in A.CHECK_SPECS] == ["scale"],
          str([k for k, _l, _t in A.CHECK_SPECS]))
    check("选项里没有「允许覆盖」勾选框",
          not any(k == "overwrite" for k, _l, _t in A.CHECK_SPECS), str(A.CHECK_SPECS))
    check("选项里没有「连同子文件夹」勾选框",
          not any(k == "recursive" for k, _l, _t in A.CHECK_SPECS), str(A.CHECK_SPECS))
    check("默认两份后缀都是空的",
          app.vars["suffix"].get() == "" and app.vars["conv_suffix"].get() == "",
          f"{app.vars['suffix'].get()!r} / {app.vars['conv_suffix'].get()!r}")
    check("⭐ 后缀变量表：两个模式各一条（不共享记忆，2026-10-08 用户要求）",
          dict(A.SUFFIX_VAR_BY_MODE) == {"swap": "suffix", "convert": "conv_suffix"},
          str(A.SUFFIX_VAR_BY_MODE))
    check("默认「开始替换」是禁用的", str(app.btn_run["state"]) == "disabled")
    check("版本号是 3.0", C.TOOL_VERSION == "3.0", C.TOOL_VERSION)

    print("\n[1a] 原件来源清单：一开始是空的")
    check("清单里一条来源都没有", app.lst_sources.size() == 0, str(app.lst_sources.size()))
    check("地址行是占位文案", app.lbl_srcdir.cget("text") == A.SRC_DIR_EMPTY,
          app.lbl_srcdir.cget("text"))
    check("还没预览过（plan 是空的）", app.plan is None)

    print("\n[1b] 默认窗口放得下 —— 不用最大化也看得全")
    root.update_idletasks()
    need_w, need_h = app.winfo_reqwidth(), app.winfo_reqheight()
    check(f"界面只要 {need_w}px 宽 ≤ 默认窗宽 {A.GEOM_DEFAULT[0]}",
          need_w <= A.GEOM_DEFAULT[0], f"要 {need_w}px")
    check(f"界面只要 {need_h}px 高 ≤ 默认窗高 {A.GEOM_DEFAULT[1]}",
          need_h <= A.GEOM_DEFAULT[1], f"要 {need_h}px")
    check("①②③ 三个模块各占一行（2 列 × 3 行，不是挤在一行）",
          app.mod_hint_row.grid_size() == (2, 3), str(app.mod_hint_row.grid_size()))
    tags = [m.tag for m in A.MODULE_ROWS]
    check("三个模块的编号不重复（①②③ 各一个）",
          len(tags) == 3 and len({t[0] for t in tags}) == 3, str(tags))
    check("界面上不再出现两个「①」（上面两块不编号）",
          not any("①" in w.cget("text") for w in app.winfo_children()
                  if w.winfo_class() == "TLabelframe"))

    print("\n[1c] 后缀必填：闸门一开始就是关着的")
    check("后缀提示写着「不填不让跑」",
          "不填不让跑" in app.lbl_suffix.cget("text"), app.lbl_suffix.cget("text"))
    check("后缀行挂着红字「必填」", app.lbl_suffix_need.cget("text") == "必填",
          app.lbl_suffix_need.cget("text"))
    check("表格上方的地址行还是占位文案",
          app.lbl_srcdir.cget("text") == A.SRC_DIR_EMPTY, app.lbl_srcdir.cget("text"))
    check("core 那边同一个判据也认空后缀不能跑",
          bool(C.suffix_error(C.Options(sources=("x",), scan_pdf="y", suffix=""))))

    print("\n[2] 只列队 → 再配对（v3.0 的两段式）")
    src, scan = make_corpus(3, 3)
    set_sources(app, [str(src)])
    root.update()
    app.on_preview()                      # 故意还没选扫描件
    check("只列队也能预览完成", pump(root, lambda: not app.busy))
    check("没选扫描件也出得了队（不阻断）",
          app.plan is not None and not app.plan.blocked,
          app.plan.block_reason if app.plan else "")
    check("⭐ 这一份还是「只列了队」（paired=False）", app.plan.paired is False)
    check("表格里 3 行都在", len(app.tree.get_children()) == 3,
          str(len(app.tree.get_children())))
    vals = [app.tree.item(r, "values") for r in app.tree.get_children()]
    check("只列队时扫描页列全是占位「—」",
          all(v[COL["scan"]] == "—" for v in vals), str([v[COL["scan"]] for v in vals]))
    check("状态栏提示「还没选扫描件」、让人照着去扫",
          "还没选扫描件" in app.status.cget("text"), app.status.cget("text"))
    check("⭐ 没配对时「开始替换」是禁用的", str(app.btn_run["state"]) == "disabled")
    check("⭐ 没配对时页码/方向按钮也是禁用的（无从调起）",
          all(str(getattr(app, n)["state"]) == "disabled" for n in A.ADJUST_BUTTONS))

    print("\n[2b] 选上扫描件 → 配对")
    app.vars["scan"].set(str(scan))
    root.update()
    app.on_preview()
    check("预览完成", pump(root, lambda: not app.busy))
    check("没被阻断", app.plan is not None and not app.plan.blocked,
          app.plan.block_reason if app.plan else "")
    rows = app.tree.get_children()
    check("表格 3 行", len(rows) == 3, str(len(rows)))
    vals = [app.tree.item(r, "values") for r in rows]
    # Treeview 取回来的值一律是字符串
    check("扫描页列 = 1、2、3",
          [v[COL["scan"]] for v in vals] == ["1", "2", "3"], str([v[COL["scan"]] for v in vals]))
    check("原件列 = 001、002、003",
          [v[COL["src"]] for v in vals] == ["001.pdf", "002.pdf", "003.pdf"], str([v[COL["src"]] for v in vals]))
    check("同层里的原件只显示文件名（不挂一长串目录）",
          all("/" not in v[COL["src"]] for v in vals), str([v[COL["src"]] for v in vals]))
    check("⭐ 后缀没填 → 配对好了也不让跑",
          str(app.btn_run["state"]) == "disabled", str(app.btn_run["state"]))
    check("状态栏点名说「后缀还没填」",
          "后缀" in app.status.cget("text") and "还没填" in app.status.cget("text"),
          app.status.cget("text"))
    check("地址行报出了原件来自哪个文件夹",
          str(src) in app.lbl_srcdir.cget("text"), app.lbl_srcdir.cget("text"))
    check("地址行说清了产物写在原件旁边",
          "产物写在各自原件旁边" in app.lbl_srcdir.cget("text"), app.lbl_srcdir.cget("text"))

    print("\n[2b] 填上后缀 → 闸门开；清空 → 闸门又关上（来回都灵）")
    app.vars["suffix"].set(SUF)
    root.update()
    check("填了后缀，右边实时显示最终名字",
          "001_已签字首页.pdf" in app.lbl_suffix.cget("text"), app.lbl_suffix.cget("text"))
    check("填了后缀 →「开始替换」可点", str(app.btn_run["state"]) == "normal")
    app.vars["suffix"].set("")
    root.update()
    check("清空后缀 → 又变禁用（不是只在建窗时判一次）",
          str(app.btn_run["state"]) == "disabled")
    check("清空后提示又变红字", "不填不让跑" in app.lbl_suffix.cget("text"))
    app.vars["suffix"].set(SUF)
    root.update()
    check("填回来 → 又能点", str(app.btn_run["state"]) == "normal")

    print("\n[2c] ⭐ 选完扫描件 → 不作废旧预览，直接把配对补上（2026-10-08 用户要求）")
    # 用户原话：「我点击扫描件 PDF 之后，它就会卡一下，然后底下的原数据就消失了……
    #            显得像是之前的原件排序感觉有问题一样」。真实流程是
    #            「先加原件 → 预览列队 → 照着扫 → 回来选扫描件」——
    #            这一步只该**把右边那几列补上**，原件那半边一个字都不该动。
    d2c = TMP / "corpus_scan_auto"
    src2c = d2c / "原件"
    src2c.mkdir(parents=True)
    for _i in (1, 2, 3):
        build_pdf(src2c / f"{_i:03d}.pdf",
                  [(595, 842, f"ORI{_i}-P1"), (595, 842, f"ORI{_i}-P2")])
    scan2c = d2c / "扫描件_签字版.pdf"
    build_pdf(scan2c, [(595, 842, f"SCAN-{k}") for k in (1, 2, 3)])

    app.vars["scan"].set("")
    set_sources(app, [str(src2c)])
    app.on_preview()
    check("先只列队（还没选扫描件）", pump(root, lambda: not app.busy))
    check("确实还没配对", app.plan is not None and app.plan.paired is False)
    rows_before = [app.tree.item(r, "values") for r in app.tree.get_children()]
    check("列队后 3 行、扫描页全是占位「—」",
          len(rows_before) == 3 and all(v[COL["scan"]] == "—" for v in rows_before),
          str([v[COL["scan"]] for v in rows_before]))

    app.vars["scan"].set(str(scan2c))       # ← 就是这一步以前会把表格清空
    root.update()
    check("⭐ 刚选完扫描件、防抖还没触发：队列**没被作废**", app.plan is not None,
          str(app.plan))
    check("⭐ 表格也**没被清空**（不是用户说的「底下数据消失了」）",
          len(app.tree.get_children()) == 3, str(len(app.tree.get_children())))
    check("⭐ 状态栏里不该出现「作废」二字（那正是看着别扭的地方）",
          "作废" not in app.status.cget("text"), app.status.cget("text"))

    check("⭐ 防抖到点后**自动**补上配对（不用再点一次「预览」）",
          pump(root, lambda: app.plan is not None and app.plan.paired and not app.busy))
    vals2c = [app.tree.item(r, "values") for r in app.tree.get_children()]
    check("⭐ 还是那 3 行，原件顺序纹丝没动",
          [v[COL["src"]] for v in vals2c] == [v[COL["src"]] for v in rows_before],
          str([v[COL["src"]] for v in vals2c]))
    check("⭐ 右边那几列补上了（扫描页 1 / 2 / 3）",
          [v[COL["scan"]] for v in vals2c] == ["1", "2", "3"],
          str([v[COL["scan"]] for v in vals2c]))
    check("扫描页尺寸 / 方向也跟着填了（不是还挂着「—」）",
          all(v[COL["scansize"]] not in ("", "—") and v[COL["rot"]] != "—" for v in vals2c),
          str([(v[COL["scansize"]], v[COL["rot"]]) for v in vals2c]))
    check("⭐ 状态栏说的是「接上了 / 顺序没变」，不是「作废」",
          ("接上了" in app.status.cget("text") or "顺序没变" in app.status.cget("text"))
          and "作废" not in app.status.cget("text"), app.status.cget("text"))
    check("配对好了 + 后缀已填 →「开始替换」亮着",
          str(app.btn_run["state"]) == "normal", str(app.btn_run["state"]))
    check("防抖句柄收干净（没有挂着的定时器）", app._scan_after is None)

    # 收尾：还原成 [2b] 结束时的样子 —— 后面的 [3]/[4]/[5] 还要用那份语料
    app.vars["scan"].set(str(scan))
    set_sources(app, [str(src)])
    app.on_preview()
    check("（收尾）还原回 3 对 3 的那一批", pump(root, lambda: not app.busy)
          and app.plan is not None and app.plan.paired)

    print("\n[3] 模块①：页序按钮（正序 ⇄ 倒序，一键来回切）")
    check("② 选项里已经没有「倒序」勾选框了",
          not any(k == "reverse" for k, _l, _t in A.CHECK_SPECS))
    check("按钮初始显示「页序：正序」", app.btn_order.cget("text") == "页序：正序",
          app.btn_order.cget("text"))
    app.on_flip_order()
    root.update()
    vals = [app.tree.item(r, "values") for r in app.tree.get_children()]
    check("点一下 → 扫描页列 = 3、2、1",
          [v[COL["scan"]] for v in vals] == ["3", "2", "1"], str([v[COL["scan"]] for v in vals]))
    check("按钮文字跟着变成「页序：倒序」", app.btn_order.cget("text") == "页序：倒序",
          app.btn_order.cget("text"))
    app.on_flip_order()
    root.update()
    vals = [app.tree.item(r, "values") for r in app.tree.get_children()]
    check("再点一下回到正序（来回点是幂等的）",
          [v[COL["scan"]] for v in vals] == ["1", "2", "3"], str([v[COL["scan"]] for v in vals]))
    check("按钮文字回到「页序：正序」", app.btn_order.cget("text") == "页序：正序")

    print("\n[4] 模块②：页码对应（行不动，只换用哪一页）")
    app.tree.selection_set(app.tree.get_children()[0])
    app._sync_adjust_buttons()
    check("选中行后页码按钮可用", str(app.btn_up["state"]) == "normal")
    app._swap_page(1)
    vals = [app.tree.item(r, "values") for r in app.tree.get_children()]
    check("对调后原件顺序一动不动（名称序是唯一的排队依据）",
          [v[COL["src"]] for v in vals] == ["001.pdf", "002.pdf", "003.pdf"], str([v[COL["src"]] for v in vals]))
    check("001 用第 2 页、002 用第 1 页",
          [v[COL["scan"]] for v in vals] == ["2", "1", "3"], str([v[COL["scan"]] for v in vals]))
    check("一页不会被两行同时用", len({v[COL["scan"]] for v in vals}) == 3, str([v[COL["scan"]] for v in vals]))
    app._swap_page(-1)          # 选中行已跟着走到第 2 行，往上换回去
    vals = [app.tree.item(r, "values") for r in app.tree.get_children()]
    check("再换回来", [v[COL["scan"]] for v in vals] == ["1", "2", "3"], str([v[COL["scan"]] for v in vals]))

    # 直接指定页码：占用别人的页 → 自动对调
    app.tree.selection_set(app.tree.get_children()[3 - 1])
    changed, note = C.assign_page(app.plan, 3, 1)
    app._apply_rotations()
    app._reselect_refresh()
    vals = [app.tree.item(r, "values") for r in app.tree.get_children()]
    check("把第 3 行改成用第 1 页 → 第 1 行自动换成第 3 页",
          [v[COL["scan"]] for v in vals] == ["3", "2", "1"], str([v[COL["scan"]] for v in vals]))
    check("对调说明里点名了被换的那份原件", "001.pdf" in note, note)
    C.assign_page(app.plan, 3, 3)
    app._apply_rotations()
    app._reselect_refresh()

    print("\n[5] 执行替换（产物落在原件旁边）")
    before = {p.name: page_texts(p) for p in src.glob("*.pdf")}
    app.on_run()
    check("替换完成", pump(root, lambda: not app.busy))
    outs = products_of(src)
    check("产出 3 份，就在原件旁边", len(outs) == 3, str(outs))
    check("产物名 = 原件名 + 后缀",
          outs == [f"00{i}{SUF}.pdf" for i in (1, 2, 3)], str(outs))
    texts = page_texts(prod(src, "002.pdf"))
    check("002 的产物第 1 页是扫描件第 2 页", texts and "SCAN-2" in texts[0], str(texts[:1]))
    check("002 的产物第 2 页还是原件内容", texts[1:2] == ["ORI2-P2"], str(texts))
    check("⭐ 原件一个字节都没动",
          {n: page_texts(src / n) for n in before} == before)
    vals = [app.tree.item(r, "values") for r in app.tree.get_children()]
    check("表格状态列刷新成「替换成功」",
          all(v[COL["status"]] == "替换成功" for v in vals), str([v[COL["status"]] for v in vals]))
    check("状态栏说清产物落在哪", "产物写在各自原件旁边" in app.status.cget("text"),
          app.status.cget("text"))

    print("\n[5b] 覆盖确认弹窗（v3.0：没有勾选框，得点一次头；只报份数不列文件名）")
    d5b = TMP / "corpus_overwrite"
    src5b = d5b / "原件"
    src5b.mkdir(parents=True)
    for i in (1, 2):
        build_pdf(src5b / f"{i:03d}.pdf", [(595, 842, f"ORI{i}-P1"), (595, 842, f"ORI{i}-P2")])
    scan5b = d5b / "扫描件.pdf"
    build_pdf(scan5b, [(595, 842, "SCAN-1"), (595, 842, "SCAN-2")])
    set_sources(app, [str(src5b)])
    app.vars["scan"].set(str(scan5b))
    app.vars["suffix"].set(SUF)
    app.on_preview()
    check("预览完成", pump(root, lambda: not app.busy))
    app.on_run()                                   # 第一次：没有同名 → 走普通确认
    check("第一次替换完成", pump(root, lambda: not app.busy))
    first = products_of(src5b)
    check("第一次产出 2 份", len(first) == 2, str(first))
    raw_first = {n: (src5b / n).read_bytes() for n in first}

    # ── 取消：什么都不能动
    global ASK_OVERWRITE
    ASKED.clear()
    ASK_OVERWRITE = None
    app.on_preview()
    check("重跑预览完成", pump(root, lambda: not app.busy))
    check("重跑预览只有 2 份原件（产物被认出来、没当成原件）",
          len(app.plan.items) == 2, str(len(app.plan.items)))
    app.on_run()
    check("弹出了覆盖确认窗", len(ASKED) == 1, str(len(ASKED)))
    check("弹窗只报重名份数（不列文件名清单）",
          "2 份" in ASKED[0] and "001_已签字首页.pdf" not in ASKED[0],
          ASKED[0][:160] if ASKED else "")
    check("弹窗说明原件不会被改动", "原件" in ASKED[0] and "不会被改动" in ASKED[0])
    # 2026-10-09 用户要求：三个选项**各占一行**、一句一个，别堆成一段。
    check("三个选项都写清楚了（是 ＝ 覆盖 / 否 ＝ 加编号另存 / 取消 ＝ 先别跑）",
          "是 ＝ 覆盖它" in ASKED[0] and "否 ＝ 加编号另存" in ASKED[0]
          and "取消 ＝ 先别跑" in ASKED[0], ASKED[0][:200] if ASKED else "")
    check("点「取消」→ 一个字节都没动",
          {n: (src5b / n).read_bytes() for n in first} == raw_first)
    check("点「取消」→ 状态栏说明已取消", "已取消" in app.status.cget("text"),
          app.status.cget("text"))
    check("点「取消」→ 没有多出任何文件", products_of(src5b) == first,
          str(products_of(src5b)))

    # ── 否：不覆盖，加编号让开
    ASK_OVERWRITE = False
    app.on_run()
    check("重跑完成（选「否」）", pump(root, lambda: not app.busy))
    check("选「否」→ 产物加编号让开",
          (src5b / f"001{SUF}_2.pdf").is_file(), str(products_of(src5b)))
    check("选「否」→ 原来那份没被动",
          {n: (src5b / n).read_bytes() for n in first} == raw_first)
    for extra in src5b.glob(f"*{SUF}_2.pdf"):     # 清掉编号件，回到干净状态
        extra.unlink()

    # ── 是：原地覆盖（重跑＝更新）
    ASK_OVERWRITE = True
    app.on_run()
    check("重跑完成（选「是」）", pump(root, lambda: not app.busy))
    check("选「是」→ 原地覆盖，不会再生出 _3",
          not list(src5b.glob(f"*{SUF}_3.pdf")), str(products_of(src5b)))
    check("选「是」→ 产物仍是 2 份（更新而不是堆积）",
          products_of(src5b) == first, str(products_of(src5b)))
    check("选「是」→ 原件依旧没被动",
          {n: page_texts(src5b / n) for n in ("001.pdf", "002.pdf")}["001.pdf"][1:]
          == ["ORI1-P2"], "")
    ASK_OVERWRITE = True

    print("\n[6] 数量对不上 → 整批拒绝")
    src2, scan2 = make_corpus(5, 3)
    set_sources(app, [str(src2)])
    app.vars["scan"].set(str(scan2))
    WARNED.clear()
    app.on_preview()
    check("预览完成", pump(root, lambda: not app.busy))
    check("被阻断", app.plan is not None and app.plan.blocked)
    check("「开始替换」仍是禁用的", str(app.btn_run["state"]) == "disabled")
    check("弹了警告说清楚差几份",
          any("差 2" in w for w in WARNED), str(WARNED))
    app.on_run()
    check("阻断时点「开始替换」不会产出任何文件",
          not products_of(src2), str(products_of(src2)))

    print("\n[7] 右键：忽略一行")
    src3, scan3 = make_corpus(3, 3)
    set_sources(app, [str(src3)])
    app.vars["scan"].set(str(scan3))
    app.on_preview()
    check("预览完成", pump(root, lambda: not app.busy))
    app.tree.selection_set(app.tree.get_children()[1])
    app.on_ignore()
    vals = [app.tree.item(r, "values") for r in app.tree.get_children()]
    check("第 2 行变成「已跳过」", vals[1][COL["status"]] == "已跳过", str(vals[1]))
    app.on_run()
    check("替换完成", pump(root, lambda: not app.busy))
    outs3 = products_of(src3)
    check("被忽略的那份没有输出",
          outs3 == [f"001{SUF}.pdf", f"003{SUF}.pdf"], str(outs3))

    print("\n[8] 后缀实时预览")
    src4, scan4 = make_corpus(2, 2)
    set_sources(app, [str(src4)])
    app.vars["scan"].set(str(scan4))
    app.on_preview()
    check("预览完成", pump(root, lambda: not app.busy))
    app.vars["suffix"].set("_已签字首页")
    root.update()
    check("后缀预览显示真实文件名", "001_已签字首页.pdf" in app.lbl_suffix.cget("text"),
          app.lbl_suffix.cget("text"))
    app.vars["suffix"].set("_a/b")
    root.update()
    check("非法字符被去掉", "001_ab.pdf" in app.lbl_suffix.cget("text"),
          app.lbl_suffix.cget("text"))

    print("\n[9] 不改名：按 WI-xxxx 原名排队（2026-10-08 用户拍板）")
    d9 = TMP / "corpus_named"
    src9 = d9 / "原件"
    src9.mkdir(parents=True)
    for nm in ("WI-3852", "WI-1032", "WI-2001"):      # 故意乱序建
        build_pdf(src9 / f"{nm}.pdf",
                  [(595, 842, f"{nm}-P1"), (595, 842, f"{nm}-P2")])
    scan9 = d9 / "扫描件_签字版.pdf"
    build_pdf(scan9, [(595, 842, f"SCAN-{k}") for k in range(1, 4)])
    set_sources(app, [str(src9)])
    app.vars["scan"].set(str(scan9))
    app.vars["suffix"].set(SUF)
    app.on_preview()
    check("预览完成", pump(root, lambda: not app.busy))
    check("没被阻断（3 份对 3 页）", app.plan is not None and not app.plan.blocked)
    vals = [app.tree.item(r, "values") for r in app.tree.get_children()]
    got9 = [v[COL["src"]] for v in vals]
    check("表格顺序 = 文件夹里按名称看到的顺序（1032 → 2001 → 3852）",
          got9 == ["WI-1032.pdf", "WI-2001.pdf", "WI-3852.pdf"], str(got9))
    check("扫描页仍是 1、2、3",
          [v[COL["scan"]] for v in vals] == ["1", "2", "3"], str([v[COL["scan"]] for v in vals]))
    app.on_run()
    check("替换完成", pump(root, lambda: not app.busy))
    t9 = page_texts(prod(src9, "WI-2001.pdf"))
    check("WI-2001 拿到扫描件第 2 页", bool(t9) and "SCAN-2" in t9[0], str(t9[:1]))
    check("WI-2001 第 2 页还是原件内容", t9[1:2] == ["WI-2001-P2"], str(t9))
    check("WI 原名下产物名 = 原名 + 后缀",
          sorted(p.name for p in src9.glob(f"*{SUF}.pdf"))
          == [f"WI-{n}{SUF}.pdf" for n in (1032, 2001, 3852)],
          str(sorted(p.name for p in src9.iterdir())))

    print("\n[10] 文案与按钮口径（UI 精简后）")
    check("输入框表里只剩「扫描件 PDF」这一行（原件来源改清单、输出文件夹早撤了）",
          [s.key for s in A.FOLDER_SPECS] == ["scan"], str([s.key for s in A.FOLDER_SPECS]))
    check("扫描件那行的提示够短（≤ 20 字，不再堆一整段）",
          len(A.FOLDER_SPECS[0].hint) <= 20, A.FOLDER_SPECS[0].hint)
    check("长说明收进了悬浮提示（detail）",
          "可以先不选" in A.FOLDER_SPECS[0].detail)
    check("长说明点明「份数必须一致、不一致整批不执行」",
          "份数必须" in A.FOLDER_SPECS[0].detail and "整批不执行" in A.FOLDER_SPECS[0].detail)
    check("来源清单那块有自己的短提示（讲清「按文件名排」）",
          "按文件名" in A.SRC_HINT_BY_MODE["swap"], A.SRC_HINT_BY_MODE["swap"])
    check("来源清单的悬浮提示写了「一次选多个」",
          "一次选多个" in A.SRC_TIP_FILES and "Cmd" in A.SRC_TIP_FILES, A.SRC_TIP_FILES)
    check("⭐ 来源清单的按钮就是表里那三个：加文件 / 移除 / 清空",
          [b.text for b in A.SRC_BUTTONS] == ["＋ 添加文件…", "－ 移除选中", "清空"],
          str([b.text for b in A.SRC_BUTTONS]))
    check("⭐ 表里每个按钮的方法都真的存在（表驱动不会指丢）",
          all(hasattr(A.App, b.method) for b in A.SRC_BUTTONS),
          str([b.method for b in A.SRC_BUTTONS]))
    check("⭐ 「加文件夹」入口彻底没了（方法也删掉了）",
          not hasattr(A.App, "on_add_folder")
          and not hasattr(A.App, "_pick_many_folders"))
    check("⭐ 两块区的标题都写明「不记住」（不记来源、也不记扫描件）",
          all("不记住" in t for t in A.SRC_BOX_TITLE_BY_MODE.values())
          and "不记住" in A.SCAN_BOX_TITLE,
          f"{A.SRC_BOX_TITLE_BY_MODE} / {A.SCAN_BOX_TITLE}")
    check("⭐ 来源框的标题也按模式换（一个收 PDF、一个收 Word / Excel）",
          set(A.SRC_BOX_TITLE_BY_MODE) == {"swap", "convert"}
          and "PDF" in A.SRC_BOX_TITLE_BY_MODE["swap"]
          and "Word" in A.SRC_BOX_TITLE_BY_MODE["convert"],
          str(A.SRC_BOX_TITLE_BY_MODE))
    check("⭐ 路径一个都不落盘：白名单里没有 scan，也没有 sources",
          "scan" not in A.CONFIG_PERSIST_VARS
          and "sources" not in A.CONFIG_PERSIST_VARS,
          str(A.CONFIG_PERSIST_VARS))
    check("后缀示例换成真实 WI 名", "WI-1032" in A.SUFFIX_SAMPLE_FALLBACK,
          A.SUFFIX_SAMPLE_FALLBACK)
    check("⭐ 后缀悬浮说明按模式**各一份**（表驱动，不再是一段里讲两件事）",
          set(A.SUFFIX_TIPS_BY_MODE) == {"swap", "convert"}, str(sorted(A.SUFFIX_TIPS_BY_MODE)))
    check("替换那份讲「必填」、转换那份讲「可以不填」",
          "必填" in A.SUFFIX_TIPS_BY_MODE["swap"]
          and "可以不填" in A.SUFFIX_TIPS_BY_MODE["convert"],
          A.SUFFIX_TIPS_BY_MODE["swap"][:40])
    check("⭐ 两份说明都点明「各记各的、互不影响」（这是用户最在意的点）",
          all("互不影响" in t for t in A.SUFFIX_TIPS_BY_MODE.values()),
          str([t[-16:] for t in A.SUFFIX_TIPS_BY_MODE.values()]))
    check("⭐ 「必填 / 可不填」那个小标签也按模式各一份",
          "必填" in A.SUFFIX_NEED_TIPS_BY_MODE["swap"]
          and "可不填" in A.SUFFIX_NEED_TIPS_BY_MODE["convert"],
          str(A.SUFFIX_NEED_TIPS_BY_MODE))
    check("预览按钮的悬浮提示讲清了「先列队、再配对」",
          "还没选扫描件" in A.PREVIEW_TIP and "选好扫描件之后" in A.PREVIEW_TIP,
          A.PREVIEW_TIP)
    for gone, why in (("on_open_out", "打开输出文件夹"), ("on_forget", "忘记记住的文件夹"),
                      ("on_help", "使用说明按钮"), ("_autofill_out", "自动补输出目录")):
        check(f"删掉了「{why}」", not hasattr(A.App, gone))
    check("不再留 HELP_TEXT 死常量", not hasattr(A, "HELP_TEXT"))
    check("①②③ 的按钮都建好了（ADJUST_BUTTONS 全在）",
          all(hasattr(app, n) for n in A.ADJUST_BUTTONS), str(A.ADJUST_BUTTONS))
    check("方向按钮都是「整批」口径（有全部 / 横向两组）",
          "全部左转" in app.btn_rot_all_l.cget("text")
          and "横向右转" in app.btn_rot_land_r.cget("text"))

    print("\n[11] 模块③：方向（默认规则 + 整批旋转）")
    d11 = TMP / "corpus_rot"
    src11 = d11 / "原件"
    src11.mkdir(parents=True)
    # 第 2 份原件是横向的（842×595），其余竖向；扫描件一律竖扫（595×842）
    for i, sz in ((1, (595, 842)), (2, (842, 595)), (3, (595, 842))):
        build_pdf(src11 / f"{i:03d}.pdf",
                  [(sz[0], sz[1], f"ORI{i}-P1"), (sz[0], sz[1], f"ORI{i}-P2")])
    scan11 = d11 / "扫描件.pdf"
    build_pdf(scan11, [(595, 842, f"SCAN-{k}") for k in range(1, 4)])
    set_sources(app, [str(src11)])
    app.vars["scan"].set(str(scan11))
    app.on_preview()
    check("预览完成", pump(root, lambda: not app.busy))
    vals = [app.tree.item(r, "values") for r in app.tree.get_children()]
    check("默认规则：横向原件那一页自动「左转 90°」，其余正向",
          [v[COL["rot"]] for v in vals] == ["正向", "左转 90°", "正向"], str([v[COL["rot"]] for v in vals]))
    check("那一页的扫描页尺寸已经宽高互换",
          [v[COL["scansize"]] for v in vals] == ["595×842", "842×595", "595×842"],
          str([v[COL["scansize"]] for v in vals]))
    check("预览后方向按钮可用", str(app.btn_rot_all_r["state"]) == "normal")

    # 整批右转：所有页一起转（横向那页 270→0，其余 0→90）
    app.on_turn_all(C.ROT_RIGHT)
    vals = [app.tree.item(r, "values") for r in app.tree.get_children()]
    check("「全部右转」所有页都动了",
          [v[COL["rot"]] for v in vals] == ["右转 90°", "正向", "右转 90°"], str([v[COL["rot"]] for v in vals]))
    app.on_turn_all(C.ROT_LEFT)
    vals = [app.tree.item(r, "values") for r in app.tree.get_children()]
    check("「全部左转」转回默认的样子",
          [v[COL["rot"]] for v in vals] == ["正向", "左转 90°", "正向"], str([v[COL["rot"]] for v in vals]))

    # 横向专用：只动「配给横向原件」的那几页
    app.on_turn_land(C.ROT_LEFT)
    vals = [app.tree.item(r, "values") for r in app.tree.get_children()]
    check("「横向左转」只影响横向原件那一行",
          [v[COL["rot"]] for v in vals] == ["正向", "倒转 180°", "正向"], str([v[COL["rot"]] for v in vals]))
    check("状态栏说明了转了几页", "横向 1 页" in app.status.cget("text"),
          app.status.cget("text"))
    app.on_reset_rot()
    vals = [app.tree.item(r, "values") for r in app.tree.get_children()]
    check("「方向复位」回到默认规则（不是全部清零）",
          [v[COL["rot"]] for v in vals] == ["正向", "左转 90°", "正向"], str([v[COL["rot"]] for v in vals]))

    # 页码一动，「这一页是谁的封面」变了 → 默认的左转跟着横向原件走
    app.tree.selection_set(app.tree.get_children()[1])      # 第 2 行 = 横向原件 002
    app._swap_page(-1)                                      # 与第 1 行对调页码
    vals = [app.tree.item(r, "values") for r in app.tree.get_children()]
    check("对调后原件顺序没变（行不动）",
          [v[COL["src"]] for v in vals] == ["001.pdf", "002.pdf", "003.pdf"], str([v[COL["src"]] for v in vals]))
    check("对调后页码换成了 2、1、3",
          [v[COL["scan"]] for v in vals] == ["2", "1", "3"], str([v[COL["scan"]] for v in vals]))
    check("横向原件那一行仍是左转（跟着原件走，不跟着页码走）",
          [v[COL["rot"]] for v in vals] == ["正向", "左转 90°", "正向"], str([v[COL["rot"]] for v in vals]))
    check("于是「要转的那页」换成了第 1 页",
          app.plan.items[1].scan_index == 1 and app.plan.items[1].rotation == 270,
          f"{app.plan.items[1].scan_index}/{app.plan.items[1].rotation}")

    app.on_run()
    check("替换完成", pump(root, lambda: not app.busy))
    check("横向原件那份：产物首页真的横过来了 842×595",
          page_sizes(prod(src11, "002.pdf"))[0] == (842.0, 595.0),
          str(page_sizes(prod(src11, "002.pdf"))))
    check("竖向原件那份：产物首页仍是 595×842（没被连坐）",
          page_sizes(prod(src11, "001.pdf"))[0] == (595.0, 842.0),
          str(page_sizes(prod(src11, "001.pdf"))))

    print("\n[12] 多来源：来源清单 + 产物各回各家")
    d12 = TMP / "corpus_multi"
    src12 = d12 / "原件"
    (src12 / "批次B").mkdir(parents=True)
    build_pdf(src12 / "001.pdf", [(595, 842, "A-ORI1-P1"), (595, 842, "A-ORI1-P2")])
    build_pdf(src12 / "002.pdf", [(595, 842, "A-ORI2-P1"), (595, 842, "A-ORI2-P2")])
    build_pdf(src12 / "批次B" / "003.pdf", [(595, 842, "B-ORI3-P1"), (595, 842, "B-ORI3-P2")])
    build_pdf(src12 / "批次B" / "004.pdf", [(595, 842, "B-ORI4-P1"), (595, 842, "B-ORI4-P2")])
    scan12 = d12 / "扫描件.pdf"
    build_pdf(scan12, [(595, 842, f"SCAN-{k}") for k in range(1, 5)])

    # 只加顶层文件夹 → 子目录里的原件不算数 → 2 份对 4 页，被拦住
    set_sources(app, [src12])
    app.vars["scan"].set(str(scan12))
    app.on_preview()
    check("预览完成（只加顶层文件夹）", pump(root, lambda: not app.busy))
    check("子目录里的原件不算数（只收 2 份 → 跟 4 页对不上，被拦住）",
          app.plan is not None and app.plan.blocked,
          f"items={len(app.plan.items) if app.plan else '?'}")

    # 把子目录也**作为来源**加进来 → 4 份；来源顺序不影响队列（按文件名统一排）
    set_sources(app, [src12 / "批次B", src12])
    app.on_preview()
    check("预览完成（两个来源）", pump(root, lambda: not app.busy))
    check("4 份都收进来了", app.plan is not None and len(app.plan.items) == 4,
          str(len(app.plan.items) if app.plan else "?"))
    check("⭐ 清单两条来源 → 地址行报「2 个来源」",
          "2 个来源" in app.lbl_srcdir.cget("text"), app.lbl_srcdir.cget("text"))
    vals = [app.tree.item(r, "values") for r in app.tree.get_children()]
    check("⭐ 队列按文件名统一排（001..004，跟来源的先后无关）",
          [v[COL["src"]] for v in vals] == ["001.pdf", "002.pdf", "003.pdf", "004.pdf"],
          str([v[COL["src"]] for v in vals]))
    check("⭐ 跨文件夹时「所在文件夹」列填上了",
          [v[COL["dir"]] for v in vals] == ["原件", "原件", "批次B", "批次B"],
          str([v[COL["dir"]] for v in vals]))

    app.vars["suffix"].set(SUF)
    app.on_run()
    check("替换完成（多来源）", pump(root, lambda: not app.busy))
    check("顶层原件的产物落在顶层", len(products_of(src12)) == 2, str(products_of(src12)))
    check("子目录原件的产物落在那个子目录里",
          len(products_of(src12 / "批次B")) == 2, str(products_of(src12 / "批次B")))
    check("没有凭空冒出一个输出文件夹",
          sorted(p.name for p in src12.iterdir() if p.is_dir()) == ["批次B"],
          str(sorted(p.name for p in src12.iterdir())))

    print("\n[12b] 改了来源 → 旧队作废（不能拿去接着调）")
    # 界面上现在只能「加文件」，这里就按界面那条路加一份 PDF（core 仍认文件夹来源，
    # [12] 那段守的就是那个能力）
    app._add_sources([str(src12 / "001.pdf")])
    check("来源加进清单了", len(app.sources) == 3, str(app.sources))
    check("⭐ 旧的配对结果被清掉（plan=None、表格也空）",
          app.plan is None and not app.tree.get_children())
    check("状态栏说清「预览作废、重新来一遍」",
          "作废" in app.status.cget("text"), app.status.cget("text"))
    check("「开始替换」跟着变回禁用", str(app.btn_run["state"]) == "disabled")

    print("\n[12c] 移除 / 清空来源")
    app.lst_sources.selection_set(2)
    app.on_remove_sources()
    check("移除选中那条（清掉最后加的那份文件）",
          app.sources == [str(src12 / "批次B"), str(src12)], str(app.sources))
    app.on_clear_sources()
    check("清空后一条不剩", app.sources == [], str(app.sources))
    check("地址行回到占位文案", app.lbl_srcdir.cget("text") == A.SRC_DIR_EMPTY,
          app.lbl_srcdir.cget("text"))

    # ══════════════════════════════════════════════════════════════════════
    #  ⭐ v3.0 的第二套界面：Office 转 PDF（顶部那两个大按钮切过来）
    #     它跟首页替换共用来源清单 / 后缀 / 覆盖三态，但流程短一截：
    #     加文件 → 自动列队 → 点「开始转换」。没有扫描件、没有①②③、没有「预览」。
    # ══════════════════════════════════════════════════════════════════════
    print("\n[13] 双模式：顶部两个大按钮切界面")
    check("模式表就是两张（导出 / 替换，导出在前）",
          [m.key for m in A.MODE_SPECS] == ["convert", "swap"],
          str([m.key for m in A.MODE_SPECS]))
    check("默认模式常量 = 导出 PDF（先转 PDF，替换是后面一步）",
          A.DEFAULT_MODE == "convert", A.DEFAULT_MODE)
    check("两个模式按钮都建好了", set(app._mode_buttons) == {"swap", "convert"},
          str(sorted(app._mode_buttons)))
    check("⭐ 当前模式靠字号加粗区分（两个按钮的 style 不一样）",
          app._mode_buttons["swap"].cget("style") == "ModeOn.TButton"
          and app._mode_buttons["convert"].cget("style") == "ModeOff.TButton",
          f"{app._mode_buttons['swap'].cget('style')} / "
          f"{app._mode_buttons['convert'].cget('style')}")

    app.on_mode("convert")
    root.update()
    check("切到转换模式", app.mode == "convert", app.mode)
    check("表格换成转换那套列（8 列，没有扫描页 / 方向）",
          tuple(app.tree["columns"]) == tuple(c[0] for c in A.CONVERT_COLUMNS),
          str(tuple(app.tree["columns"])))
    check("扫描件那一行收起来了（转换用不到扫描件）",
          app.rows["scan"].frame.master.winfo_manager() == "")
    check("①②③ 三个模块整块收起来（没有页可调）", app.mod_hint_row.winfo_manager() == "")
    check("「预览」与「开始替换」都收起来",
          app.btn_scan.winfo_manager() == "" and app.btn_run.winfo_manager() == "")
    check("「开始转换」和「重新检测引擎」露出来",
          app.btn_convert.winfo_manager() != "" and app.btn_recheck.winfo_manager() != "")
    check("引擎那只小标签也在", app.lbl_engine.winfo_manager() != "")
    check("「缩放」那一行收起来（转换走文件自带的打印设置）",
          app._opt_rows["scale_row"].winfo_manager() == "")
    check("后缀那一行留着（**输入框共用一格**，背后是转换自己那个变量）",
          app._opt_rows["suffix_row"].winfo_manager() != "")
    check("⭐ 后缀标签从「必填」变成「可不填」", app.lbl_suffix_need.cget("text") == "可不填",
          app.lbl_suffix_need.cget("text"))
    check("⭐ 那一格已切到**转换模式自己那份**变量（不跟替换共用记忆）",
          app.ent_suffix.get() == app.vars["conv_suffix"].get(),
          f"{app.ent_suffix.get()!r} vs {app.vars['conv_suffix'].get()!r}")
    check("⭐ 替换那份原封不动（SUF 还挂在 suffix 上）",
          app.vars["suffix"].get() == SUF, repr(app.vars["suffix"].get()))
    check("悬浮说明也换成转换那一份（讲「可以不填」+ 互不影响）",
          "可以不填" in app._tip_suffix.text and "互不影响" in app._tip_suffix.text,
          app._tip_suffix.text[:48])
    check("⭐ 来源说明也跟着换（不再劝人挑 PDF）",
          "Word / Excel" in app._src_tips["on_add_files"].text,
          app._src_tips["on_add_files"].text[:40])

    # ⭐ 2026-10-09：引擎不再只有「一个全局首选」—— 一族一条（见 core 的 ⑩a）。
    #    这台机器上能转的东西本来就不止一样：Windows 上 Word/Excel 走 Office、
    #    CAD 走本机的浩辰 CAD、BarTender 走 BarTender，谁也不是谁的备胎。
    _probed = A.App._probe_engines()
    check("⭐ 引擎是**一族一条**挑出来的（不是只有一个全局首选）",
          all(e.family in C.ENGINE_FAMILIES for e in _probed)
          and len({e.family for e in _probed}) == len(_probed),
          str([(e.key, e.family) for e in _probed]))
    if sys.platform == "darwin":
        check("⭐⭐ Mac 上**挑不出** CAD / BarTender 那两族的引擎"
              "（所以列队时会被明确拦下，而不是让某个应用白试一遍）",
              not [e for e in _probed if e.family in ("cad", "btw")],
              str([e.key for e in _probed]))
    # ⭐ 2026-10-10：CAD 的「注册表里有没有这个名字」在 core 里是**进程级缓存**的，
    #    「重新检测引擎」若不清它，刚装上的浩辰 CAD 照样看不见 —— 这条盯着它。
    _rc_calls: list[int] = []
    _orig_reset = C.reset_cad_cache
    C.reset_cad_cache = lambda: _rc_calls.append(1)
    try:
        app.busy = False
        app.on_recheck_engine()
    finally:
        C.reset_cad_cache = _orig_reset
    check("⭐⭐ 点「重新检测引擎」会清掉 CAD 的探测缓存"
          "（不清的话，工具开着时刚装上的 CAD 永远看不见）",
          len(_rc_calls) == 1, str(_rc_calls))

    # 装一个假引擎：真引擎要么被 macOS 的 TCC 拦（-10004 权限违例）、要么这台没装，
    # 整条「列队 → 执行 → 回填」就靠它验（补到表头 + pick_engine 直接返回它）。
    # ⭐ 2026-10-09：pick_engine 现在带参数（要问「哪一族」，见 core 的 ⑩a），
    #    所以替身必须收得下 *args/**kwargs —— 写死 0 参会被 TypeError 打成「列队出错」。
    app.engines = None
    C.CONVERT_ENGINES = (FAKE_ENGINE,) + tuple(
        e for e in C.CONVERT_ENGINES if e.key != FAKE_ENGINE_KEY)
    C.pick_engine = lambda *_a, **_k: FAKE_ENGINE
    app._refresh_engine_label()
    check("引擎标签报出用的哪个引擎", "假引擎" in app.lbl_engine.cget("text"),
          app.lbl_engine.cget("text"))

    print("\n[13b] 转换模式：来源一加进来就列好队（没有「预览」那一步）")
    d13 = TMP / "corpus_conv"
    d13.mkdir(parents=True, exist_ok=True)
    for name in ("WI-1001.xlsx", "WI-1002.docx", "WI-1003.xls"):
        (d13 / name).write_bytes(b"fake office bytes")
    (d13 / "notes.txt").write_text("x", encoding="utf-8")
    # ⭐ 转换模式：后缀**可以不填** —— 而且这跟替换那边那份是**两个变量**，
    #    哪怕替换那边挂着 SUF，转换这边照样是空的（不共享记忆）。
    check("⭐ 转换那份后缀现在是空的（可以不填）",
          app.vars["conv_suffix"].get() == "", repr(app.vars["conv_suffix"].get()))
    check("⚠️ 替换那份没被牵连（还挂着 SUF）",
          app.vars["suffix"].get() == SUF, repr(app.vars["suffix"].get()))
    set_conv_sources(app, [d13])
    check("文件夹来源只收 Word / Excel（txt / pdf 一律不算）",
          len(app.conv_plan.items) == 3 and not app.conv_plan.unknown,
          f"items={len(app.conv_plan.items)} unknown={len(app.conv_plan.unknown)}")
    check("⭐ 加完来源表格就自己列好了（不用点任何按钮）",
          len(app.tree.get_children()) == 3, str(len(app.tree.get_children())))
    vals = [app.tree.item(r, "values") for r in app.tree.get_children()]
    check("队列按文件名排（跟添加顺序无关）",
          [v[CCOL["src"]] for v in vals] == ["WI-1001.xlsx", "WI-1002.docx", "WI-1003.xls"],
          str([v[CCOL["src"]] for v in vals]))
    check("「类型」列分得出 Word / Excel",
          [v[CCOL["kind"]] for v in vals] == ["Excel", "Word", "Excel"],
          str([v[CCOL["kind"]] for v in vals]))
    check("⭐ 后缀不填也能算出产物名（.xlsx → .pdf，名字根本不会撞）",
          [v[CCOL["out"]] for v in vals] == ["WI-1001.pdf", "WI-1002.pdf", "WI-1003.pdf"],
          str([v[CCOL["out"]] for v in vals]))
    check("「产物名」列只显示短名（不带整条路径）",
          all("/" not in str(v[CCOL["out"]]) for v in vals),
          str([v[CCOL["out"]] for v in vals]))
    check("转之前「产物页数」一律是空的（那会儿还是个 Excel）",
          all(str(v[CCOL["pages"]]) == "—" for v in vals),
          str([v[CCOL["pages"]] for v in vals]))
    check("⭐ 「开始转换」在能转的时候是亮着的",
          str(app.btn_convert["state"]) == "normal", str(app.btn_convert["state"]))

    app._add_sources([str(d13 / "notes.txt")])
    check("⭐ 单文件来源里认不出格式的只统计、不进队",
          len(app.conv_plan.items) == 3 and len(app.conv_plan.unknown) == 1,
          f"items={len(app.conv_plan.items)} unknown={len(app.conv_plan.unknown)}")

    app.vars["conv_suffix"].set("_已转")     # ⭐ 动的是**转换自己那份**
    root.update()
    check("⭐ 替换那份后缀仍然没被动过",
          app.vars["suffix"].get() == SUF, repr(app.vars["suffix"].get()))
    vals = [app.tree.item(r, "values") for r in app.tree.get_children()]
    check("⭐ 填了后缀，产物名立刻跟上",
          [v[CCOL["out"]] for v in vals]
          == ["WI-1001_已转.pdf", "WI-1002_已转.pdf", "WI-1003_已转.pdf"],
          str([v[CCOL["out"]] for v in vals]))

    print("\n[13c] 真跑一遍（假引擎替身，不碰真 Office）")
    before = {p.name: (p.stat().st_mtime_ns, p.read_bytes())
              for p in d13.iterdir() if p.is_file()}
    asked_before, warned_before = len(ASKED), len(WARNED)
    app.on_convert()
    check("转换完成", pump(root, lambda: not app.busy))
    check("没有重名 → 不该弹覆盖窗", len(ASKED) == asked_before, str(ASKED[asked_before:]))
    made = sorted(p.name for p in d13.glob("*_已转.pdf"))
    check("3 份产物都落在**各自原件旁边**",
          made == ["WI-1001_已转.pdf", "WI-1002_已转.pdf", "WI-1003_已转.pdf"], str(made))
    check("⭐ 产物是真 PDF（假引擎写的是 2 页）",
          all(C.read_pdf_info(d13 / n)[0] == 2 for n in made),
          str([(n, C.read_pdf_info(d13 / n)[0]) for n in made]))
    after = {p.name: (p.stat().st_mtime_ns, p.read_bytes())
             for p in d13.iterdir() if p.is_file() and not p.name.endswith("_已转.pdf")}
    check("⭐ 原件一个字节都没动（连 mtime 都没变）", after == before,
          str(sorted(set(after) ^ set(before))))
    vals = [app.tree.item(r, "values") for r in app.tree.get_children()]
    check("表格回填「产物页数」",
          [str(v[CCOL["pages"]]) for v in vals] == ["2", "2", "2"],
          str([v[CCOL["pages"]] for v in vals]))
    check("表格回填状态「转换成功」",
          [v[CCOL["status"]] for v in vals] == ["转换成功"] * 3,
          str([v[CCOL["status"]] for v in vals]))
    check("全成功就不该弹任何警告", len(WARNED) == warned_before, str(WARNED[warned_before:]))

    print("\n[13d] 产物重名 → 覆盖三态（转换模式的说法是「同名 PDF」）")
    ASK_OVERWRITE = None
    app.on_convert()
    check("⭐ 第二次跑必然撞上上一轮的产物 → 弹覆盖窗", len(ASKED) > asked_before,
          str(ASKED[-1].splitlines()[0]) if ASKED else "")
    check("覆盖窗文案是转换模式那份（提「同名 PDF」）",
          "同名 PDF" in ASKED[-1], ASKED[-1].splitlines()[0])
    check("点「取消」→ 一份都没动", "取消" in app.status.cget("text"),
          app.status.cget("text"))

    ASK_OVERWRITE = True
    app.on_convert()
    check("点「是」跑完", pump(root, lambda: not app.busy))
    check("⭐ 点「是」＝原地更新（还是 3 份，没攒出一堆 _2）",
          len(sorted(d13.glob("*_已转*.pdf"))) == 3,
          str(sorted(p.name for p in d13.glob("*_已转*.pdf"))))

    ASK_OVERWRITE = False
    app.on_convert()
    check("点「否」跑完", pump(root, lambda: not app.busy))
    check("⭐ 点「否」＝全部加编号让开（3 份变 6 份，一个现有 PDF 都没被盖）",
          len(sorted(d13.glob("*_已转*.pdf"))) == 6,
          str(sorted(p.name for p in d13.glob("*_已转*.pdf"))))
    ASK_OVERWRITE = True

    print("\n[13e] 失败分支 / 忽略一行")
    bad = d13 / f"{FAKE_FAIL_MARK}-1.xlsx"
    bad.write_bytes(b"boom")
    app._add_sources([str(bad)])
    check("认得出格式 → 进队（4 份）", len(app.conv_plan.items) == 4,
          str(len(app.conv_plan.items)))
    warned_before = len(WARNED)
    app.on_convert()
    check("跑完（带一份失败）", pump(root, lambda: not app.busy))
    vals = [app.tree.item(r, "values") for r in app.tree.get_children()]
    row = next(v for v in vals if str(v[CCOL["src"]]).startswith(FAKE_FAIL_MARK))
    check("失败那行状态是「转换失败」", row[CCOL["status"]] == "转换失败", str(row[CCOL["status"]]))
    check("失败那行「说明」列带上了原因",
          "假引擎故意失败" in str(row[CCOL["note"]]), str(row[CCOL["note"]]))
    check("⭐ 有失败就一定弹提示（不能只悄悄写状态栏）",
          len(WARNED) == warned_before + 1, str(WARNED[warned_before:]))
    check("成功的那些照样转出来了",
          (d13 / "WI-1001_已转.pdf").is_file())

    app.tree.selection_set(app.tree.get_children()[0])
    app.on_ignore()
    vals = [app.tree.item(r, "values") for r in app.tree.get_children()]
    check("右键「忽略此行」→ 状态变「已跳过」", vals[0][CCOL["status"]] == "已跳过",
          str(vals[0][CCOL["status"]]))
    check("还有别的份要转 →「开始转换」仍可用",
          str(app.btn_convert["state"]) == "normal", str(app.btn_convert["state"]))
    app.on_convert()
    check("再跑一遍（跳过的不转）", pump(root, lambda: not app.busy))
    vals = [app.tree.item(r, "values") for r in app.tree.get_children()]
    check("⭐ 被忽略那行仍是「已跳过」（没被偷偷转）", vals[0][CCOL["status"]] == "已跳过",
          str(vals[0][CCOL["status"]]))

    print("\n[13f] 切回替换模式：两套队列各留各的")
    app.on_mode("swap")
    root.update()
    check("回到替换模式", app.mode == "swap", app.mode)
    check("表格换回替换那套列（10 列）",
          tuple(app.tree["columns"]) == tuple(c[0] for c in A.COLUMNS),
          str(tuple(app.tree["columns"])))
    check("扫描件那一行回来了", app.rows["scan"].frame.master.winfo_manager() != "")
    check("①②③ 回来了", app.mod_hint_row.winfo_manager() != "")
    check("「开始转换」收起来了", app.btn_convert.winfo_manager() == "")
    check("「预览 / 开始替换」露出来了",
          app.btn_scan.winfo_manager() != "" and app.btn_run.winfo_manager() != "")
    check("后缀标签变回「必填」", app.lbl_suffix_need.cget("text") == "必填",
          app.lbl_suffix_need.cget("text"))
    check("⭐ 后缀那一格也换回**替换自己那份**（不是转换那份 _已转）",
          app.ent_suffix.get() == SUF, repr(app.ent_suffix.get()))
    check("悬浮说明换回「必填」那一份",
          "必填" in app._tip_suffix.text and "可以不填" not in app._tip_suffix.text,
          app._tip_suffix.text[:48])
    check("⚠️ 转换那份的 _已转 还在（切来切去两份各留各的）",
          app.vars["conv_suffix"].get() == "_已转", repr(app.vars["conv_suffix"].get()))
    check("替换那边没有队列（12c 清空了来源）→ 表格也是空的",
          not app.tree.get_children(), str(app.tree.get_children()))
    app.on_mode("convert")
    root.update()
    check("⭐ 切回来转换队列还在，而且**没被重列**（结果没被抹掉）",
          len(app.tree.get_children()) == 4, str(len(app.tree.get_children())))
    vals = [app.tree.item(r, "values") for r in app.tree.get_children()]
    check("⭐ 上一轮的「转换成功 / 产物页数」原样留着",
          sum(1 for v in vals if v[CCOL["status"]] == "转换成功") == 3
          and sum(1 for v in vals if str(v[CCOL["pages"]]) == "2") == 3,
          str([(v[CCOL["status"]], v[CCOL["pages"]]) for v in vals]))

    print("\n[13g] 改后缀 → 上一轮进度清回「待转换」")
    app.vars["conv_suffix"].set("_另一版")
    root.update()
    vals = [app.tree.item(r, "values") for r in app.tree.get_children()]
    check("⭐ 后缀一改，状态全回「待转换」（上一轮说的是别的目标了）",
          all(v[CCOL["status"]] in ("待转换", "已跳过") for v in vals),
          str([v[CCOL["status"]] for v in vals]))
    check("产物名跟着新后缀走",
          any(str(v[CCOL["out"]]).endswith("_另一版.pdf") for v in vals),
          str([v[CCOL["out"]] for v in vals]))
    check("⭐ 表格里不再有「产物页数」（那批结果已经不属于这一版后缀）",
          all(str(v[CCOL["pages"]]) == "—" for v in vals),
          str([v[CCOL["pages"]] for v in vals]))

    print("\n[13h] ⭐ 来源清单也两个模式各记各的（2026-10-08 用户要求）")
    # ⚠️ 这里**别断言「清单正好等于 [d13]」**：[13e] 那节又往清单里加过两份
    #    （BAD-1.xlsx / notes.txt）。断言「那批还在」就够，写死反而会随前面几节漂移。
    conv_before = list(app.sources)
    check("转换成那边存着刚才挑的那批 Word / Excel",
          str(d13) in conv_before and len(conv_before) >= 3, str(conv_before))
    check("底层真是两个篮子（表驱动，不是靠清空凑出来的）",
          set(app._sources_by_mode) == {"swap", "convert"}
          and app._sources_by_mode["swap"] is not app._sources_by_mode["convert"],
          str(sorted(app._sources_by_mode)))
    app.on_mode("swap")
    root.update()
    check("⭐ 切到替换模式：清单立刻换成替换自己那份（不是刚才那批 xlsx / docx）",
          all(not s.lower().endswith((".xlsx", ".docx", ".xls")) for s in app.sources),
          str(app.sources))
    check("⭐ 框标题跟着换（一个收 PDF、一个收 Word / Excel）",
          "PDF" in app.box_sources.cget("text"), app.box_sources.cget("text"))
    check("⭐ 灰字也跟着换（转换那边不再讲「按文件名排队」）",
          "按文件名" in app.lbl_src_hint.cget("text"), app.lbl_src_hint.cget("text"))
    check("清单控件里显示的就是替换那份（不是转换那份）",
          app.lst_sources.size() == len(app.sources), str(app.lst_sources.size()))
    app.on_mode("convert")
    root.update()
    check("⭐ 切回转换：那批 Word / Excel 一份不多一份不少（没被替换那份挤掉）",
          app.sources == conv_before, f"{app.sources} vs {conv_before}")
    check("⭐ 框标题换回 Word / Excel 那份",
          "Word" in app.box_sources.cget("text"), app.box_sources.cget("text"))
    check("⭐ 灰字也换回转换那份（讲产物落在原件旁边）",
          "落在各自原件旁边" in app.lbl_src_hint.cget("text"),
          app.lbl_src_hint.cget("text"))
    check("清单控件里显示的也是转换那份",
          app.lst_sources.size() == len(app.sources), str(app.lst_sources.size()))
    app.on_mode("swap")           # 收尾回到默认模式，别影响 on_close

    print("\n[13i] ⭐ 转换失败**只弹「确定」** —— 不再追加「是 / 否」询问")
    # 2026-10-09 用户拍板：使用的人会被 Yes/No 问糊涂（以为要装什么软件），
    # 而且公司电脑不让随便装东西。失败原因写在表格「说明」列里已经够用。
    check("⭐ 那个「要不要现在打开设置页」的补救人已经撤掉",
          not hasattr(app, "_offer_permission_help"))
    # ⚠️ 只看「单独的 askyesno」—— 覆盖确认那个 `askyesnocancel` 是**该留着**的，
    #    直接搜 "askyesno" 会把它误伤（2026-10-09 踩过：断言假红一次）。
    _src13i = Path(A.__file__).read_text(encoding="utf-8").replace("askyesnocancel(", "")
    check("⭐ 界面源码里也不再出现单独的 askyesno（防它悄悄长回来）",
          "askyesno" not in _src13i, str(A.__file__))
    check("⭐ 报错类别 err_code 留着（类别和文案解耦，core 那边照归不误）",
          "err_code" in C.ConvertResult.__dataclass_fields__)
    check("⭐ 授权类失败的**说明**照样给得出来（只是不再弹窗问）",
          "自动化" in C.office_error_note("mac", C.format_by_key("word"), "not permitted"))

    # ⑩b：拿副本转的那几份，只在**状态栏**交代一句（不弹窗、也不往行里塞字）。
    _sync_saved = app._sync_run_button
    _plan_saved = app.conv_plan
    app._sync_run_button = lambda: None
    app.conv_plan = None                      # 走「不碰表格」那条分支，专测状态栏
    try:
        app._after_convert([
            C.ConvertResult(src=Path("/tmp/x1.docx"), out=Path("/tmp/x1.pdf"), ok=True,
                            status=C.ConvStatus.OK.value, engine_key="fake", via_copy=True),
            C.ConvertResult(src=Path("/tmp/x2.docx"), out=Path("/tmp/x2.pdf"), ok=True,
                            status=C.ConvStatus.OK.value, engine_key="fake"),
        ])
        check("⭐ 有走副本的 → 状态栏交代「已用临时副本、原件没动」",
              "临时副本" in app.status.cget("text"), app.status.cget("text"))
        app._after_convert([
            C.ConvertResult(src=Path("/tmp/x3.docx"), out=Path("/tmp/x3.pdf"), ok=True,
                            status=C.ConvStatus.OK.value, engine_key="fake"),
        ])
        check("⭐ 没人走副本 → 状态栏不多这句话（别平白多话）",
              "临时副本" not in app.status.cget("text"), app.status.cget("text"))
    finally:
        app.conv_plan, app._sync_run_button = _plan_saved, _sync_saved
    check("⭐ 失败路径上仍然只留「确定」这一种弹窗（showwarning）",
          "showwarning" in Path(A.__file__).read_text(encoding="utf-8"))

    # ⭐⭐ 2026-10-09：界面上的「支持哪些类型」必须**按表 + 平台派生** —— 手写就一定会漏，
    #     漏掉的那处从此对用户说假话（在转不了它的平台上照样写「支持」）。
    #     现在两个平台都一样（Word / Excel 都有），但派生值必须**真等于** core 算出来的那份。
    check("⭐⭐ 界面上的类型名是从 ⑩ 表 + 当前平台派生的（不是写死的字符串）",
          A.CONVERT_KINDS == C.convert_kind_words(), A.CONVERT_KINDS)
    check("  ↳ 几处关键文案都用了这个派生值（加格式时自动跟着变）",
          A.CONVERT_KINDS in A.STATUS_READY_CONVERT
          and A.CONVERT_KINDS in A.SRC_BOX_TITLE_BY_MODE["convert"]
          and A.CONVERT_KINDS in A.CONVERT_BTN_TIP)
    check("  ↳ 「一批只拉起一次」那句也按平台说（Mac 讲授权，Windows 不讲）",
          ("授权" in A.CONVERT_BATCH_NOTE) == (sys.platform == "darwin"))
    check("⭐⭐ 派生值就是 ⑩ 表里**本机转得了的**那些 label 拼起来的"
          "（加格式只动表、界面自动跟；本机转不了的格式不该出现在文案里）",
          A.CONVERT_KINDS == " / ".join(f.label for f in C.platform_formats()),
          A.CONVERT_KINDS)
    # ⚠️ 这条盯的是「手写格式名」这个老毛病：界面上出现一个 ⑩ 表里没有的格式名，
    #    就等于对用户空头支票（说支持、其实转不了）。名单按需扩。
    _stale = [n for n in ("Visio", "PowerPoint") if n not in A.CONVERT_KINDS]
    check("⭐ 界面文案里不许出现表外的格式名（说了就是空头支票）",
          all(n not in s for n in _stale for s in (
              A.STATUS_READY_CONVERT, A.SRC_BOX_TITLE_BY_MODE["convert"],
              A.CONVERT_BTN_TIP)),
          str(_stale))

    # ⭐⭐ 2026-10-09：把文件拖进来（用户要的「左边文件夹、右边软件，选中直接拖」）。
    #     分三段测：① 数据解析（tkdnd 丢过来的是个 Tcl list，不是换行文本）
    #               ② 没挂上时的**降级**（老环境没装 tkinterdnd2 就是那样）
    #               ③ 真挂上之后：文案才提、标题联动、落进清单、去重。
    print("\n[13j] 拖文件进来：解析 / 降级 / 挂载 / 提示")
    # ① 解析 —— 带空格的路径靠花括号包着，硬 .split() 会把它劈成两条假路径
    check("空数据 → 空列表（拖拽中途取消时不该塞进空条目）",
          A.drop_paths("", root.tk.splitlist) == [] and A.drop_paths("   ", None) == [])
    check("⭐⭐ 带空格的路径不被劈开（真 Tcl 解析，靠 tk.splitlist）",
          A.drop_paths("{/a/b c.pdf} /d/e.pdf", root.tk.splitlist)
          == ["/a/b c.pdf", "/d/e.pdf"],
          str(A.drop_paths("{/a/b c.pdf} /d/e.pdf", root.tk.splitlist)))
    check("⭐ 中文 + 空格路径同样解析得对",
          A.drop_paths("{/Users/x/我的 文档/图.pdf}", root.tk.splitlist)
          == ["/Users/x/我的 文档/图.pdf"])
    check("解析器不认就退回按空白切（不抛异常、不丢东西）",
          A.drop_paths("/a.pdf /b.pdf",
                       lambda _s: (_ for _ in ()).throw(ValueError()))
          == ["/a.pdf", "/b.pdf"])
    check("滤掉解析出来的空条目", A.drop_paths("x", lambda _s: ("", "  ", "/ok.pdf"))
          == ["/ok.pdf"])

    # ② 降级 —— 这一刻还没调 enable_dnd，等同于「老环境里没装 tkinterdnd2」
    check("⭐ 没挂上拖放就不硬充有能力（dnd_ok 为假、文案里不提「拖进来」）",
          app.dnd_ok is False and "拖进窗口" not in app._src_hint(), app._src_hint())
    check("  ↳ 悬浮说明里同样不提（说了却拖不动，比不说更糟）",
          "拖进这个窗口" not in app._src_tips["on_add_files"].text)

    # ③ 真挂上
    if A.enable_dnd(root) and app._setup_dnd():
        app.dnd_ok = True
        app._refresh_source_tips()
        app.lbl_src_hint.config(text=app._src_hint())
        check("⭐⭐ 挂上之后文案才提「拖进来」", "拖进窗口" in app._src_hint())
        check("  ↳ 悬浮说明也跟着提", "拖进这个窗口" in app._src_tips["on_add_files"].text)
        before = list(app.sources)
        app._drag_hint(True)
        check("  ↳ 拖着悬停时箱子标题换成「松手」那句",
              app.box_sources.cget("text") == A.DND_BOX_TITLE,
              app.box_sources.cget("text"))
        app._drag_hint(False)
        check("  ↳ 离开 / 松手后标题恢复成这个模式原本那句",
              app.box_sources.cget("text") == A.SRC_BOX_TITLE_BY_MODE[app.mode],
              app.box_sources.cget("text"))
        check("⭐⭐ 提示只改标题 —— 拖进来转一圈不该动任何数据", app.sources == before)
        app.sources[:] = []
        app._on_drop(_DropEvent("/tmp/drop_a.pdf"))
        check("⭐ 松手 → 落进「原件来源」清单（走的就是 _add_sources 那条路）",
              app.sources == ["/tmp/drop_a.pdf"], str(app.sources))
        app._on_drop(_DropEvent("{/tmp/带 空格.pdf} /tmp/drop_b.docx"))
        check("  ↳ 一次拖多个也照收（含带空格的）",
              app.sources == ["/tmp/drop_a.pdf", "/tmp/带 空格.pdf", "/tmp/drop_b.docx"],
              str(app.sources))
        app.sources[:] = []
        app._on_drop(_DropEvent("/tmp/drop_a.pdf"))
        app._on_drop(_DropEvent("/tmp/drop_a.pdf"))
        check("  ↳ 同一个文件拖两次不会加重复（_add_sources 本来就去重）",
              app.sources == ["/tmp/drop_a.pdf"], str(app.sources))
        app.sources[:] = []
    else:
        print("  --  这台机器没挂上拖放（没装 tkinterdnd2？），③ 挂载那段跳过")

    app.on_close()          # 关窗这条路也要走一遍（它会保存配置）
    try:                    # on_close 已经把窗口销毁了，再 destroy 会报错
        root.destroy()
    except tk.TclError:
        pass
    shutil.rmtree(TMP, ignore_errors=True)
    print(f"\n通过 {len(PASS)} 项，失败 {len(FAIL)} 项")
    for n in FAIL:
        print("  失败：" + n)
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
