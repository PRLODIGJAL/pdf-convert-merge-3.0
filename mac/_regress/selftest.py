#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""核心自测 —— 全部产物写在系统临时目录，不污染桌面。

跑法：.venv/bin/python -B _regress/selftest.py
"""
from __future__ import annotations

import contextlib
import inspect
import io
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import seq_core as C  # noqa: E402

PASS: list[str] = []
FAIL: list[str] = []
# 默认测试后缀 —— v3.0 起「输出文件名后缀」是必填的（产物写在原件旁边）
SUF = "_已替换"
# 名字里带这个标记的原件，会被注入的假转换引擎故意判失败（用来验失败分支）
FAKE_FAIL_MARK = "BAD"


def check(name: str, cond: bool, detail: str = "") -> None:
    (PASS if cond else FAIL).append(name)
    print(f"  {'OK ' if cond else 'NG '} {name}" + (f"  <{detail}>" if detail and not cond else ""))


def _raises(fn) -> bool:
    """这个调用应该抛异常（错了要炸，不能静默退回默认）。"""
    try:
        fn()
    except Exception:  # noqa: BLE001
        return True
    return False


# ── 纯标准库构造 PDF（每页写一句不同的文字，好验证内容有没有被换掉）─────────
def build_pdf(path: Path, pages: list[tuple[float, float, str]]) -> None:
    objs: dict[int, bytes] = {}
    page_nums = [4 + 2 * i for i in range(len(pages))]
    objs[1] = b"<< /Type /Catalog /Pages 2 0 R >>"
    kids = " ".join(f"{n} 0 R" for n in page_nums)
    objs[2] = f"<< /Type /Pages /Kids [{kids}] /Count {len(pages)} >>".encode()
    objs[3] = b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"
    for i, (w, h, label) in enumerate(pages):
        pn, cn = page_nums[i], page_nums[i] + 1
        objs[pn] = (
            f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {w} {h}] "
            f"/Resources << /Font << /F1 3 0 R >> >> /Contents {cn} 0 R >>"
        ).encode()
        stream = f"BT /F1 28 Tf 60 100 Td ({label}) Tj ET".encode()
        objs[cn] = f"<< /Length {len(stream)} >>\nstream\n".encode() + stream + b"\nendstream"

    out = io.BytesIO()
    out.write(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets: dict[int, int] = {}
    for num in sorted(objs):
        offsets[num] = out.tell()
        out.write(f"{num} 0 obj\n".encode() + objs[num] + b"\nendobj\n")
    xref_pos = out.tell()
    max_num = max(objs) + 1
    out.write(f"xref\n0 {max_num}\n".encode())
    out.write(b"0000000000 65535 f \n")
    for num in range(1, max_num):
        off = offsets.get(num)
        out.write((f"{off:010d} 00000 n \n" if off else "0000000000 65535 f \n").encode())
    out.write(
        f"trailer\n<< /Size {max_num} /Root 1 0 R >>\nstartxref\n{xref_pos}\n%%EOF\n".encode()
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(out.getvalue())


def write_min_xlsx(path: Path, sheets: list[tuple[str, str]]) -> None:
    """造一个最小可用的 .xlsx —— sheets = [(表名, 可见性)]，可见性 = visible/hidden。

    只写「数可见工作表」要读的那两个成员（[Content_Types].xml + xl/workbook.xml）：
    回归要验的是**数得出几个可见表**，不需要它真能被 Excel 打开。
    """
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    items = "".join(
        f'<sheet name="{t}" sheetId="{i}" state="{st}" r:id="rId{i}"/>'
        for i, (t, st) in enumerate(sheets, start=1))
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", "<Types/>")
        z.writestr("xl/workbook.xml",
                   f'<workbook xmlns="{ns}" xmlns:r="{rns}">'
                   f"<sheets>{items}</sheets></workbook>")


def prod(d: Path, name: str, suffix: str = SUF) -> Path:
    """产物路径 —— v3.0 起产物就写在**原件旁边**，名字 = 原件名 + 后缀。

    测试里所有「去输出目录找结果」的地方都改成走这个函数：
    落点变了，但「结果对不对」的判据一个字都没变。
    """
    p = Path(name)
    return d / f"{p.stem}{suffix}{p.suffix}"


def mkopts(**kw) -> C.Options:
    """造 Options —— suffix 默认给 SUF。

    v3.0 起 suffix 是**必填**的（产物写在原件旁边，没后缀就会盖掉原件），
    所以测试里不给的话默认补一个；要验「后缀为空会怎样」的用例显式传空串。
    """
    kw.setdefault("suffix", SUF)
    return C.Options(**kw)


def page_texts(pdf: Path) -> list[str]:
    from pypdf import PdfReader

    return [(p.extract_text() or "").strip() for p in PdfReader(str(pdf)).pages]


def _disp(pdf: Path) -> list[tuple[float, float]]:
    """每页的「显示尺寸」（含 /Rotate 修正），用来验旋转有没有真的生效。"""
    from pypdf import PdfReader

    return [(round(float(b[0]), 1), round(float(b[1]), 1))
            for b in (C.display_size(p) for p in PdfReader(str(pdf)).pages)]


def _has_rot_matrix(pdf: Path, page: int = 0) -> bool:
    """内容流里有没有带 90° 旋转分量的 cm 矩阵（a b c d e f cm 里的 b、c 非零）。

    只比 mediabox 是不够的 —— 框变大了但内容没转，拼出来还是躺着的。
    """
    import re
    from pypdf import PdfReader

    try:
        data = PdfReader(str(pdf)).pages[page].get_contents().get_data().decode("latin-1")
    except Exception:  # noqa: BLE001
        return False
    for m in re.finditer(r"([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+"
                         r"([-\d.]+)\s+([-\d.]+)\s+cm", data):
        try:
            _a, b, c, _d, _e, _f = (float(m.group(i)) for i in range(1, 7))
        except ValueError:
            continue
        if abs(b) > 1e-6 and abs(c) > 1e-6:
            return True
    return False


def page_sizes(pdf: Path) -> list[tuple[float, float]]:
    from pypdf import PdfReader

    out = []
    for p in PdfReader(str(pdf)).pages:
        b = p.mediabox
        out.append((round(float(b.width), 1), round(float(b.height), 1)))
    return out


def make_corpus(root: Path, n_original: int, pages_each: int,
                n_scan: int, size=(595.0, 842.0)) -> tuple[Path, Path]:
    """造语料：原件 001..N（每份 pages_each 页），扫描件 n_scan 页。"""
    src_d = root / "原件"
    src_d.mkdir(parents=True, exist_ok=True)
    for i in range(1, n_original + 1):
        build_pdf(src_d / f"{i:03d}.pdf",
                  [(size[0], size[1], f"ORI{i}-P{j}") for j in range(1, pages_each + 1)])
    scan = root / "扫描件_签字版.pdf"
    build_pdf(scan, [(size[0], size[1], f"SCAN-{k}") for k in range(1, n_scan + 1)])
    return src_d, scan


def main() -> int:
    root = Path(tempfile.gettempdir()) / "pdfseq_selftest"
    if root.exists():
        shutil.rmtree(root)
    root.mkdir(parents=True)
    print(f"临时工作目录：{root}\n")

    # ── [1] 文件名顺序（＝文件夹里按「名称」看到的顺序） ──────────────────
    print("[1] 排序：文件名顺序")
    names = ["010.pdf", "002.pdf", "001.pdf", "100.pdf", "020.pdf"]
    paths = sorted((Path(n) for n in names), key=lambda p: C.natural_key(p.name))
    got = [p.name for p in paths]
    check("010 不会排到 002 前面（不是字典序）",
          got == ["001.pdf", "002.pdf", "010.pdf", "020.pdf", "100.pdf"], str(got))
    mixed = ["WI-9.pdf", "WI-10.pdf", "WI-100.pdf", "WI-2.pdf"]
    got2 = [p.name for p in sorted((Path(n) for n in mixed),
                                   key=lambda p: C.natural_key(p.name))]
    check("带前缀的编号也按数值排",
          got2 == ["WI-2.pdf", "WI-9.pdf", "WI-10.pdf", "WI-100.pdf"], str(got2))

    # 2026-10-08 用户拍板：不改名成 001/002，直接用 WI-xxxx 原名排
    real = ["WI-3852.pdf", "WI-1032.pdf", "WI-2001.pdf", "WI-999.pdf"]
    got1b = [p.name for p in sorted((Path(n) for n in real),
                                    key=lambda p: C.natural_key(p.name))]
    check("真实 WI 原名按名称排（999 在 1032 前、3852 在最后）",
          got1b == ["WI-999.pdf", "WI-1032.pdf", "WI-2001.pdf", "WI-3852.pdf"], str(got1b))
    case_mixed = ["banana.pdf", "Apple.pdf", "cherry.pdf"]
    got1c = [p.name for p in sorted((Path(n) for n in case_mixed),
                                    key=lambda p: C.natural_key(p.name))]
    check("字母不分大小写（和资源管理器一致）",
          got1c == ["Apple.pdf", "banana.pdf", "cherry.pdf"], str(got1c))
    check("默认排序规则就是「文件名顺序」", C.DEFAULT_ORDER == "name", C.DEFAULT_ORDER)
    check("排序表里能查到文件名顺序这条",
          any(r.key == "name" for r in C.ORDER_RULES))
    check("改名成 001/002 不再是前提（排序键不吃「必须是纯数字」这一套）",
          [p.name for p in sorted((Path(n) for n in ["WI-1032.pdf", "WI-9.pdf"]),
                                  key=lambda p: C.natural_key(p.name))]
          == ["WI-9.pdf", "WI-1032.pdf"])

    # 备用规则：按修改时间（原件没有编号时的退路）
    w_mt = root / "按时间排"
    w_mt.mkdir(parents=True)
    a_p, b_p, c_p = w_mt / "a.pdf", w_mt / "b.pdf", w_mt / "c.pdf"
    for p in (a_p, b_p, c_p):
        p.write_bytes(b"%PDF-1.4\n")
    # 故意让修改时间的先后（b → c → a）和文件名顺序（a → b → c）不一样
    for k, p in enumerate([b_p, c_p, a_p]):
        os.utime(p, (1_700_000_000 + k * 100, 1_700_000_000 + k * 100))
    got_mt = [p.name for p in C.sort_originals([a_p, b_p, c_p], "mtime")]
    check("备用规则：按修改时间从早到晚（b → c → a）",
          got_mt == ["b.pdf", "c.pdf", "a.pdf"], str(got_mt))
    got_nm = [p.name for p in C.sort_originals([c_p, a_p, b_p], "name")]
    check("同一批文件用文件名顺序又是另一回事（a → b → c）",
          got_nm == ["a.pdf", "b.pdf", "c.pdf"], str(got_nm))
    check("传一个不存在的规则会直接报错（不静默退回默认）",
          _raises(lambda: C.sort_originals([a_p], "不存在的规则")))

    # ── [2] 收集原件：排除扫描件自身、排除上一轮的产物 ────────────────────
    print("\n[2] 收集原件")
    w = root / "收集"
    src_d, scan = make_corpus(w, 3, 2, 3)
    opts = mkopts(sources=(str(src_d),), scan_pdf=str(scan), suffix=SUF)
    got3 = [p.name for p in C.collect_originals(opts.sources, opts)]
    check("没有产物时三份原件都在",
          got3 == ["001.pdf", "002.pdf", "003.pdf"], str(got3))

    # ⑧ 产物识别（v3.0 新规则）：产物落在原件旁边，第二次跑不能被当成原件收
    prod1 = src_d / f"001{SUF}.pdf"
    build_pdf(prod1, [(595, 842, "PROD")])
    skipped: list[Path] = []
    got4 = [p.name for p in C.collect_originals(opts.sources, opts, skipped)]
    check("上一轮的产物被认出来并跳过（数量不会翻倍）",
          got4 == ["001.pdf", "002.pdf", "003.pdf"], str(got4))
    check("跳过了哪几份也如实报出来",
          [p.name for p in skipped] == [f"001{SUF}.pdf"], str([p.name for p in skipped]))
    check("sibling_source_of 能指回那份原件",
          C.sibling_source_of(prod1, SUF) == src_d / "001.pdf",
          str(C.sibling_source_of(prod1, SUF)))
    check("后缀为空时什么都不猜（不把任何文件当产物）",
          C.sibling_source_of(prod1, "") is None)
    check("名字带后缀、但同目录没有对应原件的，不当产物（宁可漏判也不误伤）",
          C.sibling_source_of(src_d / f"没有原件的{SUF}.pdf", SUF) is None)
    weird = src_d / f"{SUF}.pdf"          # 名字整个就是后缀
    weird.write_bytes(b"%PDF-1.4\n")
    check("名字整个就是后缀的也不认", C.sibling_source_of(weird, SUF) is None)
    weird.unlink()

    # 「加编号让开」留下的尾巴也要认 —— 否则用户选了「否」之后，
    # 下一轮这些 _2 会被当成原件收进来（数量翻倍 → 整批拒绝），还得手工删。
    num2 = src_d / f"001{SUF}_2.pdf"
    num10 = src_d / f"001{SUF}_10.pdf"
    build_pdf(num2, [(595, 842, "PROD")])
    build_pdf(num10, [(595, 842, "PROD")])
    check("编号尾巴 _2 也算产物（指回同一份原件）",
          C.sibling_source_of(num2, SUF) == src_d / "001.pdf",
          str(C.sibling_source_of(num2, SUF)))
    check("两位数的编号 _10 也认", C.sibling_source_of(num10, SUF) == src_d / "001.pdf")
    check("编号件跟无编号件一起被跳过",
          [p.name for p in C.collect_originals(opts.sources, opts)]
          == ["001.pdf", "002.pdf", "003.pdf"],
          str([p.name for p in C.collect_originals(opts.sources, opts)]))
    num2.unlink()
    num10.unlink()
    # 反例：后缀本身不以「_数字」结尾时，`001_2.pdf` 这种不许被误认成产物
    plain2 = src_d / "001_2.pdf"
    build_pdf(plain2, [(595, 842, "我是一份真原件")])
    check("不带后缀的 `001_2.pdf` 不许被误认成产物",
          C.sibling_source_of(plain2, SUF) is None)
    plain2.unlink()

    prod1.unlink()
    check("产物被删掉之后又回到三份原件",
          len(C.collect_originals(opts.sources, opts)) == 3)

    # 扫描件就放在原件文件夹里（用户最可能的摆法）—— 它绝不能被当成一份原件
    w2 = root / "扫描件混在里面"
    src2, scan2 = make_corpus(w2, 3, 2, 3)
    shutil.move(str(scan2), str(src2 / "zz扫描件.pdf"))
    opts2 = mkopts(sources=(str(src2),), scan_pdf=str(src2 / "zz扫描件.pdf"))
    got5 = [p.name for p in C.collect_originals(opts2.sources, opts2)]
    check("扫描件本身放在原件文件夹里会被排除",
          got5 == ["001.pdf", "002.pdf", "003.pdf"], str(got5))

    # ── 多来源（v3.0 新模型）─────────────────────────────────────────────
    # 「连同子文件夹」已删：子目录里的原件**不算**，要收就把它作为来源加进清单。
    w3 = root / "多来源"
    src3, scan3 = make_corpus(w3, 2, 2, 4)
    sub3 = src3 / "子目录"
    sub3.mkdir()
    build_pdf(sub3 / "003.pdf", [(595, 842, "ORI3-P1"), (595, 842, "ORI3-P2")])
    build_pdf(sub3 / "004.pdf", [(595, 842, "ORI4-P1"), (595, 842, "ORI4-P2")])
    opt_bare = mkopts(sources=(str(src3),), scan_pdf=str(scan3))
    check("只给文件夹：子目录里的**不收**（递归已删）",
          len(C.collect_originals(opt_bare.sources, opt_bare)) == 2,
          str([p.name for p in C.collect_originals(opt_bare.sources, opt_bare)]))
    opts3 = mkopts(sources=(str(src3), str(sub3)), scan_pdf=str(scan3))
    got5 = [p.name for p in C.collect_originals(opts3.sources, opts3)]
    check("把子目录也作为来源加进来 → 4 份都在队里，且按文件名统一排",
          got5 == ["001.pdf", "002.pdf", "003.pdf", "004.pdf"], str(got5))

    # 单个文件也能当来源；文件夹 + 单文件混着来
    opts3c = mkopts(sources=(str(src3 / "002.pdf"), str(sub3 / "004.pdf")))
    check("来源可以就是单个 PDF 文件（挑着处理）",
          [p.name for p in C.collect_originals(opts3c.sources, opts3c)] == ["002.pdf", "004.pdf"],
          str([p.name for p in C.collect_originals(opts3c.sources, opts3c)]))

    # 去重：既加了文件夹、又单加了里面的某个文件 —— 不能算两份
    opts3d = mkopts(sources=(str(src3), str(src3 / "001.pdf")), scan_pdf=str(scan3))
    check("同一个文件被指到两次会自动去重",
          len(C.collect_originals(opts3d.sources, opts3d)) == 2,
          str([p.name for p in C.collect_originals(opts3d.sources, opts3d)]))

    # ⭐ 排序是「全局按文件名」—— 来自哪个文件夹不影响先后
    mixed = root / "混排"
    ma = mixed / "甲"; mb = mixed / "乙"
    for d in (ma, mb):
        d.mkdir(parents=True)
    build_pdf(ma / "WI-2001.pdf", [(595, 842, "A1")])
    build_pdf(mb / "BOM-1998.pdf", [(595, 842, "B1")])
    build_pdf(mb / "WI-0001.pdf", [(595, 842, "B2")])
    o_mix = mkopts(sources=(str(ma), str(mb)))
    check("跨文件夹时按文件名统一排（BOM-1998 在 WI-0001 前面）",
          [p.name for p in C.collect_originals(o_mix.sources, o_mix)]
          == ["BOM-1998.pdf", "WI-0001.pdf", "WI-2001.pdf"],
          str([p.name for p in C.collect_originals(o_mix.sources, o_mix)]))

    # ── [3] 阻断闸门 ─────────────────────────────────────────────────────
    print("\n[3] 阻断闸门（有问题就整批不输出）")
    g = root / "闸门"
    src_a, scan_a = make_corpus(g, 5, 2, 5)

    p1 = C.plan_batch(mkopts(sources=(str(g / "不存在"),), scan_pdf=str(scan_a)))
    check("原件文件夹不存在 → 阻断", p1.blocked and "不存在" in p1.block_reason, p1.block_reason)

    empty = g / "空文件夹"
    empty.mkdir()
    p2 = C.plan_batch(mkopts(sources=(str(empty),), scan_pdf=str(scan_a)))
    check("原件文件夹里没有 PDF → 阻断", p2.blocked, p2.block_reason)

    p3 = C.plan_batch(mkopts(sources=(str(src_a),), scan_pdf=str(g / "没有这个.pdf")))
    check("扫描件不存在 → 阻断", p3.blocked and "不存在" in p3.block_reason, p3.block_reason)

    bad = g / "坏扫描件.pdf"
    bad.write_bytes(b"not a pdf at all")
    p4 = C.plan_batch(mkopts(sources=(str(src_a),), scan_pdf=str(bad)))
    check("扫描件读不出来 → 阻断", p4.blocked, p4.block_reason)

    # 数量对不上（用两个独立目录造语料，避免互相覆盖）
    src_many, scan_short = make_corpus(g / "少页", 5, 2, 3)
    p5 = C.plan_batch(mkopts(sources=(str(src_many),), scan_pdf=str(scan_short)))
    check("原件 5 份 vs 扫描件 3 页 → 阻断",
          p5.blocked and "5" in p5.block_reason and "3" in p5.block_reason, p5.block_reason)
    check("阻断原因里写清差几份", "差 2" in p5.block_reason, p5.block_reason)

    src_eq, scan_eq = make_corpus(g / "相等", 5, 2, 5)
    p5b = C.plan_batch(mkopts(sources=(str(src_eq),), scan_pdf=str(scan_eq)))
    check("两边都是 5 → 不阻断", not p5b.blocked, p5b.block_reason)

    src_few, scan_many = make_corpus(g / "多页", 3, 2, 5)
    p5c = C.plan_batch(mkopts(sources=(str(src_few),), scan_pdf=str(scan_many)))
    check("扫描件比原件多也照样阻断",
          p5c.blocked and "差 2" in p5c.block_reason, p5c.block_reason)

    src6, scan6 = make_corpus(g / "坏原件", 3, 2, 3)
    (src6 / "002.pdf").write_bytes(b"broken")
    p6 = C.plan_batch(mkopts(sources=(str(src6),), scan_pdf=str(scan6)))
    check("原件读不出页数 → 阻断（顺序会错乱）",
          p6.blocked and "002.pdf" in p6.block_reason, p6.block_reason)

    # ── [4] 顺序配对 ─────────────────────────────────────────────────────
    print("\n[4] 顺序配对")
    w4 = root / "配对"
    src_b, scan_b = make_corpus(w4, 4, 2, 4)
    plan = C.plan_batch(mkopts(sources=(str(src_b),), scan_pdf=str(scan_b)))
    check("配对数量正确", len(plan.items) == 4)
    check("第 N 份原件 ← 扫描件第 N 页",
          [it.scan_index for it in plan.items] == [1, 2, 3, 4],
          str([it.scan_index for it in plan.items]))
    check("原件顺序就是 001/002/003/004",
          [it.src.name for it in plan.items] == ["001.pdf", "002.pdf", "003.pdf", "004.pdf"])
    C.reindex(plan, True)
    check("倒序后 001 用最后一页",
          [it.scan_index for it in plan.items] == [4, 3, 2, 1],
          str([it.scan_index for it in plan.items]))
    C.reindex(plan, False)
    check("切回正序", [it.scan_index for it in plan.items] == [1, 2, 3, 4])

    # 调页码：行不动（行的顺序＝名称序，那是唯一的排队依据），只换「用哪一页」
    check("与下一行对调：两行的页码互换",
          C.swap_pages(plan, 1, 1) is True
          and [it.scan_index for it in plan.items] == [2, 1, 3, 4],
          str([it.scan_index for it in plan.items]))
    check("对调时行的顺序一动不动",
          [it.src.name for it in plan.items] == ["001.pdf", "002.pdf", "003.pdf", "004.pdf"])
    C.swap_pages(plan, 1, 1)
    check("再对调一次换回来（同一个按钮点两次 = 没变）",
          [it.scan_index for it in plan.items] == [1, 2, 3, 4])
    check("首行往上一行对调会失败（它没有上一行）", C.swap_pages(plan, 1, -1) is False)
    check("末行往下一行对调会失败（它没有下一行）", C.swap_pages(plan, 4, 1) is False)
    check("指定超出扫描件范围的页码 → 拒绝", C.assign_page(plan, 1, 99)[0] is False)

    changed, note = C.assign_page(plan, 4, 2)
    check("把第 4 行改成用第 2 页 → 原先用第 2 页的那行自动改成第 4 页",
          changed and [it.scan_index for it in plan.items] == [1, 4, 3, 2],
          str([it.scan_index for it in plan.items]))
    check("一页不会被两行同时用（4 行 4 个不同页码）",
          len({it.scan_index for it in plan.items}) == 4)
    check("对调说明里点名了另一份原件", "002.pdf" in note, note)
    check("改成自己已经在用的页码 → 无事发生",
          C.assign_page(plan, 1, 1)[0] is False)

    # ── [5] 替换结果 ─────────────────────────────────────────────────────
    print("\n[5] 替换结果")
    w5 = root / "替换"
    src_c, scan_c = make_corpus(w5, 3, 3, 3)
    before = {p.name: page_texts(p) for p in src_c.glob("*.pdf")}
    plan_c = C.plan_batch(mkopts(sources=(str(src_c),), scan_pdf=str(scan_c)))
    results, _ = C.run_batch(plan_c, mkopts(sources=(str(src_c),), scan_pdf=str(scan_c)))
    check("全部成功", all(r.ok for r in results),
          str([(r.src.name, r.note) for r in results if not r.ok]))
    for i, r in enumerate(results, start=1):
        texts = page_texts(r.out)
        check(f"001…00{i}：第 1 页是扫描件第 {i} 页",
              texts and f"SCAN-{i}" in texts[0], str(texts[:1]))
        check(f"001…00{i}：第 2 页起还是原件的内容",
              texts[1:] == before[f"{i:03d}.pdf"][1:], str(texts[1:]))
        check(f"001…00{i}：页数没变（3 页）", len(texts) == 3, str(len(texts)))
    # ⚠️ v3.0：产物就躺在同一目录里，所以「原件没被动」不能整个目录 glob ——
    #    只会把新产出的那几份也算进去。只比对**跑之前记下的那几份原件**。
    check("原件本身没被改动",
          {n: page_texts(src_c / n) for n in before} == before)
    check("输出第 1 页尺寸 = 原件首页尺寸（贴合）",
          page_sizes(results[0].out)[0] == (595.0, 842.0), str(page_sizes(results[0].out)[0]))

    # 关掉缩放
    opts_asis = mkopts(sources=(str(src_c),), scan_pdf=str(scan_c),
                          scale_to_original=False, overwrite=True)
    plan_c2 = C.plan_batch(opts_asis)
    results2, _ = C.run_batch(plan_c2, opts_asis)
    check("关掉缩放后依然成功", all(r.ok for r in results2))
    check("关掉缩放：第 1 页是扫描页原尺寸",
          page_sizes(results2[0].out)[0] == (595.0, 842.0))

    # 不同尺寸：验证「贴合」真的在缩放
    w6 = root / "异形"
    src_d2 = w6 / "原件"
    src_d2.mkdir(parents=True)
    for i in range(1, 3):
        build_pdf(src_d2 / f"{i:03d}.pdf",
                  [(400.0, 400.0, f"ORI{i}-P1"), (400.0, 400.0, f"ORI{i}-P2")])
    scan_d2 = w6 / "扫描件.pdf"
    build_pdf(scan_d2, [(800.0, 800.0, "SCAN-1"), (800.0, 800.0, "SCAN-2")])
    opts_d = mkopts(sources=(str(src_d2),), scan_pdf=str(scan_d2),
                       overwrite=True)
    plan_d = C.plan_batch(opts_d)
    res_d, _ = C.run_batch(plan_d, opts_d)
    check("比例一致时缩放到原件首页尺寸",
          page_sizes(res_d[0].out)[0] == (400.0, 400.0), str(page_sizes(res_d[0].out)[0]))

    # 比例差太多 → 原样
    w7 = root / "比例不同"
    src_e = w7 / "原件"
    src_e.mkdir(parents=True)
    for i in range(1, 3):
        build_pdf(src_e / f"{i:03d}.pdf",
                  [(842.0, 595.0, f"ORI{i}-P1"), (842.0, 595.0, f"ORI{i}-P2")])
    scan_e = w7 / "扫描件.pdf"
    build_pdf(scan_e, [(595.0, 842.0, "SCAN-1"), (595.0, 842.0, "SCAN-2")])
    # 关掉⑦自动方向：竖向扫描页套不进横向原件 → 原样保留，不硬塞
    opts_e = mkopts(sources=(str(src_e),), scan_pdf=str(scan_e),
                       overwrite=True, auto_rotate=False)
    plan_e = C.plan_batch(opts_e)
    res_e, _ = C.run_batch(plan_e, opts_e)
    check("关掉自动方向：版面比例差太多 → 原样保留，不硬塞",
          res_e[0].fit_mode == C.FIT_ASIS and page_sizes(res_e[0].out)[0] == (595.0, 842.0),
          f"{res_e[0].fit_mode}/{page_sizes(res_e[0].out)[0]}")

    # 默认（开着自动方向）：扫描页先左转 90° 变横向 → 正好贴合，不再是一片白边
    opts_e2 = mkopts(sources=(str(src_e),), scan_pdf=str(scan_e),
                        overwrite=True)
    plan_e2 = C.plan_batch(opts_e2)
    res_e2, _ = C.run_batch(plan_e2, opts_e2)
    check("自动左转后：比例吻合，从「原样」变成「贴合」",
          res_e2[0].fit_mode == C.FIT_FIT and page_sizes(res_e2[0].out)[0] == (842.0, 595.0),
          f"{res_e2[0].fit_mode}/{page_sizes(res_e2[0].out)[0]}")

    # ── [6] 产物落点（⑨ 表）+ 后缀必填（④）────────────────────────────────
    print("\n[6] 产物落点 / 后缀必填")
    w8 = root / "命名"
    src_f, scan_f = make_corpus(w8, 2, 2, 2)
    base_names = sorted(p.name for p in src_f.glob("*.pdf"))

    # ④ 后缀必填：core 层直接拒绝，一个字节都不写
    empty = mkopts(sources=(str(src_f),), scan_pdf=str(scan_f), suffix="")
    check("后缀为空 → run_batch 直接拒绝",
          C.run_batch(C.plan_batch(empty), empty) == ([], [C.SUFFIX_REQUIRED_MSG]))
    check("被拒之后磁盘上一个产物都没有",
          sorted(p.name for p in src_f.glob("*.pdf")) == base_names,
          str(sorted(p.name for p in src_f.glob("*.pdf"))))
    check("suffix_error 是唯一判据",
          bool(C.suffix_error(empty)) and not C.suffix_error(mkopts(suffix=SUF)))

    # 正常跑一次：产物落在原件旁边，名字 = 原名 + 后缀
    of = mkopts(sources=(str(src_f),), scan_pdf=str(scan_f), suffix="_已签字首页")
    C.run_batch(C.plan_batch(of), of)
    check("后缀生效：产物就在原件旁边",
          prod(src_f, "001.pdf", "_已签字首页").is_file(),
          str(sorted(p.name for p in src_f.iterdir())))
    check("原件本身还在（只读承诺没变）", (src_f / "001.pdf").is_file())
    check("落点 = 原件自己所在的目录（resolve_target）",
          C.resolve_target(src_f / "001.pdf", mkopts(), set()) == prod(src_f, "001.pdf"),
          str(C.resolve_target(src_f / "001.pdf", mkopts(), set())))
    check("产物名必然 ≠ 原件名（后缀非空换来的安全）",
          C.out_name_for(src_f / "001.pdf", SUF) != "001.pdf")

    # ⑨ 表第二行：目标是**认得出的上一轮产物** → 默认就覆盖（重跑＝原地更新）
    C.run_batch(C.plan_batch(of), of)
    check("重跑：默认覆盖上次的产物，不产生 _2",
          prod(src_f, "001.pdf", "_已签字首页").is_file()
          and not list(src_f.glob("*已签字首页_2.pdf")),
          str(sorted(p.name for p in src_f.iterdir())))
    check("重跑时预览里只有 2 份原件（产物被认出来、没当成原件收）",
          len(C.plan_batch(of).items) == 2, str(len(C.plan_batch(of).items)))

    # ⑨ 表第三行的兜底（加编号让开）：逻辑单测，保证「谁都不覆盖」
    src_z, _scan_z = make_corpus(root / "撞名", 1, 2, 1)
    taken = prod(src_z, "001.pdf")
    build_pdf(taken, [(595, 842, "别人放的文件")])
    # ⚠️ 别拿 page_texts 比中文：build_pdf 写进去的是 UTF-8 字节、pypdf 取出来是
    #    latin-1 乱码（字体没带 ToUnicode）。要验「一个字节都没动」就**真的比字节**。
    raw_z = taken.read_bytes()
    used: set[Path] = set()
    got_z = C._unique_out_path(taken, used)
    check("撞上已有文件 → 加编号让开（_2）", got_z.name == f"001{SUF}_2.pdf", got_z.name)
    check("再撞一次就轮到 _3",
          C._unique_out_path(taken, used).name == f"001{SUF}_3.pdf")
    check("那个已有文件一个字节都没动",
          taken.read_bytes() == raw_z and taken.is_file())

    # ── ⑨ 的三态（AUTO / YES / NO）────────────────────────────────────────
    # 这是「界面弹窗点否」能不能真生效的关键：只分 True/False 两态时，
    # 点「否」照样会被 AUTO 的「认得出是产物就覆盖」吃掉 —— 弹窗等于摆设。
    print("\n[6b] 撞名三态：AUTO 更新 / YES 一律盖 / NO 一律让开")
    src_t, scan_t = make_corpus(root / "三态", 2, 2, 2)
    tgt1 = prod(src_t, "001.pdf")
    run_of = mkopts(sources=(str(src_t),), scan_pdf=str(scan_t))
    no_of = mkopts(sources=(str(src_t),), scan_pdf=str(scan_t), overwrite=C.OVERWRITE_NO)
    yes_of = mkopts(sources=(str(src_t),), scan_pdf=str(scan_t), overwrite=C.OVERWRITE_YES)

    # 先按默认口径跑一轮 → 磁盘上就有了「认得出是自家产物」的底稿
    C.run_batch(C.plan_batch(run_of), run_of)
    check("底稿就位：产物已存在且认得出是自家产物",
          tgt1.is_file() and C.is_product(tgt1, SUF))
    old_raw = tgt1.read_bytes()

    check("三态常量是 AUTO(None) / YES(True) / NO(False)",
          (C.OVERWRITE_AUTO, C.OVERWRITE_YES, C.OVERWRITE_NO) == (None, True, False))
    check("Options 默认就是 AUTO", mkopts().overwrite is C.OVERWRITE_AUTO)

    check("AUTO：认得出的产物 → 原地更新（不产生 _2）",
          C.resolve_target(src_t / "001.pdf", run_of, set()) == tgt1)
    check("NO：用户点了「否」→ 连自家产物也让开（加 _2）",
          C.resolve_target(src_t / "001.pdf", no_of, set()).name == f"001{SUF}_2.pdf")
    check("YES：用户点了「是」→ 就用原名覆盖",
          C.resolve_target(src_t / "001.pdf", yes_of, set()) == tgt1)

    # 真跑一遍 NO：磁盘上多出 _2，原来那份产物一个字节都不能动
    C.run_batch(C.plan_batch(no_of), no_of)
    check("NO 实跑：两份产物都被让开成 _2",
          (src_t / f"001{SUF}_2.pdf").is_file()
          and (src_t / f"002{SUF}_2.pdf").is_file(),
          str(sorted(p.name for p in src_t.iterdir())))
    check("NO 实跑：原来那份产物一个字节都没动", tgt1.read_bytes() == old_raw)
    check("NO 实跑：原件仍然没被动", (src_t / "001.pdf").is_file())
    check("NO 实跑：下一次预览仍是 2 份原件（_2 编号件也被认成产物、没被当原件收）",
          len(C.plan_batch(run_of).items) == 2, str(len(C.plan_batch(run_of).items)))

    # 手放一个「名字像产物、但找不到对应原件」的文件：认不出 → AUTO 不许碰
    stray = src_t / f"999{SUF}.pdf"
    build_pdf(stray, [(595, 842, "别人手放的同名文件")])
    stray_raw = stray.read_bytes()
    check("认不出的同名文件：AUTO 让开、一个字节都不动",
          C.resolve_target(src_t / "999.pdf", run_of, set()).name == f"999{SUF}_2.pdf"
          and stray.read_bytes() == stray_raw)
    check("认不出的同名文件：YES 才敢直接盖（是用户自己点的）",
          C.resolve_target(src_t / "999.pdf", yes_of, set()) == stray)

    # ── [7] 后缀规范化 ───────────────────────────────────────────────────
    print("\n[7] 后缀规范化")
    check("去掉路径分隔符", C.clean_output_suffix("_a/b.pdf") == "_ab.pdf",
          C.clean_output_suffix("_a/b.pdf"))
    check("去掉首尾空白与点", C.clean_output_suffix("  _x.  ") == "_x",
          C.clean_output_suffix("  _x.  "))
    check("限长 40", len(C.clean_output_suffix("_" + "x" * 60)) == 40)
    check("空后缀", C.clean_output_suffix("") == "")
    check("不做自动补下划线的猜测", C.clean_output_suffix("已替换") == "已替换")

    # ── [8] 多来源：产物跟着原件落在各自的目录 ───────────────────────────
    print("\n[8] 多来源：产物落在各自原件所在的目录")
    w9 = root / "多来源输出"
    src_g = w9 / "原件"
    (src_g / "批次B").mkdir(parents=True)
    for i in range(1, 3):
        build_pdf(src_g / f"{i:03d}.pdf", [(595, 842, f"A{i}-P1"), (595, 842, f"A{i}-P2")])
    for i in range(3, 5):
        build_pdf(src_g / "批次B" / f"{i:03d}.pdf",
                  [(595, 842, f"B{i}-P1"), (595, 842, f"B{i}-P2")])
    scan_g = w9 / "扫描件.pdf"
    build_pdf(scan_g, [(595, 842, f"SCAN-{k}") for k in range(1, 5)])
    # 顶层文件夹 + 子目录，两个来源一起加
    og = mkopts(sources=(str(src_g), str(src_g / "批次B")), scan_pdf=str(scan_g))
    pg = C.plan_batch(og)
    check("两个来源一起收，数量对得上（4 vs 4）", not pg.blocked, pg.block_reason)
    check("队列按文件名统一排（001..004，跟来自哪个文件夹无关）",
          [it.src.name for it in pg.items] == ["001.pdf", "002.pdf", "003.pdf", "004.pdf"],
          str([it.src.name for it in pg.items]))
    C.run_batch(pg, og)
    check("顶层原件：产物落在顶层", prod(src_g, "001.pdf").is_file(),
          str(sorted(str(p) for p in src_g.rglob("*.pdf"))))
    check("另一个来源（子目录）的原件：产物落在它自己那个目录里",
          prod(src_g / "批次B", "003.pdf").is_file(),
          str(sorted(str(p) for p in src_g.rglob("*.pdf"))))
    check("没有在别处凭空建输出文件夹", not (w9 / "输出").exists())

    # 第二次跑：产物被认出来，不会变成 8 份
    pg2 = C.plan_batch(og)
    check("第二次跑：产物被跳过，还是 4 份", len(pg2.items) == 4, str(len(pg2.items)))
    check("也如实报出跳过了几份", len(pg2.products) == 4, str(len(pg2.products)))

    # ── [8b] 只列队（没选扫描件）：不配对也能把队排出来 ──────────────────
    print("\n[8b] 只列队（还没选扫描件）")
    oq = mkopts(sources=(str(src_g), str(src_g / "批次B")))     # 故意不给 scan_pdf
    pq = C.plan_batch(oq)
    check("没有扫描件也能出队（不阻断）", not pq.blocked and len(pq.items) == 4,
          f"blocked={pq.blocked} items={len(pq.items)}")
    check("这是「只列了队」：paired 是 False", pq.paired is False)
    check("只列队时每行没有扫描页", all(it.scan_index == 0 for it in pq.items))
    check("只列队也把每份的页数读出来了", all(it.src_pages > 0 for it in pq.items))
    check("只列队时不算方向（全是 0）", all(it.rotation == 0 for it in pq.items))
    n_before = len(list(src_g.rglob(f"*{SUF}.pdf")))
    rq, notes = C.run_batch(pq, oq)
    check("没配对就直接跑 → 拒绝，一个文件都不写",
          not rq and notes and len(list(src_g.rglob(f"*{SUF}.pdf"))) == n_before,
          str(notes))
    # 选了扫描件之后同一批来源 → 变成配对态
    oq2 = mkopts(sources=(str(src_g), str(src_g / "批次B")), scan_pdf=str(scan_g))
    pq2 = C.plan_batch(oq2)
    check("选上扫描件后同一批来源变成配对态", pq2.paired and not pq2.blocked
          and all(it.scan_index for it in pq2.items))

    # ── [9] 演练不写文件 ─────────────────────────────────────────────────
    print("\n[9] dry_run")
    w10 = root / "演练"
    src_h, scan_h = make_corpus(w10, 2, 2, 2)
    oh = mkopts(sources=(str(src_h),), scan_pdf=str(scan_h), dry_run=True)
    ph = C.plan_batch(oh)
    rh, _ = C.run_batch(ph, oh)
    check("演练报告成功", all(r.ok for r in rh))
    check("演练不产生任何文件", not list(src_h.glob(f"*{SUF}.pdf")),
          str(sorted(p.name for p in src_h.glob("*.pdf"))))

    # ── [10] 跳过 ────────────────────────────────────────────────────────
    print("\n[10] 手动忽略")
    ph.items[1].status = C.Status.SKIPPED.value
    oi = mkopts(sources=(str(src_h),), scan_pdf=str(scan_h))
    ri, _ = C.run_batch(ph, oi)
    check("被忽略的那行不计入输出",
          [r.out.name for r in ri if r.out.name] == [f"001{SUF}.pdf"],
          str([r.out.name for r in ri]))

    # ── [11] ⑥ 人工判方向（横版表格被竖着扫进来） ────────────────────────
    print("\n[11] 扫描页方向（左转 / 右转）")
    w11 = root / "方向"
    src11 = w11 / "原件"
    src11.mkdir(parents=True)
    # 原件首页是横向 842×595（横版表格）；扫描页被竖着扫进来，是竖向 595×842
    build_pdf(src11 / "001.pdf", [(842, 595, "ORI1-P1"), (842, 595, "ORI1-P2")])
    scan11 = w11 / "扫描件.pdf"
    build_pdf(scan11, [(595, 842, "SCAN-1")])

    check("角度语义：正数=右转、负数=左转", (C.ROT_RIGHT, C.ROT_LEFT) == (90, -90))
    check("转四次回到原样",
          C.turn_deg(C.turn_deg(C.turn_deg(C.turn_deg(0, 90), 90), 90), 90) == 0)
    check("左转一次 = 270°", C.turn_deg(0, C.ROT_LEFT) == 270)
    check("90/270 宽高互换，180 不变",
          C.rotated_size((595, 842), 90) == (842.0, 595.0)
          and C.rotated_size((595, 842), 180) == (595.0, 842.0))

    # ⚠️ 这一节专门验证「不转」的对照，所以关掉⑦自动方向（否则横向原件会自动左转）
    base = mkopts(sources=(str(src11),), scan_pdf=str(scan11),
                     overwrite=True, auto_rotate=False)
    p_straight = C.plan_batch(base)
    check("不转时：预览里扫描页尺寸是竖的 595×842",
          p_straight.items[0].scan_size == (595.0, 842.0), str(p_straight.items[0].scan_size))
    r_straight, _ = C.run_batch(p_straight, base)
    check("不转也能出（只是套不进横版原件，原样保留）", r_straight[0].ok, r_straight[0].note)
    check("不转时输出首页仍是竖的 595×842",
          _disp(prod(src11, "001.pdf"))[0] == (595.0, 842.0), str(_disp(prod(src11, "001.pdf"))))
    check("不转时内容流里没有旋转矩阵",
          not _has_rot_matrix(prod(src11, "001.pdf")))

    rots = mkopts(sources=(str(src11),), scan_pdf=str(scan11),
                     rotations={1: 90}, overwrite=True)
    p_rot = C.plan_batch(rots)
    check("右转 90° 后预览里的尺寸就宽高互换了",
          p_rot.items[0].scan_size == (842.0, 595.0), str(p_rot.items[0].scan_size))
    check("方向标签是「右转 90°」",
          C.ROT_LABELS[p_rot.items[0].rotation] == "右转 90°", str(p_rot.items[0].rotation))
    r_rot, _ = C.run_batch(p_rot, rots)
    check("转完替换成功", r_rot[0].ok, r_rot[0].note)
    check("输出首页变成横向 842×595（正好贴合原件的横版首页）",
          _disp(prod(src11, "001.pdf"))[0] == (842.0, 595.0), str(_disp(prod(src11, "001.pdf"))))
    check("内容是真转了，不只是框变了（内容流里有旋转矩阵）",
          _has_rot_matrix(prod(src11, "001.pdf")))
    check("原件本身没被转（原件只读）",
          _disp(src11 / "001.pdf") == [(842.0, 595.0), (842.0, 595.0)])
    check("第 2 页起仍是原件内容",
          page_texts(prod(src11, "001.pdf"))[1:2] == ["ORI1-P2"], str(page_texts(prod(src11, "001.pdf"))))

    p_rot2 = C.plan_batch(mkopts(sources=(str(src11),), scan_pdf=str(scan11), rotations={1: 270}))
    check("左转 270° 同样是横向（宽高互换）",
          p_rot2.items[0].scan_size == (842.0, 595.0), str(p_rot2.items[0].scan_size))
    check("左转的标签对得上", C.ROT_LABELS[270] == "左转 90°")

    check("命令行能解析旋转设置", C.parse_rotations("3:90,5:270") == {3: 90, 5: 270})
    check("角度不是 90 的倍数会报错", _raises(lambda: C.parse_rotations("3:45")))
    check("页码格式写错会报错", _raises(lambda: C.parse_rotations("3")))

    # ── [12] ⑦ 自动方向（用户 2026-10-08 拍板的默认规则）────────────────
    print("\n[12] 自动方向：原件横向 → 它那一页默认左转 90°")
    w12 = root / "自动方向"
    src12 = w12 / "原件"
    src12.mkdir(parents=True)
    # 第 2 份是横向原件（842×595），其余竖向；扫描件一律竖扫 595×842
    for i, sz in ((1, (595, 842)), (2, (842, 595)), (3, (595, 842))):
        build_pdf(src12 / f"{i:03d}.pdf",
                  [(sz[0], sz[1], f"O{i}-P1"), (sz[0], sz[1], f"O{i}-P2")])
    scan12 = w12 / "扫描件.pdf"
    build_pdf(scan12, [(595, 842, f"S-{k}") for k in range(1, 4)])
    o12 = mkopts(sources=(str(src12),), scan_pdf=str(scan12), overwrite=True)

    check("原件首页横竖判得出来",
          [C.is_landscape(it.src_size) for it in
           C.plan_batch(o12).items] == [False, True, False])
    p12 = C.plan_batch(o12)
    check("横向原件那一页默认左转 90°（270），其余正向",
          [it.rotation for it in p12.items] == [0, 270, 0],
          str([it.rotation for it in p12.items]))
    check("预览里那一页的尺寸已经宽高互换",
          p12.items[1].scan_size == (842.0, 595.0), str(p12.items[1].scan_size))

    # 自动值跟着【行→页】映射走：页码一动，「这一页是谁的封面」变了，默认值跟着变。
    # 期望按**行**看：002 是横向原件，不管它改用哪一页，它那一行都得是左转 90°。
    C.swap_pages(p12, 2, -1)          # 第 2 行（横向原件）改用第 1 页
    C.apply_rotations(p12, {})
    check("页码对调后，横向原件那一行仍是左转 90°（跟原件走，不跟页码走）",
          [it.rotation for it in p12.items] == [0, 270, 0],
          str([it.rotation for it in p12.items]))
    check("对调后第 1 页归横向原件用 → 第 1 页被标成左转",
          p12.items[1].scan_index == 1 and p12.items[1].rotation == 270)

    # 人工 > 自动：给第 1 页指定右转 90°，压过「横向原件默认左转」
    C.apply_rotations(p12, {1: 90})
    check("人工判的方向压过自动值", p12.items[1].rotation == 90,
          str(p12.items[1].rotation))
    check("人工只管自己那一页（第 2 页仍走自动）", p12.items[0].rotation == 0,
          str(p12.items[0].rotation))
    check("反复 apply 不会把上一次的结果当输入（从原始尺寸重算）",
          C.plan_batch(o12).items[1].scan_size == (842.0, 595.0))

    p12b = C.plan_batch(mkopts(sources=(str(src12),), scan_pdf=str(scan12),
                                  auto_rotate=False))
    check("关掉自动方向后全部正向", [it.rotation for it in p12b.items] == [0, 0, 0])

    # 守卫：这一页「本来就是横向」→ 已经是横着扫出来的，不能再转一次
    scan12L = w12 / "扫描件_横.pdf"
    build_pdf(scan12L, [(842, 595, "S1"), (842, 595, "S2"), (842, 595, "S3")])
    p12L = C.plan_batch(mkopts(sources=(str(src12),), scan_pdf=str(scan12L)))
    check("扫描页本来就是横向 → 不再自动左转（否则会转反）",
          [it.rotation for it in p12L.items] == [0, 0, 0],
          str([it.rotation for it in p12L.items]))
    check("并在「说明」列出声：这批纸不是竖扫的",
          C.NOTE_SCAN_LANDSCAPE in p12L.items[1].note, p12L.items[1].note)

    # 界面③「整批转」的底层口径：在当前方向基础上再转 90°
    p12c = C.plan_batch(o12)
    C.apply_rotations(p12c, {it.scan_index: C.turn_deg(it.rotation, C.ROT_RIGHT)
                             for it in p12c.items})
    check("整批右转：横向那页 270→0，其余 0→90",
          [it.rotation for it in p12c.items] == [90, 0, 90],
          str([it.rotation for it in p12c.items]))

    # 真的跑一遍：自动左转要落到输出文件里
    o12d = mkopts(sources=(str(src12),), scan_pdf=str(scan12),
                     overwrite=True)
    r12, _ = C.run_batch(C.plan_batch(o12d), o12d)
    check("自动方向下全部替换成功", all(r.ok for r in r12), str([r.note for r in r12]))
    check("横向原件那份：输出首页是横向 842×595（转过来了）",
          _disp(prod(src12, "002.pdf"))[0] == (842.0, 595.0), str(_disp(prod(src12, "002.pdf"))))
    check("横向原件那份：内容是真的转了（内容流里有旋转矩阵）",
          _has_rot_matrix(prod(src12, "002.pdf")))
    check("竖向原件那份：输出首页仍是竖向 595×842（没被连坐）",
          _disp(prod(src12, "001.pdf"))[0] == (595.0, 842.0), str(_disp(prod(src12, "001.pdf"))))
    check("竖向原件那份：内容流里没有旋转矩阵",
          not _has_rot_matrix(prod(src12, "001.pdf")))

    # ══════════════════════════════════════════════════════════════════════
    #  [13] 格式转换：Word / Excel → PDF（v3.0 新增的第二套流程）
    #       ⚠️ 真引擎要么被 macOS 的 TCC 拦（-10004 权限违例）、要么这台没装，
    #          所以整条「列队 → 执行 → 回填」一律用**注入的假引擎**验。
    # ══════════════════════════════════════════════════════════════════════
    print("\n[13] ⑩ 格式表 / ⑪ 引擎注册表 / ⑫ 转换闸门")
    wf, ef = C.format_by_key("word"), C.format_by_key("excel")
    # ⭐ 凡是要「拼一段 Mac 脚本」的断言，都跑 **macOS 上真有应用** 的那批行
    #    （platform_formats("mac")）。现在它＝全部两行，但用派生值而不是直接遍历
    #    CONVERT_FORMATS，是为了将来若加进「只在 Windows 上有的格式」时这些断言
    #    不会被带偏 —— 那种行本来就不该拼 Mac 脚本。
    _mfs = C.platform_formats("mac")
    check("⑩ 表里 Word / Excel 两行都在", wf is not None and ef is not None)
    check("⑩ Word 认 .doc/.docx/.docm/.rtf 四种",
          set(wf.exts) == {".doc", ".docx", ".docm", ".rtf"}, str(wf.exts))
    check("⑩ Excel 认 .xls/.xlsx/.xlsm 三种",
          set(ef.exts) == {".xls", ".xlsx", ".xlsm"}, str(ef.exts))
    # ⚠️ 2026-10-09：这两个字段**只有 office 族用得上** —— CAD / BarTender 那一族
    #    走的是各自的专用程序（⑪ 表的 cad / bartender），没有 win_kind 也没有
    #    LibreOffice 过滤器。所以这里按族过滤，别再拿全表一刀切。
    _office_fs = [f for f in C.CONVERT_FORMATS if f.family == "office"]
    check("⑩ office 那两行都带了 Windows 应用名 / LibreOffice 过滤器（引擎靠它挑）",
          all(f.win_app and f.soffice_filter for f in _office_fs),
          str([(f.key, f.win_app, f.soffice_filter) for f in _office_fs]))
    check("⑩ macOS 上真有应用的那些行，都写了 Mac 应用名",
          all(f.mac_app for f in _mfs),
          str([(f.key, f.mac_app) for f in _mfs]))
    # ⭐⭐ 2026-10-08 的教训（AppleScript 里 tell application 必须跟字面量）
    #      在 2026-10-09 改成了**结构上不可能犯**：不再有 __APP__ 运行时替换，
    #      app 名由 build_mac_script() 直接拼进字面量。下面这几条盯死新结构。
    check("⭐ ⑩ 每行三步都齐（open / export / close，字段名由 OFFICE_STEPS 派生）",
          all(all(getattr(f, f"mac_{s}").strip() for s in C.OFFICE_STEPS)
              for f in _mfs),
          str([(f.key, [s for s in C.OFFICE_STEPS if not getattr(f, f"mac_{s}").strip()])
               for f in _mfs]))
    check("⭐⭐ 拼出来的脚本里 tell application 后面是**字面量**（写成变量 AppleScript 编译不过）",
          all(f'tell application "{f.mac_app}"' in C.build_mac_script(f)
              for f in _mfs)
          and all("tell application __APP__" not in C.build_mac_script(f)
                  for f in _mfs),
          C.build_mac_script(wf).strip().splitlines()[3])
    check("换出来的是这个 app 自己的名字（表里改了名字就跟着变）",
          '"Microsoft Word"' in C.build_mac_script(wf)
          and '"Microsoft Excel"' in C.build_mac_script(ef))
    # ⚠️ 步骤本身是多行的（mac_open 有十几行），拼装时每行统一加了缩进 →
    #    不能拿整段字符串去 index。改比每步的**第一句**，既精确又不怕缩进变。
    def _step_head(s: str) -> str:
        return s.splitlines()[0].strip()

    check("⭐⭐ 三步在脚本里的**先后顺序**就是 OFFICE_STEPS（不是随手写的）",
          all([C.build_mac_script(f).index(_step_head(getattr(f, f"mac_{s}")))
               for s in C.OFFICE_STEPS]
              == sorted(C.build_mac_script(f).index(_step_head(getattr(f, f"mac_{s}")))
                        for s in C.OFFICE_STEPS)
              for f in _mfs))
    check("⭐ Word / Excel **各用各的脚本**（一个 tell 块只认一个 app 的词典，混写编译不过）",
          C.build_mac_script(wf) != C.build_mac_script(ef)
          and "open workbook" in C.build_mac_script(ef)
          and "open workbook" not in C.build_mac_script(wf))

    # ⭐⭐ 2026-10-09（第二轮）：粒度过改 —— 一段脚本转**整批**。
    #    用户实测的两条症状都出自「一份文件一段脚本」：
    #      「每一份都要点一下授权」＋「3 份只成功 1 份 / 试三次才全成」。
    #    下面这几条把新粒度盯死（退回一份一段就会红）。
    _batch_body = C.build_mac_script(wf)
    check("⭐⭐ 生成的是**整批循环**脚本（一个格式只拉一次应用，不是一份一段）",
          "repeat with i from 1 to nFiles" in _batch_body
          and "set srcHFS to item i of srcList" in _batch_body
          and all(C.build_mac_script(f).count("repeat with i from 1 to nFiles") == 1
                  for f in _mfs))
    check("⭐⭐ `activate` 全脚本只出现**一次**、且在循环外面"
          "（每份都激活一次＝窗口乱跳 + 反复问授权）",
          all(C.build_mac_script(f).count("activate") == 1 for f in _mfs)
          and _batch_body.index("activate")
          < _batch_body.index("repeat with i from 1 to nFiles"),
          str([(f.key, C.build_mac_script(f).count("activate")) for f in _mfs]))
    check("⭐⭐ 三步都在那个循环**里面**（循环起点之后）",
          all(C.build_mac_script(f).index("repeat with i from 1 to nFiles")
              < C.build_mac_script(f).index(_step_head(getattr(f, f"mac_{s}")))
              for f in _mfs for s in C.OFFICE_STEPS))
    check("⭐ 循环上界在 tell 块**外面**算好"
          "（免得在 app 词典作用域里对脚本变量求 count）",
          all(C.build_mac_script(f).index("set nFiles to (count of srcList)")
              < C.build_mac_script(f).index(f'tell application "{f.mac_app}"')
              for f in _mfs))
    check("⭐⭐ 脚本里**不许再出现 srcPath** —— 那是「一份一段」时代的变量，"
          "批量化后它不存在了，引用它就是 -2753 复发",
          all("srcPath" not in C.build_mac_script(f) for f in _mfs))
    check("⭐ 一份失败只记一行错误、后面的照样转（整批模式带来的额外好处）",
          "on error errMsg number errNum" in _batch_body
          and "|ERR|" in _batch_body and "my neat(errMsg)" in _batch_body)
    check("⭐ 结果按「一行一份」交回（结尾把列表拼成文本），Python 才解析得动",
          "set AppleScript's text item delimiters to linefeed" in _batch_body
          and "return out as text" in _batch_body)
    try:
        C.assert_format_complete(wf._replace(mac_export=""))
        _miss = False
    except ValueError:
        _miss = True
    check("⑩ 表里哪行漏了 macOS 的某一步 → 当场抛错（不变成一个看不懂的 osascript 语法错）",
          _miss)
    try:
        C.assert_format_complete(wf._replace(win_kind="ppt"))
        _miss2 = False
    except ValueError:
        _miss2 = True
    check("⑩ 表里 win_kind 写了个没有实现的 → 当场抛错（Windows 侧不会跑到一半才发现）",
          _miss2)
    check("⭐ 两个平台覆盖同一套 OFFICE_STEPS（Mac 三步字段 / Win 的 win_kind 都在表里）",
          set(C.OFFICE_STEPS) == {"open", "export", "close"}
          and all(f.win_kind in C.WIN_OFFICE_RUNNERS for f in _office_fs),
          str([(f.key, f.win_kind) for f in _office_fs]))

    # ⭐⭐ 2026-10-09（第六轮）：CAD / BarTender —— 「只在一个平台上成立」这件事
    #     **真有人用了**（用户点名要的那两种；Mac 上根本没有对应软件，也用不到）。
    #     所以不再拿合成格式钉机制了，直接拿真行验；再补一条**合成格式**验
    #     「非 office 族的 Office 字段必须全空」—— 表里没有反例，得造一个。
    cf, bf = C.format_by_key("cad"), C.format_by_key("btw")
    check("⑩ 表里有 CAD 与 BarTender 那一行（用户 2026-10-09 点名要的）",
          cf is not None and bf is not None)
    check("⭐ ⑩ CAD 把那几种图纸后缀都认下（.dwg/.dxf 是常用的，.dwt/.dws 也一并认）",
          set(cf.exts) == {".dwg", ".dxf", ".dwt", ".dws"}, str(cf.exts))
    check("⭐ ⑩ BarTender 只认 .btw（.btxml 是 XML 交换格式，不是标签文档）",
          set(bf.exts) == {".btw"}, str(bf.exts))
    check("认扩展名大小写不敏感（.DWG 也认）", C.format_of(Path("图.DWG")) is cf)
    check("⭐⭐ CAD / BarTender 都**只在 Windows 上**成立（platforms=(\"win\",)）",
          cf.platforms == ("win",) and bf.platforms == ("win",),
          f"{cf.platforms} / {bf.platforms}")
    check("⭐⭐ 它们那一行的 Office 字段**全空**（不是 Office 文档，不走那套三步曲）",
          not any([cf.mac_app, cf.win_app, cf.win_kind, cf.soffice_filter,
                   *[getattr(cf, f"mac_{s}").strip() for s in C.OFFICE_STEPS]])
          and not any([bf.mac_app, bf.win_app, bf.win_kind, bf.soffice_filter,
                       *[getattr(bf, f"mac_{s}").strip() for s in C.OFFICE_STEPS]]))
    check("⭐⭐ 非 office 族若**填了** Office 字段 → 当场炸（表自己矛盾）",
          _raises(lambda: C.assert_format_complete(cf._replace(win_app="Acad.Application"))))
    check("⭐⭐ Office 族若**说 Mac 上能转、却没有 Mac 应用名** → 当场炸",
          _raises(lambda: C.assert_format_complete(ef._replace(mac_app=""))))
    check("⭐⭐ platforms 写空 → 当场炸（没有哪个格式两边都转不了）",
          _raises(lambda: C.assert_format_complete(ef._replace(platforms=()))))
    check("⭐ family 写了个不存在的 → 当场炸",
          _raises(lambda: C.assert_format_complete(ef._replace(family="不存在"))))
    check("⭐ 每行都带上了 family 与 platforms（将来加格式别忘写）",
          all(f.family in C.ENGINE_FAMILIES and f.platforms for f in C.CONVERT_FORMATS),
          str([(f.key, f.family, f.platforms) for f in C.CONVERT_FORMATS]))
    check("⭐ 在 Mac 上给 CAD 拼脚本 → 当场拦下（不去拼一段不存在的脚本）",
          _raises(lambda: C.build_mac_script(cf)))
    _cnote = C.mac_unavailable_note(cf)
    check("⭐⭐ Mac 上碰到 CAD，说的是「去 Windows 上转」而不是「装一个」"
          "（Mac 上根本没有它可装）",
          "CAD" in _cnote and "Windows" in _cnote and "macOS" in _cnote
          and C.family_maker("cad") in _cnote, _cnote)
    check("  ↳ 而 Word / Excel 这种两边都有的格式，这句话是空的（不打扰）",
          C.mac_unavailable_note(wf) == "" and C.mac_unavailable_note(ef) == "")
    check("⭐⭐ 平台派生把「只在 Windows 上有的格式」挡在 Mac 那侧外面"
          "（界面就不会在 Mac 上说支持它）",
          [f.key for f in C.platform_formats("mac")] == ["word", "excel"]
          and {"cad", "btw"} <= {f.key for f in C.platform_formats("win")},
          str([f.key for f in C.platform_formats("mac")]))
    check("⭐ 界面上的「支持哪些类型」按表 + 平台派生：Mac 上不提它、Windows 上提它",
          C.convert_kind_words("mac") == "Word / Excel"
          and C.convert_kind_words("win") == "Word / Excel / CAD / BarTender",
          C.convert_kind_words("mac") + " ／ " + C.convert_kind_words("win"))
    check("⭐ 分页规则那句话也按平台派生（Mac 上不会冒出 CAD / 标签的规则）",
          ("图纸 / 标签按它自己排的页" in C.page_rule_label())
          == (sys.platform != "darwin"),
          C.page_rule_label())
    if sys.platform == "darwin" and shutil.which("osacompile"):
        _badc = []
        for f in _mfs:
            _r = subprocess.run(["osacompile", "-o", "/dev/null", "-e",
                                 C.build_mac_script(f)],
                                capture_output=True, text=True)
            if _r.returncode != 0:
                _badc.append((f.key, (_r.stderr or "").strip()[:80]))
        check("⭐⭐ 生成的 AppleScript **在这台机器上真能编译**（-2741 那段坑的实机回归）",
              not _badc, str(_badc))

    # ⭐⭐ 2026-10-09（第五轮）：⑩c 分页规则 —— 「一份文件该出几页」。
    #     用户给的规律（配 WI-2019 实例：13 个子表格 ↔ 13 页，一一对应）：
    #     **Excel 底下每个工作表 = 一页**；Word 按自己的分页一页一页。
    #     规则在 ⑩ 表里声明一次，能核对页数的（Excel）顺带核对 —— 不符才出声。
    print("\n[13a5] ⑩c 分页规则：Excel 一表一页 / 规则声明 + 页数核对")
    check("⑩c 分页规则表现在三档（flow / sheet / own）",
          set(C.PAGE_RULES) == {"flow", "sheet", "own"}, str(list(C.PAGE_RULES)))
    check("⭐ ⑩ 每行都指到 ⑩c 里一条**真**规则（Word 分页 / Excel 一表一页 / 图纸标签按自己）",
          wf.page_rule == "flow" and ef.page_rule == "sheet"
          and cf.page_rule == "own" and bf.page_rule == "own"
          and all(f.page_rule in C.PAGE_RULES for f in C.CONVERT_FORMATS),
          str([(f.key, f.page_rule) for f in C.CONVERT_FORMATS]))
    check("⭐ ⑩ 表里 page_rule 写了个不存在的 → 当场抛错（不拖到运行期）",
          _raises(lambda: C.assert_format_complete(ef._replace(page_rule="不存在"))))
    check("⭐ 只有 Excel 那条规则报得出「应该几页」"
          "（Word 报不出；图纸 / 标签也报不出 —— 它们几页由软件自己排，绝不瞎猜）",
          C.page_rule_of(ef).expect is not None
          and C.page_rule_of(wf).expect is None
          and C.page_rule_of(cf).expect is None
          and C.page_rule_of(bf).expect is None)
    check("⭐ 规则那句话是按平台 + ⑩ 表演化出来的",
          "工作表一页" in C.page_rule_label() and "分页一页一页" in C.page_rule_label(),
          C.page_rule_label())

    _prdir = Path(tempfile.mkdtemp(prefix="pagerule_"))
    _x3 = _prdir / "三表全可见.xlsx"
    write_min_xlsx(_x3, [("封面", "visible"), ("工序", "visible"), ("附表", "visible")])
    _xhide = _prdir / "两可见一隐藏.xlsx"
    write_min_xlsx(_xhide, [("A", "visible"), ("B", "hidden"), ("C", "visible")])
    check("⭐ 数得出 .xlsx 里几个**可见**工作表（Excel 导 PDF 只导可见的那些）",
          C._xlsx_sheet_count(_x3) == 3 and C._xlsx_sheet_count(_xhide) == 2,
          f"{C._xlsx_sheet_count(_x3)} / {C._xlsx_sheet_count(_xhide)}")
    check("老 .xls（二进制）读不了 → None（这一份就不核对，不瞎猜一个数）",
          C._xlsx_sheet_count(_prdir / "老格式.xls") is None)
    check("文件不存在 → None", C._xlsx_sheet_count(_prdir / "没有这个.xlsx") is None)
    _junkx = _prdir / "坏文件.xlsx"
    _junkx.write_bytes(b"not a zip at all - fake xlsx")
    check("⭐ 坏文件 / 不是 zip → None（核对绝不抛异常、绝不拖垮转换）",
          C._xlsx_sheet_count(_junkx) is None)
    check("⭐⭐ 页数**正对上**规则 → 一个字都不说（用户嫌啰嗦）",
          C.page_rule_report(ef, _x3, 3) == "")
    _pwarn = C.page_rule_report(ef, _x3, 2)
    check("⭐⭐ 页数**对不上** → 给一句人话（说清几个工作表、转出几页）",
          "3 个工作表" in _pwarn and "2 页" in _pwarn, _pwarn)
    check("Word 那条规则没有「应该几页」，所以永远不报（不打扰）",
          C.page_rule_report(wf, _x3, 9) == "")
    check("页数读不出来（0）时不报 —— 那是另一条提示管的事",
          C.page_rule_report(ef, _x3, 0) == "")

    check("⭐ CONVERT_EXTS 是从表里派生出来的（绝不手写第二份，免得脱节）",
          set(C.CONVERT_EXTS) == {e for f in C.CONVERT_FORMATS for e in f.exts}
          and list(C.CONVERT_EXTS) == sorted(C.CONVERT_EXTS),
          str(C.CONVERT_EXTS))
    check("认扩展名不区分大小写（.XLSX 也认）", C.format_of(Path("a.XLSX")) is ef)
    check("认不出的格式返回 None，不抛异常",
          C.format_of(Path("a.pdf")) is None and C.format_of(Path("无扩展名")) is None)
    check("格式表的 key 查不到也返回 None", C.format_by_key("pdf") is None)

    # ⑪ 引擎表的**顺序**就是优先级：保真度高的排前面（Office 认文件自带的打印设置）
    keys13 = [e.key for e in C.CONVERT_ENGINES]
    check("⑪ 引擎表里 Office 排在 LibreOffice 前面（保真度优先）",
          keys13.index("libreoffice") > max(i for i, k in enumerate(keys13)
                                            if k.startswith("office")), str(keys13))
    check("⑪ 每行都有 probe / run，且都可调用",
          all(callable(e.probe) and callable(e.run) for e in C.CONVERT_ENGINES), str(keys13))
    check("⭐ 本机 Office 那两条都声明了**批量口**（整批塞进一个应用会话）",
          all(C.engine_by_key(k).run_many is not None
              for k in ("office_mac", "office_win")),
          str([(e.key, e.run_many is not None) for e in C.CONVERT_ENGINES]))
    check("LibreOffice 是命令行引擎：没有批量口也不算错（流程会退回按单份循环）",
          C.engine_by_key("libreoffice").run_many is None)
    check("⑪ 引擎的说明文案里讲清了「一次只拉起一次应用」",
          "一次" in C.engine_by_key("office_mac").detail
          and "整批" in C.engine_by_key("office_win").detail,
          C.engine_by_key("office_mac").detail[:40])
    check("⑪ 每行都写清了「用得上什么」的说明（界面拿它当悬浮提示）",
          all(e.detail for e in C.CONVERT_ENGINES), str(keys13))
    check("⭐ 没引擎时**只**让装 Word / Excel（不指路去装第三方，免得使用者犯迷糊）",
          "Office" in C.engine_no_engine_msg()
          and "LibreOffice" not in C.engine_no_engine_msg(),
          C.engine_no_engine_msg())
    check("⭐ 失败补救文案里也不再劝人装别的转换软件（2026-10-09 用户拍板）",
          not any("LibreOffice" in v for v in C.OFFICE_ERR_HINTS.values()),
          str([k for k, v in C.OFFICE_ERR_HINTS.items() if "LibreOffice" in v]))

    # ══════════════════════════════════════════════════════════════════════
    #  [13b2] ⑪ 引擎的「平台 + 族」
    #     ⭐ 用户 2026-10-09 的疑问：「Mac 版里是不是装着一个 Windows 引擎？
    #        Windows 上跑的时候，Word 是不是先让 Mac 那条试一遍？」
    #     —— 从来没有（平台不符的引擎探测直接 False），但**表上看不出来**。
    #     所以把「平台」和「族」写到表里，并把下面这些规矩钉死。
    # ══════════════════════════════════════════════════════════════════════
    print("\n[13b2] ⑪ 引擎的「平台 + 族」—— 一台机器上只走自己那条路")
    check("⭐ 每条引擎都标了 platform（mac / win / any）",
          all(e.platform in ("mac", "win", "any") for e in C.CONVERT_ENGINES),
          str([(e.key, e.platform) for e in C.CONVERT_ENGINES]))
    check("⭐ 每条引擎都标了 family，且都在 ⑩a 族表里",
          all(e.family in C.ENGINE_FAMILIES for e in C.CONVERT_ENGINES),
          str([(e.key, e.family) for e in C.CONVERT_ENGINES]))
    check("⭐⭐ Mac 上**根本不会列出** Windows 的引擎（不是「列出来了再判 False」）",
          [e.key for e in C.active_engines("mac")] == ["office_mac", "libreoffice"],
          str([e.key for e in C.active_engines("mac")]))
    check("⭐⭐ Windows 上也不会列出 Mac 的那条，但会多出 CAD / BarTender 两条",
          {e.key for e in C.active_engines("win")}
          == {"office_win", "libreoffice", "cad", "bartender"},
          str([e.key for e in C.active_engines("win")]))
    check("⭐⭐ CAD / BarTender 在 Mac 上**一条引擎都没有**（列队就会被拦下并说明白）",
          C.family_engines("cad", "mac") == [] and C.family_engines("btw", "mac") == []
          and [e.key for e in C.family_engines("cad", "win")] == ["cad"]
          and [e.key for e in C.family_engines("btw", "win")] == ["bartender"])
    check("⭐⭐ 兜底梯队**只在同族内**换：CAD 那条不会拉上 LibreOffice / Word 来顶",
          [e.key for e in C.engine_chain(C.engine_by_key("cad"))] == ["cad"],
          str([e.key for e in C.engine_chain(C.engine_by_key("cad"))]))
    if sys.platform == "darwin":
        _oc = [e.key for e in C.engine_chain(C.engine_by_key("office_mac"))]
        _lo_ok = C.probe_engine(C.engine_by_key("libreoffice"))
        check("⭐⭐ Office 那族该拉的还是拉（同族内兜底没被这次改动带坏；"
              "本机没装 LibreOffice 就只有它自己）",
              _oc[0] == "office_mac"
              and all(C.engine_by_key(k).family == "office" for k in _oc)
              and (("libreoffice" in _oc) == _lo_ok),
              f"{_oc}（本机 LibreOffice 可用={_lo_ok}）")
    check("⭐ 引擎的名字里**不出现别族的格式**（office 那条不许自称能转 CAD）",
          C.engine_by_key("office_win").label == "Microsoft Word / Excel（本机）"
          and C.engine_by_key("office_mac").label == "Microsoft Word / Excel（本机）",
          C.engine_by_key("office_win").label)
    check("⭐ 引擎的「说人话」名字也只数自己那一族（加格式时不会串到别的引擎头上）",
          C.family_kind_words("office", "win") == "Word / Excel"
          and C.family_kind_words("cad", "win") == "CAD"
          and C.family_kind_words("cad", "mac") == "",
          f'{C.family_kind_words("office", "win")} / {C.family_kind_words("cad", "win")}')
    _cad_msg = C.engine_no_engine_msg("cad")
    check("⭐⭐ CAD 没引擎时说的是「去装 CAD（浩辰 CAD）/ 去 Windows 上转」，"
          "不是笼统的「装个软件」——品牌名由 ⑩d 品牌表派生，文案里不许再手写一遍",
          C.cad_brand_words() in _cad_msg and "Windows" in _cad_msg
          and "Office" not in _cad_msg, _cad_msg)
    check("⭐ families_of 保序去重（一族只挑一条引擎，同一族的前后不重复）",
          C.families_of([ef, wf, ef]) == ["office"]
          and C.families_of([wf, cf, bf]) == ["office", "cad", "btw"])

    # ══════════════════════════════════════════════════════════════════════
    #  [13b3] ⑩d CAD 品牌表 + 调用层 —— 「到底能不能信」
    #     ⚠️ 这一段**在这台机器上跑不了真的**（Mac 上装不了 CAD / BarTender，
    #        用户公司那台装的是**浩辰 CAD 2025**，要拷到同事的 Windows 上试）。
    #        所以这里只钉**结构**：牌子认不认得、设备是不是问出来的、签名对不对、
    #        产物核对在不在、探测会不会瞎试。
    #     ⭐ 2026-10-10 重写：原来钉的是「accoreconsole 的 17 行 -PLOT 应答表」，
    #        用户一句「其实是浩辰 CAD 2025」那套就全废了 —— 现在钉的是**表 + COM**。
    # ══════════════════════════════════════════════════════════════════════
    print("\n[13b3] ⑩d CAD 品牌表 / 调用层的结构")
    check("⭐⭐ CAD 靠 **COM 名字**认程序，不是猜安装路径"
          "（实测：浩辰装在 `C:\\Program Files\\浩辰软件\\…`，按路径 glob 必然漏）",
          bool(C.CAD_APPS)
          and all(a.progids and all("." in p for p in a.progids) for a in C.CAD_APPS)
          and not any(("\\" in p or "/" in p or p.endswith(".exe"))
                      for a in C.CAD_APPS for p in a.progids),
          str([(a.key, a.progids) for a in C.CAD_APPS]))
    check("⭐⭐ 品牌表里**只剩浩辰一家**（2026-10-10 用户拍板：AutoCAD 那一行、"
          "中望 / BricsCAD / nanoCAD 三条注释预留位全部删掉）",
          [a.key for a in C.CAD_APPS] == ["gstarcad"] and "浩辰" in C.CAD_APPS[0].name,
          str([(a.key, a.name) for a in C.CAD_APPS]))
    check("⭐ 浩辰的 ProgID 收全了（`GStarCAD` 大写 S 是注册表里实际登记的写法 —— "
          "用户 `reg query HKCR /f \"GStarCAD\" /k` 实测过）",
          {"GStarCAD.Application", "GStarCAD.Application.25", "GstarCAD.Application"}
          <= set(C.CAD_APPS[0].progids),
          str(C.CAD_APPS[0].progids))
    check("⭐ 品牌文案由表派生（加一家品牌只加一行；别处不许再手写品牌名）——"
          "⑩a 的 FAMILY_MAKER 与 ⑩d 品牌表**不许各说各的**",
          C.cad_brand_words() == " / ".join(a.name for a in C.CAD_APPS)
          and C.cad_brand_words() == "浩辰 CAD"
          and C.family_maker("cad") == C.cad_brand_words(),
          f"{C.cad_brand_words()} ／ family_maker={C.family_maker('cad')}")
    check("⭐⭐ 打印设备**从本机问、不写死**（写死 `DWG To PDF.pc3` = 换台机器就错）："
          "按 ⑩d2 关键词挑，挑不出就老实说挑不出",
          C.pick_pdf_device(["HP LaserJet 1020", "DWG To PDF.pc3"]) == "DWG To PDF.pc3"
          and C.pick_pdf_device(["浩辰 PDF 打印机"]) == "浩辰 PDF 打印机"
          and C.pick_pdf_device(["Microsoft XPS Document Writer"]) is None
          and C.pick_pdf_device([]) is None
          and C.pick_pdf_device(None) is None)
    check("⭐ 挑设备时**关键词表的顺序就是优先级**（先用 CAD 自带的出图驱动）",
          C.pick_pdf_device(["Acme PDF Printer", "DWG To PDF.pc3"]) == "DWG To PDF.pc3",
          str(C.CAD_PDF_DEVICE_WORDS))
    check("⭐⭐ 出图前**必须关掉后台打印**（BACKGROUNDPLOT=0）—— "
          "开着的话 PlotToFile 会提前返回，**写了一半的 PDF** 会被当成成功",
          ("BACKGROUNDPLOT", 0) in C.CAD_QUIET_VARS, str(C.CAD_QUIET_VARS))
    # ── 2026-10-10 真机 bug：只打 `ActiveLayout` 一页 → 有参照那张布局漏打 ──
    #    `_cad_layouts` 是纯函数（只吃 doc 对象），用假 doc 在 Mac 上就能盯住
    #    「跳过模型 + 按 TabOrder 排序」这套口径，真机才验 `ActiveLayout = lay` 那一行。
    class _Lay:
        def __init__(self, name, tab):
            self.Name = name
            self.TabOrder = tab
    class _Layouts:
        def __init__(self, items):
            self._items = items
        @property
        def Count(self):
            return len(self._items)
        def Item(self, i):
            return self._items[i]
    class _Doc:
        def __init__(self, items):
            self.Layouts = _Layouts(items)
    check("⭐⭐ CAD 出图**枚举所有布局、跳过模型空间**（2026-10-10 真机发现："
          "只打 ActiveLayout 会把「有参照那张布局」漏掉）—— 纯函数按 TabOrder 排序、"
          "模型空间不计入、布局名里有无数字都只看 TabOrder",
          C._cad_layouts(_Doc([_Lay("Model", 0), _Lay("布局2", 2), _Lay("布局1", 1)]))
          == [(1, "布局1", 2), (2, "布局2", 1)]
          and C._cad_layouts(_Doc([_Lay("Model", 0)])) == []
          and C._cad_layouts(_Doc([])) == []
          and C._cad_layouts(_Doc([_Lay("Layout1", 1), _Lay("Layout2", 2)]))
          == [(1, "Layout1", 0), (2, "Layout2", 1)],
          str(C._cad_layouts(_Doc([_Lay("Model", 0), _Lay("布局2", 2), _Lay("布局1", 1)]))))
    # ⭐ 反向哨兵：源码里不许再出现「直接打当前布局一页」的旧写法。
    #    旧 bug 就两行：`layout = doc.ActiveLayout` + `PlotToFile`，一次只打一页。
    #    现在必须先 `_cad_layouts()` 枚举、再逐页 PlotToFile —— 抓源码盯着「枚举在出图前」。
    _plot_src = inspect.getsource(C._plot_one)
    check("⭐⭐ `_plot_one` 先枚举布局再出图（不许退回去只打 ActiveLayout 一页）",
          "_cad_layouts(doc)" in _plot_src
          and "_cad_layouts" in _plot_src
          and "doc.ActiveLayout = lay" in _plot_src
          and "PdfWriter" in _plot_src,
          "（源码里缺 枚举/切布局/拼多页 之一）")
    # ⭐ 反向哨兵（2026-10-10 第三十波语义翻转）：**禁止**在 `_plot_one` 里 `lay.ConfigName = dev`
    #    换设备 —— 换设备会把布局的纸张/PlotOrigin/StandardScale 全套重置（真机铁证：
    #    PlotOrigin (13.27,0)→(0,54.18) → 内容顶左上、白边 54mm），连续七波「读回设回」都治不干净。
    #    ⭐ 正路 = 跟「方向」那次成功同一招：**别动布局一个字**，把 PDF 设备名当
    #    `PlotToFile(plotFile, plotConfig)` 的第二参数传（Autodesk 官方 Solved 写法）。
    check("⭐⭐ `_plot_one` 不许换布局设备（改走 PlotToFile 第二参数传 PDF 设备名）",
          "lay.ConfigName = dev" not in _plot_src
          and "lay.ConfigName" not in _plot_src
          and "PlotToFile(str(page_pdf), dev)" in _plot_src,
          "（换设备会污染布局的纸张/原点/比例，白边就是这么来的）")
    check("⭐⭐ `_plot_one` 不再读回设回任何打印属性（读回设回=承认换设备会污染）",
          "_plot_restore" not in _plot_src
          and "orig_pw" not in _plot_src
          and "lay.PlotOrigin" not in _plot_src
          and "lay.CenterPlot" not in _plot_src,
          "（出现读回设回，说明又退回「换设备再救」的老路）")
    # ── 2026-10-10 方向诊断：`_cad_fix_orientation` 只读不写，按 AcPlotRotation 枚举判 ──
    #    PlotRotation 语义（AutoCAD AcPlotRotation / 浩辰兼容）：
    #        0 = ac0degrees  = 0°   = 横
    #        1 = ac90degrees  = 90°  = 竖
    #        2 = ac180degrees = 180° = 横
    #        3 = ac270degrees = 270° = 竖
    #    ⭐ 2026-10-10 真机踩坑后改口径：出图**不碰方向**（浩辰自己按存的设置出），
    #       这里只读出来判断「纸张横竖 vs 当前方向」是否一致，供诊断看。
    class _FakeLay:
        def __init__(self, pw, ph, rot):
            self._pw, self._ph, self.rot = pw, ph, rot
        def GetPaperSize(self):
            return self._pw, self._ph
        @property
        def PlotRotation(self):
            return self.rot
        @PlotRotation.setter
        def PlotRotation(self, v):
            self.rot = v
        def RefreshPlotDeviceInfo(self):
            pass

    # 真机横纸那份：GetPaperSize=(297,210) 横纸 + PlotRotation=0（枚举=横）→ 一致
    _l1 = _FakeLay(297, 210, 0.0)
    _r1 = C._cad_fix_orientation(_l1)
    check("⭐⭐ 横纸 + PlotRotation=0（枚举=横）→ 方向一致，'ok'（且**不写**）",
          _r1 == "ok" and _l1.PlotRotation == 0.0,
          f"状态={_r1}, 方向={_l1.PlotRotation}")

    # 真机竖纸那份：GetPaperSize=(210,297) 竖纸 + PlotRotation=1（枚举=竖）→ 一致
    _l2 = _FakeLay(210, 297, 1.0)
    _r2 = C._cad_fix_orientation(_l2)
    check("⭐⭐ 竖纸 + PlotRotation=1（枚举=竖）→ 方向一致，'ok'（且**不写**）",
          _r2 == "ok" and _l2.PlotRotation == 1.0,
          f"状态={_r2}, 方向={_l2.PlotRotation}")

    # 横纸但当前方向是竖（枚举=1/90°）→ 不一致，报 mismatch（不掰）
    _l3 = _FakeLay(297, 210, 1.0)
    _r3 = C._cad_fix_orientation(_l3)
    check("⭐ 横纸 + PlotRotation=1（竖）→ 不一致，'mismatch'（只报不掰）",
          _r3 == "mismatch" and _l3.PlotRotation == 1.0,
          f"状态={_r3}, 方向={_l3.PlotRotation}")

    # 竖纸但当前方向是横（枚举=0）→ 不一致
    _l4 = _FakeLay(210, 297, 0.0)
    _r4 = C._cad_fix_orientation(_l4)
    check("⭐ 竖纸 + PlotRotation=0（横）→ 不一致，'mismatch'（只报不掰）",
          _r4 == "mismatch" and _l4.PlotRotation == 0.0,
          f"状态={_r4}, 方向={_l4.PlotRotation}")

    # 180° 也算横（枚举=2）→ 横纸 + 180 = 一致
    _l5 = _FakeLay(297, 210, 2.0)
    _r5 = C._cad_fix_orientation(_l5)
    check("⭐ 横纸 + PlotRotation=2（180°=横）→ 一致，'ok'",
          _r5 == "ok" and _l5.PlotRotation == 2.0,
          f"状态={_r5}, 方向={_l5.PlotRotation}")

    # GetPaperSize 读不到 → 没法比
    class _BadLay:
        def GetPaperSize(self):
            raise RuntimeError("读不到纸张")
    check("⭐ GetPaperSize 读不到 → 返回 'paper'（不瞎判）",
          C._cad_fix_orientation(_BadLay()) == "paper",
          "（读不到纸张尺寸就别乱判方向）")

    # PlotRotation 读不到 → 没法比
    class _NoRotLay:
        def GetPaperSize(self):
            return 297.0, 210.0
        @property
        def PlotRotation(self):
            raise RuntimeError("读不到方向")
    check("⭐ PlotRotation 读不到 → 返回 'rot'（不瞎判）",
          C._cad_fix_orientation(_NoRotLay()) == "rot",
          "（读不到方向就别乱判）")

    # ── 2026-10-10 第二十二波：`_cad_match_media` 按尺寸找回被 ConfigName 塌掉的纸张 ──
    #    真机 MediaBox 铁证：换 "DWG To PDF.pc3" 后横 A4(297×210) 塌成竖 A4(210×297)。
    class _MediaLay:
        def __init__(self, medias):
            self.medias = medias                 # [(纸名, 宽, 高)]
            self.canon = medias[0][0]
        def GetCanonicalMediaNames(self):
            return tuple(m[0] for m in self.medias)
        @property
        def CanonicalMediaName(self):
            return self.canon
        @CanonicalMediaName.setter
        def CanonicalMediaName(self, v):
            self.canon = v
        def RefreshPlotDeviceInfo(self):
            pass
        def GetPaperSize(self):
            for n, w, h in self.medias:
                if n == self.canon:
                    return w, h
            raise RuntimeError("no paper")

    # 正匹配：清单里就有 (297,210) 正着的那张 → 优先选它、不用反向
    _m1 = _MediaLay([("A4_竖", 210.0, 297.0), ("A4_横", 297.0, 210.0)])
    _r = C._cad_match_media(_m1, 297.0, 210.0)
    check("⭐⭐ 按尺寸找回纸张：清单里有正着的 (297,210) → 选中且不反向",
          _r is not None and _r[0] == "A4_横" and _r[1] is False,
          f"结果={_r}")

    # 反向兜底：清单里只有 (210,297)，图纸要 (297,210) → 选竖纸、标记反向
    _m2 = _MediaLay([("A4_竖", 210.0, 297.0)])
    _r2 = C._cad_match_media(_m2, 297.0, 210.0)
    check("⭐ 只有反向的纸 → 选它并标记「反向」（出图会翻转方向补偿）",
          _r2 is not None and _r2[0] == "A4_竖" and _r2[1] is True,
          f"结果={_r2}")

    # 一张都对不上 → None（保持默认，别瞎设）
    _m3 = _MediaLay([("A2", 594.0, 420.0)])
    check("⭐ 清单里没有同尺寸的纸 → 返回 None（不瞎设）",
          C._cad_match_media(_m3, 297.0, 210.0) is None,
          "（匹配不到就保持 PC3 默认，别设个错的）")

    # ⭐ 反向哨兵（2026-10-10 第二十七波）：-PUBLISH 路已弃用 —— 浩辰不认命令行 DSD，
    #    会弹「发布」对话框等人点、卡死批量。出图只走 COM 路，`_plot_one` 里不能再有
    #    -PUBLISH 的痕迹。
    check("⭐⭐ -PUBLISH 路已弃用（浩辰不认命令行 DSD，弹窗卡死）",
          not hasattr(C, "_cad_plot_via_publish")
          and not hasattr(C, "_cad_write_dsd")
          and "CAD_PUBLISH_FIRST" not in inspect.getsource(C._plot_one)
          and "PUBLISH" not in inspect.getsource(C._plot_one),
          "（浩辰命令行 -PUBLISH 弹窗等人点，出图只能走 COM，别再挂这条路）")
    def _first3_ok(fn, want: tuple[str, ...]) -> bool:
        """`fn` 的前几个位置参数必须正好是 want，多出来的**必须能省**（有默认值）。

        ⭐ 这就是「`eng.run(src, dst, fmt)` 一定调得通」的形式化写法 ——
           比对参数个数更有意义：Office 那几条后面还挂着一个可选 `timeout`，合法；
           而上一版 CAD 写成 `(src, dst)` 少一个，就必炸。
        """
        ps = list(inspect.signature(fn).parameters.values())
        if [p.name for p in ps[:len(want)]] != list(want):
            return False
        return all(p.default is not inspect.Parameter.empty for p in ps[len(want):])

    check("⭐⭐ 每条引擎的 `run` 前三个参数必须正好是 **(src, dst, fmt)**"
          "（转换层统一三参调用；写成两参＝真机上必 TypeError，"
          "而注入假引擎的回归**根本走不到**那条路 —— 2026-10-10 真踩过）",
          all(_first3_ok(e.run, ("src", "dst", "fmt")) for e in C.CONVERT_ENGINES),
          str([(e.key, str(inspect.signature(e.run))) for e in C.CONVERT_ENGINES]))
    check("⭐⭐ 有 `run_many` 的话前两个参数必须正好是 **(jobs, fmt)**（批量口同理不能写歪）",
          all(e.run_many is None or _first3_ok(e.run_many, ("jobs", "fmt"))
              for e in C.CONVERT_ENGINES),
          str([(e.key, str(inspect.signature(e.run_many)) if e.run_many else None)
               for e in C.CONVERT_ENGINES]))
    check("⭐ CAD 那条**有**批量口、BarTender 那条**没有**"
          "（一个 CAD 会话串着打完整批；BarTender 一份 .btw 要出 jpg+pdf 两个，走单份）",
          C.engine_by_key("cad").run_many is not None
          and C.engine_by_key("bartender").run_many is None)
    _magic = Path(tempfile.mkdtemp(prefix="pdfmagic_"))
    (_magic / "真.pdf").write_bytes(b"%PDF-1.7\n%\xe2\xe3\xcf\xd3\n")
    (_magic / "假的.pdf").write_bytes(b"\x1b%-12345X@PJL SET ... printer instruction stream")
    check("⭐⭐ 「产物是不是**真** PDF」这道核对认得出来 —— "
          "CAD / BarTender 都不是格式转换而是**打印**，设备选错会写出打不开的东西，"
          "只判「文件在不在」会把假成功放过去",
          C.looks_like_pdf(_magic / "真.pdf")
          and not C.looks_like_pdf(_magic / "假的.pdf")
          and not C.looks_like_pdf(_magic / "根本没有这个.pdf"))
    check("⭐ 非 Windows 机器上找不到这两个程序 → 探测直接返回 False，不去瞎试、不抛异常",
          (os.name == "nt")
          or (C.find_cad() is None and C.find_bartender() is None
              and not C._probe_cad() and not C._probe_bartender()),
          f"{C.find_cad()} / {C.find_bartender()}")
    check("⭐ 「重新检测引擎」会清掉 CAD 的探测缓存（刚装上的 CAD 才看得见）",
          callable(C.reset_cad_cache) and C.find_cad() is None
          or os.name == "nt")
    _cap = io.StringIO()
    with contextlib.redirect_stdout(_cap):
        _rc = C._main_show_cad(None)          # 这个口子不读 argv，纯只读体检
    _out = _cap.getvalue()
    check("⭐⭐ `--show-cad` 这个体检口**不碰任何文件**就能说清情况"
          "（列品牌表 + 说找没找到；本机是 Mac → 明说「非 Windows 上没有属正常」）",
          _rc == 0 and "品牌表" in _out and C.CAD_APPS[0].name in _out
          and (os.name == "nt" or "非 Windows" in _out or "正常的" in _out),
          f"rc={_rc} out={_out[-160:]!r}")

    # ⭐ BarTender 反向哨兵：出图必须走 COM（Formats.Open + Export…），
    #    不许再退回命令行打印（bartend.exe /F=…/P/PRNFILE=…）。这是 2026-10-10
    #    用户点名「跟 CAD 一样优先调用原生接口」的口径：导出不是打印，不碰模板打印机。
    _bt_src = inspect.getsource(C._run_bartender)
    check("⭐⭐ BarTender 出图走 COM（`find_bartender_com` + `Formats.Open`），"
          "**不再用命令行打印**（`bartend.exe /F=` 那套已废弃）",
          "find_bartender_com" in _bt_src and "Formats.Open" in _bt_src
          and "bartend.exe /F=" not in _bt_src and "PRNFILE" not in _bt_src)
    check("⭐⭐ BarTender 一份 .btw 要出**两个产物**（JPG + PDF），"
          "两个导出方法都在（ExportPrintPreviewToImage + ExportPrintPreviewToFile）",
          "_bartender_export_jpg" in _bt_src and "_bartender_export_pdf" in _bt_src
          and "ExportPrintPreviewToImage" in _bt_src and "ExportPrintPreviewToFile" in _bt_src)
    _bt_probe = inspect.getsource(C._probe_bartender)
    check("⭐⭐ BarTender 的「装没装」看 **COM 名字**（find_bartender_com），不是找 exe",
          "find_bartender_com" in _bt_probe)

    # ══════════════════════════════════════════════════════════════════════
    #  [13b4] ⑫ 闸门：只有「一族都轮不到」才整批拦下
    #     夹了一份 .dwg 不该把用户那几十份 Word 也拖下水（用户常态是混合来源）。
    # ══════════════════════════════════════════════════════════════════════
    print("\n[13b4] ⑫ 闸门：局部没有引擎 ≠ 整批不能转")
    _OFF_ONLY = C.ConvEngine("only_office", "只有 Office（假）", "", lambda: True,
                             lambda s, d, f: "不该被调用")
    _keep_eng = C.CONVERT_ENGINES
    _gd = Path(tempfile.mkdtemp(prefix="gatefam_"))
    (_gd / "规范.docx").write_bytes(b"x")
    (_gd / "图纸.dwg").write_bytes(b"x")
    try:
        C.CONVERT_ENGINES = (_OFF_ONLY,)          # office 有、cad 一族一条都没有
        _p1 = C.plan_convert(C.Options(sources=(str(_gd),)))
        check("⭐⭐ 只有 office 有引擎 → 队列**不阻断**"
              "（夹了一份 .dwg 不该让那批 Word 也转不了）", not _p1.blocked, _p1.block_reason)
        _dwg = [it for it in _p1.items if it.src.suffix == ".dwg"]
        check("⭐⭐ 没有引擎的那几份在**队列里就写清了原因**（不是等点完才报）",
              len(_dwg) == 1 and C.cad_brand_words() in _dwg[0].note,
              _dwg[0].note if _dwg else "（没进队列）")
        check("⭐ 有引擎的那几份 note 是空的（不打扰）",
              all(not it.note for it in _p1.items if it.src.suffix == ".docx"))
        C.CONVERT_ENGINES = ()
        _p2 = C.plan_convert(C.Options(sources=(str(_gd),)))
        check("⭐⭐ 一族都轮不到 → 整批拦下，并**分别**说清哪一族缺什么程序",
              _p2.blocked and C.cad_brand_words() in _p2.block_reason
              and "Office" in _p2.block_reason, _p2.block_reason)
    finally:
        C.CONVERT_ENGINES = _keep_eng

    print("\n[13a2] ⑪c 失败归类 / 统一报错 —— Mac 与 Win 共用一套口径")
    check("归类的候选 code 都在标题表里有话可说（不会冒出个没定义的 code）",
          all(code in C.OFFICE_ERR_TITLE for code, _ in C.OFFICE_ERR_RULES),
          str([code for code, _ in C.OFFICE_ERR_RULES]))
    check("⭐ 只有**一张**规则表（不是 Mac 一串 if、Win 又串一遍）",
          isinstance(C.OFFICE_ERR_RULES, tuple) and len(C.OFFICE_ERR_RULES) >= 3)
    for _raw, _want in (("-10004", "no_permission"), ("权限违例", "no_permission"),
                        ("-2741 syntax error", "syntax"),
                        ("Class not registered 0x80040154", "no_app"),
                        ("文件正在使用中", "busy"),
                        ("0x80070005 拒绝访问", "no_permission")):
        check(f"归类 {_raw!r} → {_want}", C.classify_office_error(_raw) == _want,
              C.classify_office_error(_raw))
    check("归不出来的不抛异常，落到 other",
          C.classify_office_error("") == "other"
          and C.classify_office_error("莫名其妙的一段话") == "other")
    _mac_note = C.office_error_note("mac", wf, "execution error: 权限违例。 (-10004)")
    _win_note = C.office_error_note("win", wf, "com_error: 拒绝访问 0x80070005")
    check("⭐ 同一类失败，两个平台都**说得出下一步干什么**（含各自那条补救说明）",
          "系统设置" in _mac_note and "对话框" in _win_note,
          f"{_mac_note[:40]} / {_win_note[:40]}")
    check("⭐ 两平台同 code 但**说明不同**（同一张表、不同条目）", _mac_note != _win_note)
    check("原始报错一个字不丢地带上（用户能原样转给我们）",
          "-10004" in _mac_note and "0x80070005" in _win_note)
    check("抬头写清了是哪个 app / 哪一类", "Microsoft Word" in _mac_note
          and "转换失败" in _mac_note, _mac_note.splitlines()[0])
    check("❓ 界面能拿到的那个 code 就是这里归出来的（ConvertResult.err_code）",
          "err_code" in C.ConvertResult.__dataclass_fields__)

    print("\n[13a4] 整批脚本的返回解析（一行一份 → 每份的失败原因）")
    check("OK 行 → 这份没有失败原因",
          C._parse_mac_batch_output("1|OK\n2|OK\n", 2) == ["", ""])
    check("ERR 行 → 号挂在尾巴上（⑪c 的归类靠它认）",
          C._parse_mac_batch_output("1|ERR|-10004|权限违例。", 1)
          == ["权限违例。 (-10004)"])
    check("⭐ 解析出来的话照样能被归好类（界面才拿得到补救入口）",
          C.classify_office_error(C._parse_mac_batch_output(
              "1|ERR|-10004|not permitted", 1)[0]) == "no_permission")
    check("脚本没给出结果的那一份 = None（交给调用方补一句人话）",
          C._parse_mac_batch_output("2|OK\n", 2) == [None, ""])
    check("瞎编的行 / 越界的序号一律忽略（不抛异常、不串行）",
          C._parse_mac_batch_output("hello\n9|OK\nx|OK\n1|OK\n", 1) == [""])
    check("报错话里带 | 也不串行（只切前 3 个，尾巴整段留给说明）",
          C._parse_mac_batch_output("1|ERR|-1|a|b|c", 1) == ["a|b|c (-1)"])
    check("结果乱序回来也按序号归位（脚本不保证顺序，别拿行序当身份）",
          C._parse_mac_batch_output("2|OK\n1|ERR|-2753|没打开", 2)
          == ["没打开 (-2753)", ""])

    # _engine_run_group 是「批量 / 单份」两条路的唯一入口 ——
    # 先备一个记录型的批量假引擎（真正跑整条流程的用例见 [13d3]）。
    _grp_jobs = [(Path("/x/a.docx"), Path("/x/a.pdf")),
                 (Path("/x/b.docx"), Path("/x/b.pdf"))]
    _grp_calls: list[int] = []

    def _grp_probe_single(_s, _d, _f):        # 单份口：有批量口时不该被走到
        raise AssertionError("有批量口还走单份口 = 分组没生效")

    def _grp_probe_many(jobs, _f):
        _grp_calls.append(len(jobs))
        return ["" for _ in jobs]

    GROUPY = C.ConvEngine("groupy", "整批（假）", "", lambda: True,
                          _grp_probe_single, _grp_probe_many)

    def _grp_probe_short(jobs, _f):
        return []                             # 故意一个结果都不给

    SHORTY = GROUPY._replace(key="shorty", run_many=_grp_probe_short)
    check("⭐ 有批量口：一批只调**一次**（不再一份一次拉应用）",
          C._engine_run_group(GROUPY, _grp_jobs, wf) == ["", ""]
          and _grp_calls == [2], str(_grp_calls))
    check("⭐ 批量口少给结果 → 按 jobs 补齐、每份都算没成（不静默当成成功）",
          len(C._engine_run_group(SHORTY, _grp_jobs, wf)) == 2
          and all(C._engine_run_group(SHORTY, _grp_jobs, wf)), "少给结果时没兜住")
    check("空的一批不调引擎（省掉一次白拉应用）",
          C._engine_run_group(GROUPY, [], wf) == [] and _grp_calls == [2],
          str(_grp_calls))

    # 假引擎（回归专用）：正常写一份真 PDF，名字带 BAD 的故意失败
    def _fake_run(src: Path, dst: Path, fmt) -> str:
        if FAKE_FAIL_MARK in src.name:
            return "假引擎故意失败（回归用）"
        build_pdf(dst, [(595.0, 842.0, f"CONV-{src.stem}"),
                        (595.0, 842.0, f"CONV-{src.stem}-P2")])
        return ""

    def _boom_probe() -> bool:
        raise RuntimeError("炸了")

    BOOM = C.ConvEngine("boom", "会炸的引擎", "", _boom_probe, lambda s, d, f: "x")
    FAKE = C.ConvEngine("fake_regress", "假引擎（回归用）", "", lambda: True, _fake_run)
    ORIG_ENGINES = C.CONVERT_ENGINES
    C.CONVERT_ENGINES = (BOOM, FAKE)
    check("⭐ 某个引擎探测时自己炸了，不拖垮整张表（跳过它继续往下探）",
          C.pick_engine() is FAKE)
    C.CONVERT_ENGINES = (FAKE,)
    check("表里只剩一个可用引擎时能选出来", C.pick_engine() is FAKE)
    C.CONVERT_ENGINES = ()
    check("一个引擎都没有 → 返回 None（不抛异常）", C.pick_engine() is None)
    C.CONVERT_ENGINES = (FAKE,)
    # ⭐ 没有批量口的引擎（假引擎 / LibreOffice）必须照样能整批走完 ——
    #    退回按单份循环，两条路都从 _engine_run_group 这同一个入口进。
    _loop_jobs = [(root / "loop1.docx", root / "loop1.pdf"),
                  (root / "loop2.docx", root / "loop2.pdf")]
    _loop_errs = C._engine_run_group(FAKE, _loop_jobs, wf)
    check("⭐ 没有批量口的引擎：退回按单份循环，结果照样一份对一份",
          _loop_errs == ["", ""]
          and all(dst.is_file() for _src, dst in _loop_jobs),
          str(_loop_errs))

    # ── ⑫ 转换闸门：三道，各自给一句人话 ────────────────────────────────
    p_no_src = C.plan_convert(C.Options())
    check("⑫ 没来源 → 阻断", p_no_src.blocked)
    check("⑫ 文案说清「还没有添加任何原件」",
          "还没有添加任何原件" in p_no_src.block_reason, p_no_src.block_reason)

    empty_d = root / "转换空夹"
    empty_d.mkdir(parents=True, exist_ok=True)
    p_empty = C.plan_convert(C.Options(sources=(str(empty_d),)))
    check("⑫ 来源在、但没有能转的文件 → 阻断", p_empty.blocked)
    check("⑫ 文案把认哪些扩展名列了出来",
          ".xlsx" in p_empty.block_reason and ".docx" in p_empty.block_reason,
          p_empty.block_reason)

    conv_d = root / "转换语料"
    conv_d.mkdir(parents=True, exist_ok=True)
    for n13 in ("WI-1032.xlsx", "WI-2001.docx", "WI-9.docm", "notes.txt", "WI-1032.PDF"):
        (conv_d / n13).write_bytes(b"fake office bytes")
    sub13 = conv_d / "子目录"
    sub13.mkdir(exist_ok=True)
    (sub13 / "WI-8888.xlsx").write_bytes(b"x")
    C.CONVERT_ENGINES = ()
    p_no_eng = C.plan_convert(C.Options(sources=(str(conv_d),)))
    check("⑫ 有来源有文件、但一个引擎都没有 → 阻断", p_no_eng.blocked)
    check("⑫ 阻断原因就是「没引擎」那份说明",
          p_no_eng.block_reason == C.engine_no_engine_msg(), p_no_eng.block_reason)
    C.CONVERT_ENGINES = (FAKE,)

    print("\n[13b] 收集待转文件（与「收集原件」同一套来源口径）")
    files13, unknown13 = C.collect_convertibles([str(conv_d)])
    check("文件夹来源只收它**这一层**（子目录里的不算）",
          [p.name for p in files13] == ["WI-9.docm", "WI-1032.xlsx", "WI-2001.docx"],
          str([p.name for p in files13]))
    check("这一层里的 .txt / .PDF 不收（PDF 不用转）",
          not any(p.name in ("notes.txt", "WI-1032.PDF") for p in files13),
          str([p.name for p in files13]))
    check("文件夹来源不会把不收的记成「认不出」（它们压根不是候选）",
          not unknown13, str([p.name for p in unknown13]))
    check("队列按文件名排（跟来源的先后无关）",
          [p.name for p in files13]
          == [p.name for p in C.sort_originals(files13, "name")])
    f13b, u13b = C.collect_convertibles([str(conv_d / "notes.txt"), str(conv_d / "WI-9.docm")])
    check("单文件来源里认不出的进 unknown（只统计、不进队）",
          [p.name for p in u13b] == ["notes.txt"] and [p.name for p in f13b] == ["WI-9.docm"],
          f"files={[p.name for p in f13b]} unknown={[p.name for p in u13b]}")
    f13c, _ = C.collect_convertibles([str(conv_d / "WI-9.docm"), str(conv_d / "WI-9.docm")])
    check("同一份文件被指到两次只收一次", len(f13c) == 1, str(len(f13c)))
    files13s, _ = C.collect_convertibles([str(sub13)])
    check("子目录**作为来源**加进来时才收（显式、可控，不会误收）",
          [p.name for p in files13s] == ["WI-8888.xlsx"], str([p.name for p in files13s]))

    print("\n[13c] 转换产物名 / 落点 / 自家产物识别")
    check("产物名 = 原件名 + 后缀 + .pdf",
          C.out_pdf_name_for(Path("/x/WI-1032.xlsx"), "_已转") == "WI-1032_已转.pdf")
    check("⭐ 后缀为空也安全（.xlsx → .pdf，扩展名就不同，名字必然不同）",
          C.out_pdf_name_for(Path("/x/WI-1032.xlsx"), "") == "WI-1032.pdf")
    check("后缀里的非法字符照样被清掉（与替换那边同一口径）",
          C.out_pdf_name_for(Path("/x/a.xlsx"), "_a/b*c") == "a_abc.pdf",
          C.out_pdf_name_for(Path("/x/a.xlsx"), "_a/b*c"))
    check("⭐ 转换不套用「后缀必填」（suffix_error 只管替换那条线）",
          C.clean_output_suffix("") == "")

    dd13 = root / "转换落点"
    dd13.mkdir(parents=True, exist_ok=True)
    src13f = dd13 / "WI-1032.xlsx"
    src13f.write_bytes(b"x")
    o13 = C.Options(suffix="_已转")
    t13 = C.resolve_convert_target(src13f, o13, set())
    check("产物写在**原件旁边**（没有输出文件夹这一说）",
          t13 == dd13 / "WI-1032_已转.pdf", str(t13))
    t13.write_bytes(b"%PDF-1.4\n")
    check("⭐ 认得出是自家产物（同目录里有同名的 .xlsx）",
          C.is_convert_product(t13, "_已转"))
    check("⭐ 认不出就别乱认（旁边没有同名的 Word / Excel）",
          not C.is_convert_product(dd13 / "别人放的_已转.pdf", "_已转"))
    check("后缀为空时不去猜 _2 尾巴（宁可不认，也不误伤别人的文件）",
          not C.is_convert_product(dd13 / "WI-1032_2.pdf", ""))
    check("原件自己不会被认成产物（它不是 .pdf）",
          not C.is_convert_product(src13f, "_已转"))
    check("AUTO：自家的产物原地更新（重跑＝刷新，不攒 _2）",
          C.resolve_convert_target(src13f, C.Options(suffix="_已转",
                                                      overwrite=C.OVERWRITE_AUTO), set()) == t13)
    check("YES：一律覆盖",
          C.resolve_convert_target(src13f, C.Options(suffix="_已转",
                                                      overwrite=C.OVERWRITE_YES), set()) == t13)
    check("NO：一律加编号让开（连自家上一轮的也不碰）",
          C.resolve_convert_target(src13f, C.Options(suffix="_已转",
                                                      overwrite=C.OVERWRITE_NO), set()
                                   ).name == "WI-1032_已转_2.pdf")

    print("\n[13d] plan_convert / run_convert（假引擎跑全程）")
    o13p = C.Options(sources=(str(conv_d),), suffix="_已转")
    p13 = C.plan_convert(o13p)
    check("没被阻断", not p13.blocked, p13.block_reason)
    check("引擎信息写进了队列（界面拿它显示「引擎：…」）",
          p13.engine_key == "fake_regress" and p13.engine_label, f"{p13.engine_key}/{p13.engine_label}")
    check("队列 3 份，序号 1 起连续",
          [it.index for it in p13.items] == [1, 2, 3], str([it.index for it in p13.items]))
    check("每份都标好了格式标签（界面「类型」列用）",
          [it.fmt_label for it in p13.items] == ["Word", "Excel", "Word"],
          str([it.fmt_label for it in p13.items]))
    check("列队阶段**不读文件内容**（只是个 Excel，页数无从谈起）",
          all(it.status == C.ConvStatus.PENDING.value for it in p13.items))
    check("⑩ 转换规则那句话有明确口径", bool(C.CONVERT_RULE_NOTE), C.CONVERT_RULE_NOTE)

    before13 = {p.stat().st_mtime_ns: p.name for p in conv_d.iterdir() if p.is_file()}
    res13, notes13 = C.run_convert(p13, o13p)
    check("跑完没有额外提示", not notes13, str(notes13))
    check("3 份全成功", all(r.ok for r in res13) and len(res13) == 3,
          str([(r.src.name, r.note) for r in res13]))
    check("⭐ 产物页数读得出来（页数有值＝PDF 真的生成了）",
          all(r.pages == 2 for r in res13), str([r.pages for r in res13]))
    check("产物就落在各自原件旁边",
          all(r.out is not None and r.out.parent == conv_d for r in res13),
          str([str(r.out) for r in res13]))
    check("产物名 = 原件名 + 后缀 + .pdf",
          sorted(r.out.name for r in res13)
          == ["WI-1032_已转.pdf", "WI-2001_已转.pdf", "WI-9_已转.pdf"],
          str(sorted(r.out.name for r in res13)))
    check("⭐ 原件一个字节都没动（连 mtime 都没变）",
          {p.stat().st_mtime_ns: p.name for p in conv_d.iterdir()
           if p.is_file() and not p.name.endswith("_已转.pdf")} == before13,
          "原件被改过")
    check("重跑一遍：自家产物认得出 → 原地更新，不攒 _2",
          len(sorted(conv_d.glob("*_已转.pdf"))) == 3,
          str(sorted(p.name for p in conv_d.glob("*_已转*.pdf"))))
    res13b, _ = C.run_convert(C.plan_convert(o13p), o13p)
    check("重跑之后产物仍是 3 份（原地更新）",
          len(sorted(conv_d.glob("*_已转*.pdf"))) == 3,
          str(sorted(p.name for p in conv_d.glob("*_已转*.pdf"))))
    check("重跑的结果同样都是成功", all(r.ok for r in res13b), str([r.note for r in res13b]))

    # 覆盖清单只回答「选覆盖会动到谁」，与三态无关
    will13 = C.convert_overwrite_list(p13, o13p)
    check("覆盖清单点名了全部 3 份（它们本来就撞上了上一轮产物）",
          len(will13) == 3, str([(t.name, s.name) for t, s in will13]))

    # dry_run 不写盘
    o13dry = C.Options(sources=(str(conv_d),), suffix="_另一版", dry_run=True)
    p13dry = C.plan_convert(o13dry)
    res13dry, _ = C.run_convert(p13dry, o13dry)
    check("dry_run：报成功但不真写文件",
          all(r.ok and r.status == C.ConvStatus.OK.value for r in res13dry)
          and not list(conv_d.glob("*_另一版.pdf")),
          str([r.note for r in res13dry]))

    # 失败 / 跳过两条分支
    bad_src = conv_d / f"{FAKE_FAIL_MARK}-1.xlsx"
    bad_src.write_bytes(b"boom")
    o13bad = C.Options(sources=(str(bad_src),), suffix="_已转")
    p13bad = C.plan_convert(o13bad)
    res13bad, _ = C.run_convert(p13bad, o13bad)
    check("失败分支：状态是「转换失败」",
          res13bad[0].status == C.ConvStatus.FAILED.value, str(res13bad[0].status))
    check("失败分支：把引擎给的原因原样带出来（不吞）",
          res13bad[0].note == "假引擎故意失败（回归用）", res13bad[0].note)
    check("⭐ 失败分支：顺手归好了类（界面靠 err_code 决定给不给「打开设置」）",
          res13bad[0].err_code == "other", res13bad[0].err_code)
    check("成功的行不带 err_code（别让界面误弹补救入口）",
          all(r.err_code == "" for r in res13), str([r.err_code for r in res13]))
    check("失败分支：没有留下半成品",
          not list(conv_d.glob(f"{FAKE_FAIL_MARK}-1*.pdf")),
          str(sorted(p.name for p in conv_d.glob("BAD*"))))
    p13sk = C.plan_convert(o13p)
    p13sk.items[0].status = C.ConvStatus.SKIPPED.value
    res13sk, _ = C.run_convert(p13sk, o13p)
    check("手动忽略的份：原样带出「已跳过」，且没转它",
          res13sk[0].status == C.ConvStatus.SKIPPED.value
          and res13sk[0].out is None, str(res13sk[0]))
    C.CONVERT_ENGINES = (FAKE,)
    p13blk = C.plan_convert(C.Options())
    res13blk, notes13blk = C.run_convert(p13blk, C.Options())
    check("被阻断的队列：一份都不转，并把原因当提示带出来",
          not res13blk and notes13blk and notes13blk[0] == p13blk.block_reason,
          str(notes13blk))

    # ── ⑩c 分页规则走完整条链路：真 .xlsx（3 个可见表）→ 假引擎只吐 2 页/3 页 ──
    print("\n[13d1] ⑩c 分页规则在 run_convert 里真生效（页数对不上要说一句）")
    _pr13 = Path(tempfile.mkdtemp(prefix="pagerule13_"))
    _xlsx13 = _pr13 / "三表.xlsx"
    write_min_xlsx(_xlsx13, [("封面", "visible"), ("工序", "visible"), ("附表", "visible")])

    def _run_n_pages(n: int):
        def _run(src: Path, dst: Path, fmt) -> str:
            build_pdf(dst, [(595.0, 842.0, f"第{i}页") for i in range(1, n + 1)])
            return ""
        return _run

    _fake2 = C.ConvEngine("fake2", "假引擎2", "", lambda: True, _run_n_pages(2))
    _o13pr = C.Options(sources=(str(_xlsx13),), suffix="_已转")
    _res_mis, _ = C.run_convert(C.plan_convert(_o13pr), _o13pr, engine=_fake2)
    check("⭐⭐ 3 个工作表却转出 2 页 → 「说明」列点出对不上",
          _res_mis[0].ok and "3 个工作表" in _res_mis[0].note
          and "2 页" in _res_mis[0].note, _res_mis[0].note)

    _fake3 = C.ConvEngine("fake3", "假引擎3", "", lambda: True, _run_n_pages(3))
    _res_ok, _ = C.run_convert(C.plan_convert(_o13pr), _o13pr, engine=_fake3)
    check("⭐⭐ 3 个工作表、正好 3 页 → 说明列**空空如也**（不打扰）",
          _res_ok[0].ok and _res_ok[0].pages == 3 and _res_ok[0].note == "",
          f"pages={_res_ok[0].pages} note={_res_ok[0].note!r}")

    # ── 引擎兜底梯队（2026-10-09）：首选失手不该让整批白跑 ────────────────
    #    为什么值得单开一节：Mac 上 Office 自动化天生脆（授权 / 沙箱 / 弹窗），
    #    Windows 上也会卡宏提示 —— 机器上还有 LibreOffice 就应该让它接着干完。
    print("\n[13d2] 引擎兜底梯队（首选失手 → 自动换下一个）")

    def _ok_run(src: Path, dst: Path, fmt) -> str:
        build_pdf(dst, [(595.0, 842.0, f"FALLBACK-{src.stem}")])
        return ""

    def _bad_run(src: Path, dst: Path, fmt) -> str:
        return "首选引擎故意失败（回归用）"

    PRIM = C.ConvEngine("prim_fail", "首选（假）", "", lambda: True, _bad_run)
    BACK = C.ConvEngine("backup_ok", "兜底（假）", "", lambda: True, _ok_run)
    C.CONVERT_ENGINES = (PRIM, BACK)
    check("梯队 = 首选 + 表里其他可用的引擎（按表序）",
          [e.key for e in C.engine_chain(PRIM)] == ["prim_fail", "backup_ok"],
          str([e.key for e in C.engine_chain(PRIM)]))
    check("显式关掉兜底 → 只剩首选自己（回归注入钦定引擎时正是这条）",
          [e.key for e in C.engine_chain(PRIM, allow_fallback=False)] == ["prim_fail"])

    fb_d = root / "兜底语料"
    fb_d.mkdir(parents=True, exist_ok=True)
    (fb_d / "F-1.docx").write_bytes(b"fake office bytes")
    o_fb = C.Options(sources=(str(fb_d),), suffix="_已转")
    res_fb, _ = C.run_convert(C.plan_convert(o_fb), o_fb)   # engine=None → 自动 → 走梯队
    check("⭐ 首选失手，兜底把活干完了（整批不至于白跑）",
          res_fb[0].ok and res_fb[0].engine_key == "backup_ok",
          f"ok={res_fb[0].ok} engine={res_fb[0].engine_key}")
    check("⭐ 实际是哪个引擎干的活按份记（不拿计划里那个当答案）",
          res_fb[0].engine_key != C.plan_convert(o_fb).engine_key,
          f"{res_fb[0].engine_key} vs {C.plan_convert(o_fb).engine_key}")
    check("⭐ 「换过引擎」这件事摆在说明列里，用户看得见",
          "自动改用" in res_fb[0].note, res_fb[0].note)
    check("产物是兜底引擎真写出来的（页数读得出来）", res_fb[0].pages == 1,
          str(res_fb[0].pages))

    o_fb2 = C.Options(sources=(str(fb_d),), suffix="_钦定")
    res_inj, _ = C.run_convert(C.plan_convert(o_fb2), o_fb2, engine=PRIM)
    check("钦定引擎（回归注入）失败就是失败，不会偷偷换成别人",
          not res_inj[0].ok and res_inj[0].note == "首选引擎故意失败（回归用）",
          f"ok={res_inj[0].ok} note={res_inj[0].note}")

    C.CONVERT_ENGINES = (PRIM, C.ConvEngine("backup_fail", "兜底也失败（假）", "",
                                            lambda: True, _bad_run))
    o_fb3 = C.Options(sources=(str(fb_d),), suffix="_全败")
    res_all, _ = C.run_convert(C.plan_convert(o_fb3), o_fb3)
    check("全都失手 → 报**首选**那句原因（用户最该看的是第一个）",
          not res_all[0].ok and res_all[0].note == "首选引擎故意失败（回归用）",
          res_all[0].note)
    C.CONVERT_ENGINES = (FAKE,)           # 交还给 [13e] 用

    # ── ⭐⭐ 2026-10-09（第二轮）：整批分组（一个格式 = 一次应用会话）──────────
    #    这是「每一份都要点一次授权」和「3 份只成功 1 份」那个毛病的**结构断言**：
    #    一个格式的所有文件交给同一个应用会话，换格式才换应用。
    #    做法：装一个会记账的批量假引擎，数「被调了几次、每次带几份」。
    print("\n[13d3] 整批分组（同格式只调一次引擎）")
    batch_d = root / "整批语料"
    batch_d.mkdir(parents=True, exist_ok=True)
    for n_b in ("B-1.docx", "B-2.xlsx", "B-3.docx", "B-4.xlsx", "B-5.docx"):
        (batch_d / n_b).write_bytes(b"fake office bytes")

    calls: list[tuple[str, int]] = []

    def _b_single(src, dst, fmt):
        calls.append((f"single:{fmt.key}", 1))
        return "这个用例不该走单份口"

    def _b_many(jobs, fmt):
        calls.append((fmt.key, len(jobs)))
        for src, dst in jobs:
            build_pdf(dst, [(595.0, 842.0, f"BATCH-{src.stem}")])
        return ["" for _ in jobs]

    BATCHY = C.ConvEngine("batchy", "整批（假）", "", lambda: True, _b_single, _b_many)
    C.CONVERT_ENGINES = (BATCHY,)
    o_b = C.Options(sources=(str(batch_d),), suffix="_已转")
    res_b, _ = C.run_convert(C.plan_convert(o_b), o_b)
    check("5 份全转成功", len(res_b) == 5 and all(r.ok for r in res_b),
          str([(r.src.name, r.note) for r in res_b]))
    check("⭐⭐ 引擎只被调 **2 次**（Word 一批 + Excel 一批），不是 5 次",
          [c[0] for c in calls] == ["word", "excel"], str(calls))
    check("⭐⭐ 每批带的是**同格式的全部文件**（3 份 Word 一起 / 2 份 Excel 一起）",
          sorted(c[1] for c in calls) == [2, 3], str(calls))
    check("⭐ 一次都没走单份口（分组真的生效，没退回一份一次）",
          not any(c[0].startswith("single") for c in calls), str(calls))
    check("⭐ 结果顺序仍与队列一一对应（界面一行对一行）",
          [r.src.name for r in res_b]
          == [it.src.name for it in C.plan_convert(o_b).items],
          str([r.src.name for r in res_b]))
    check("⭐ 页数照样按份读回来", all(r.pages == 1 for r in res_b),
          str([r.pages for r in res_b]))
    check("⭐ 批量转出来的产物照样写在各自原件旁边",
          all(r.out.parent == batch_d for r in res_b),
          str([str(r.out) for r in res_b]))

    def _b_partial(jobs, fmt):
        out = []
        for src, dst in jobs:
            if "B-3" in src.name:
                out.append("这一份故意失败（回归用）")
                continue
            build_pdf(dst, [(595.0, 842.0, f"PART-{src.stem}")])
            out.append("")
        return out

    PARTIAL = BATCHY._replace(key="partial", run_many=_b_partial)
    C.CONVERT_ENGINES = (PARTIAL,)
    o_b2 = C.Options(sources=(str(batch_d),), suffix="_半失")
    res_b2, _ = C.run_convert(C.plan_convert(o_b2), o_b2)
    ok_b2 = [r for r in res_b2 if r.ok]
    bad_b2 = [r for r in res_b2 if not r.ok]
    check("⭐⭐ 批量里一份失手：只失败那一份，同批另外 4 份照样转好",
          len(ok_b2) == 4 and len(bad_b2) == 1,
          str([(r.src.name, r.status) for r in res_b2]))
    check("⭐ 失败那行把引擎给的原因原样带出来（不吞）",
          bool(bad_b2) and bad_b2[0].note == "这一份故意失败（回归用）",
          bad_b2[0].note if bad_b2 else "没有失败的行")
    check("⭐ 整批成功的那份不会因为同批有人失手而被冤枉",
          len(set(r.out.name for r in ok_b2)) == 4, str([r.out.name for r in ok_b2]))
    C.CONVERT_ENGINES = (FAKE,)           # 收工，交还给 [13e]

    print("\n[13d4] ⭐ 转前预处理 ⑩b：带「来自网络」标记的原件改走临时副本")
    prep_d = root / "预处理语料"
    prep_d.mkdir(parents=True, exist_ok=True)
    q_src = prep_d / "带标记.docx"
    c_src = prep_d / "干净.docx"
    for _p in (q_src, c_src):
        _p.write_bytes(b"fake office bytes" * 8)

    check("⑩b 表结构完整（key / 原因 / 检测 / 处理 都能用）",
          all(r.key and r.reason and callable(r.detect) and callable(r.stage)
              for r in C.PREP_RULES), str([r.key for r in C.PREP_RULES]))
    check("⑩b 表里有「网络标记」这一条",
          any(r.key == "quarantine" for r in C.PREP_RULES), str([r.key for r in C.PREP_RULES]))

    def _mark_q(p: Path) -> None:
        subprocess.run(["/usr/bin/xattr", "-w", C.QUARANTINE_XATTR,
                        "0086;1;Microsoft Word;", str(p)], capture_output=True)

    if sys.platform == "darwin" and Path("/usr/bin/xattr").exists():
        _mark_q(q_src)
        check("⭐ 认得出「带网络标记」的文件",
              (C.first_prep_rule(q_src) or None) is not None
              and C.first_prep_rule(q_src).key == "quarantine")
        check("⭐ 干净文件不命中任何规则（绝大多数文件走的就是这条）",
              C.first_prep_rule(c_src) is None)

        pdir = root / "预处理临时"
        pdir.mkdir(parents=True, exist_ok=True)
        feed_q, why_q = C.prep_source(q_src, pdir)
        check("⭐⭐ 带标记的：换成临时副本去转，不再直接喂原件",
              feed_q != q_src and feed_q.parent == pdir, str(feed_q))
        check("⭐ 副本**自己不带**标记（否则受保护视图照样出现）",
              not C.has_quarantine(feed_q))
        check("⭐ 副本内容与原件完全一致（转出来的东西不会变）",
              feed_q.read_bytes() == q_src.read_bytes())
        check("⭐⭐ 原件**一个字节都没动、扩展属性也没动**（标记还在原件上）",
              C.has_quarantine(q_src), "原件被动过了")
        check("副本沿用原文件名（日志/命名好认）", feed_q.name == q_src.name, feed_q.name)
        feed_c, why_c = C.prep_source(c_src, pdir)
        check("干净的：原样喂原件、不给原因",
              feed_c == c_src and why_c == "", f"{feed_c} / {why_c}")

        # ⭐ 这一层绝不许「帮倒忙」：检测炸、处理炸，都必须退回原件
        def _boom_detect(_p: Path) -> bool:
            raise RuntimeError("检测故意炸（回归用）")

        def _boom_stage(_s: Path, _t: Path) -> Path | None:
            raise RuntimeError("处理故意炸（回归用）")

        _saved_rules = C.PREP_RULES
        C.PREP_RULES = (C.PrepRule("boomd", "检测就炸", _boom_detect, lambda s, t: None),
                        _saved_rules[0])
        check("⭐ 检测函数自己抛异常 → 不算命中，后面的规则照常生效",
              (C.first_prep_rule(q_src) or None) is not None
              and C.first_prep_rule(q_src).key == "quarantine")
        check("⭐ 检测抛异常 + 后面也没命中 → 老实返回 None（不抛、不崩）",
              C.first_prep_rule(c_src) is None)
        C.PREP_RULES = (C.PrepRule("booms", "处理就炸", lambda _p: True, _boom_stage),)
        check("⭐⭐ 预处理失败 → 老实退回原件（宁可撞受保护视图，也不能凭空失败）",
              C.prep_source(q_src, pdir)[0] == q_src)
        C.PREP_RULES = _saved_rules

        # 同批两份同名原件 → 副本不许互相盖掉
        d_a, d_b = root / "同名A", root / "同名B"
        d_a.mkdir(exist_ok=True)
        d_b.mkdir(exist_ok=True)
        same1, same2 = d_a / "同名.docx", d_b / "同名.docx"
        same1.write_bytes(b"one")
        same2.write_bytes(b"two")
        _mark_q(same1)
        _mark_q(same2)
        w_map, w_tmp = C._prep_sources([(0, same1, None, None), (1, same2, None, None)])
        check("⭐ 同批两份同名原件 → 两个副本各有各的落点（不互相盖掉）",
              len(set(w_map.values())) == 2, str([str(p) for p in w_map.values()]))
        check("⭐ 两份副本内容各归各的（没串）",
              {w_map[0].read_bytes(), w_map[1].read_bytes()} == {b"one", b"two"},
              str([w_map[0].read_bytes(), w_map[1].read_bytes()]))
        check("临时目录建在系统临时区，不落用户文件夹",
              w_tmp is not None and str(w_tmp).startswith(tempfile.gettempdir()), str(w_tmp))
        shutil.rmtree(w_tmp, ignore_errors=True)

        # 整批跑一遍：喂的是副本 / 结果里仍是原件 / 产物在原处 / 跑完不留垃圾
        fed: list[Path] = []

        def _p_many(jobs, fmt):
            for src, dst in jobs:
                fed.append(Path(src))
                build_pdf(dst, [(595.0, 842.0, f"PREP-{Path(src).stem}")])
            return ["" for _ in jobs]

        PREPY = C.ConvEngine("prepy", "整批（假）", "", lambda: True,
                             lambda s, d, f: "这个用例不该走单份口", _p_many)
        C.CONVERT_ENGINES = (PREPY,)
        o_p = C.Options(sources=(str(prep_d),), suffix="_已转")
        before_p = sorted(p.name for p in prep_d.iterdir())
        res_p, notes_p = C.run_convert(C.plan_convert(o_p), o_p)
        by_name = {r.src.name: r for r in res_p}
        check("⭐⭐ 带标记的那份：喂给引擎的是副本（不是原件）",
              any(p.parent != prep_d for p in fed), str([str(p) for p in fed]))
        check("⭐ 干净的那份：照旧直接喂原件",
              any(p.parent == prep_d and p.name == "干净.docx" for p in fed),
              str([str(p) for p in fed]))
        check("⭐ 结果里的 src 仍然是**原件**（界面不许冒出临时路径）",
              all(r.src.parent == prep_d for r in res_p), str([str(r.src) for r in res_p]))
        check("⭐ 产物照旧落在原件旁边",
              all(r.out is not None and r.out.parent == prep_d for r in res_p),
              str([str(r.out) for r in res_p]))
        check("⭐ 两份都成功（走副本不影响成败）",
              all(r.ok for r in res_p), str([(r.src.name, r.note) for r in res_p]))
        check("⭐ via_copy 只在带标记的那份上是 True",
              by_name["带标记.docx"].via_copy and not by_name["干净.docx"].via_copy,
              str({k: v.via_copy for k, v in by_name.items()}))
        check("⭐ 提示语交代了「已改用临时副本、原件没动」",
              any("临时副本" in n for n in notes_p), str(notes_p))
        check("⭐⭐ 跑完临时副本被清掉（不留垃圾）",
              not list(Path(tempfile.gettempdir()).glob("pdfseq_src_*")),
              str([p.name for p in Path(tempfile.gettempdir()).glob("pdfseq_src_*")]))
        check("⭐⭐ 整批跑完，原件那份的标记原样还在（从头到尾没碰原件属性）",
              C.has_quarantine(q_src), "原件被动过了")
        check("⭐ 目录里只多了两份产物，没多别的（副本落在临时区）",
              sorted(p.name for p in prep_d.iterdir())
              == sorted(before_p + ["带标记_已转.pdf", "干净_已转.pdf"]),
              str(sorted(p.name for p in prep_d.iterdir())))
    else:
        check("非 macOS：⑩b 检测一律不命中（这套标记是 macOS 专有）",
              not C.has_quarantine(q_src))
    C.CONVERT_ENGINES = (FAKE,)           # 再次交还给 [13e]

    print("\n[13e] 命令行 --mode convert")
    # ⚠️ 另开一个干净目录：上面 [13d] 故意留了一份会失败的文件在里面，
    #    拿同一个目录跑命令行，退出码会被那份失败拉成 1（跟被测的东西无关）。
    cli_d = root / "转换命令行语料"
    cli_d.mkdir(parents=True, exist_ok=True)
    for n13 in ("C-1.xlsx", "C-2.docx"):
        (cli_d / n13).write_bytes(b"fake office bytes")
    rc = C.main(["--mode", "convert", "--src", str(cli_d), "--suffix", "_已转"])
    check("命令行 convert 模式能整条跑通（假引擎）", rc == 0, str(rc))
    check("命令行也把产物写在原件旁边",
          sorted(p.name for p in cli_d.glob("*_已转.pdf"))
          == ["C-1_已转.pdf", "C-2_已转.pdf"],
          str(sorted(p.name for p in cli_d.glob("*_已转.pdf"))))
    rc2 = C.main(["--mode", "convert"])
    check("命令行没给 --src → 退出码 2（整批不执行）", rc2 == 2, str(rc2))

    # ── 命令行自检（2026-10-09 新增）：转换失败时用来自查原因 ──────────────
    rc3 = C.main(["--check-convert", str(cli_d / "C-2.docx")])
    check("自检 --check-convert：假引擎下走通、退出码 0", rc3 == 0, str(rc3))
    (cli_d / "readme.txt").write_bytes(b"not office")
    rc4 = C.main(["--check-convert", str(cli_d / "readme.txt")])
    check("自检遇到认不出的后缀 → 退出码 2（不瞎试）", rc4 == 2, str(rc4))
    rc5 = C.main(["--check-convert"])
    check("自检既没给文件名、--src 里也没有可转的 → 退出码 2 并给出用法",
          rc5 == 2, str(rc5))
    _buf = io.StringIO()
    with contextlib.redirect_stdout(_buf):
        C.main(["--check-convert", str(cli_d / "C-2.docx")])
    check("自检会把「选中哪个引擎」打出来（排查第一步就看这个）",
          "选中引擎" in _buf.getvalue(), _buf.getvalue()[:80])
    C.CONVERT_ENGINES = ORIG_ENGINES      # 收工，把引擎表还原

    print(f"\n通过 {len(PASS)} 项，失败 {len(FAIL)} 项")
    if FAIL:
        for n in FAIL:
            print("  失败：" + n)
    shutil.rmtree(root, ignore_errors=True)
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
