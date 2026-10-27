#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""真实语料验证 —— 拿桌面上真实的 WI 文件跑一遍，不写回原处。

做法：真实原件**拷进 /tmp** → 用它们各自的首页拼成一份「扫描件」→ 批量替换 →
      逐页比对内容/页数/尺寸，并确认原件一个字节都没变。

⚠️ v3.0 起没有输出文件夹了，产物写在**每份原件旁边** —— 也就是 /tmp 那份副本
   所在的目录。桌面上的真实文件全程只读（只被 copy2 读走过）。
   所以下面「原件没被动 / 产物有几份」都必须按**拷贝前记下的名字**分开数。

⭐ v3.0 还多验了一段「转换模式」：真实目录里除了 PDF **确实混着 .xlsx**
   （共表 / 数据表），所以那一段不能拍脑袋断言「一份 Word/Excel 都没有」，
   得按实际内容量。同时反手验一个「只有 PDF 的目录必须拦住」的反例。

跑法：.venv/bin/python -B _regress/verify_real_corpus.py
"""
from __future__ import annotations

import hashlib
import shutil
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import seq_core as C  # noqa: E402

PASS: list[str] = []
FAIL: list[str] = []


def check(name: str, cond: bool, detail: str = "") -> None:
    (PASS if cond else FAIL).append(name)
    print(f"  {'OK ' if cond else 'NG '} {name}" + (f"  <{detail}>" if detail and not cond else ""))


def md5(p: Path) -> str:
    return hashlib.md5(p.read_bytes()).hexdigest()


def texts(pdf: Path) -> list[str]:
    from pypdf import PdfReader

    return [(pg.extract_text() or "").strip() for pg in PdfReader(str(pdf)).pages]


def sizes(pdf: Path) -> list[tuple[float, float]]:
    """纸面尺寸（mediabox 原样，不看 /Rotate）。"""
    from pypdf import PdfReader

    out = []
    for pg in PdfReader(str(pdf)).pages:
        b = pg.mediabox
        out.append((round(float(b.width), 1), round(float(b.height), 1)))
    return out


def disp_sizes(pdf: Path) -> list[tuple[float, float]]:
    """显示尺寸（带 /Rotate 90/270 的页按横放后的尺寸算）—— 贴合就是对齐这个。

    ⚠️ 真实 WI 的封面常常是 /Rotate 90 的横向页：纸面 595×842，显示却是 842×595。
       工具按显示尺寸贴合，所以测试也必须按显示尺寸比，否则会误报。
    """
    from pypdf import PdfReader

    return [tuple(round(v, 1) for v in C.display_size(pg))
            for pg in PdfReader(str(pdf)).pages]


CORPUS = [Path("/Users/langzi/Desktop/成功1"), Path("/Users/langzi/Desktop/成功2")]
SUF = "_已签字首页"          # ⭐ v3.0：后缀必填，产物名 = 原件名 + 它


def main() -> int:
    reals = sorted({p for d in CORPUS if d.is_dir()
                    for p in d.glob("*.pdf") if p.is_file()})
    if not reals:
        print("  跳过：桌面上的真实语料（成功1 / 成功2）不在了")
        return 0
    print(f"真实原件 {len(reals)} 份")

    tmp = Path(tempfile.gettempdir()) / "pdfseq_real"
    if tmp.exists():
        shutil.rmtree(tmp)
    src = tmp / "原件"
    src.mkdir(parents=True)
    before: dict[str, str] = {}
    for p in reals:
        dst = src / p.name
        shutil.copy2(p, dst)
        before[dst.name] = md5(dst)

    # ── 用真实首页拼一份「扫描件」（页数 = 原件份数）─────────────────────
    from pypdf import PdfReader, PdfWriter

    boot_opts = C.Options(sources=(str(src),), suffix=SUF)
    ordered = C.collect_originals(boot_opts.sources, boot_opts)
    writer = PdfWriter()
    for p in ordered:
        r = PdfReader(str(p))
        if r.is_encrypted:
            try:
                r.decrypt("")
            except Exception:  # noqa: BLE001
                pass
        writer.add_page(r.pages[0])
    scan = tmp / "扫描件_签字版.pdf"
    with open(scan, "wb") as fh:
        writer.write(fh)
    print(f"扫描件 {len(ordered)} 页 → {scan.name}\n")

    # ⭐ v3.0：没有输出文件夹了 —— 产物写在每份原件旁边（也就是 src 里）
    opts = C.Options(sources=(str(src),), scan_pdf=str(scan), suffix=SUF)
    plan = C.plan_batch(opts)
    check("真实原件全部读得出，没被阻断", not plan.blocked, plan.block_reason)
    if plan.blocked:
        return 1
    check("配对份数 = 原件份数", len(plan.items) == len(ordered), str(len(plan.items)))

    scan_texts = texts(scan)
    ori_texts = {p.name: texts(p) for p in ordered}
    ori_disp = {p.name: disp_sizes(p) for p in ordered}

    results, _ = C.run_batch(plan, opts)
    bad = [r for r in results if not r.ok]
    check("全部替换成功", not bad, str([(r.src.name, r.note) for r in bad]))
    if bad:
        return 1

    for r in results:
        t = texts(r.out)
        o = ori_texts[r.src.name]
        i = r.scan_index
        check(f"{r.src.name[:28]}：第 1 页 = 扫描件第 {i} 页",
              t and t[0] == scan_texts[i - 1], str(t[:1])[:60])
        check(f"{r.src.name[:28]}：第 2 页起与原件逐页一致",
              t[1:] == o[1:], f"{len(t)} vs {len(o)}")
        check(f"{r.src.name[:28]}：页数不变", len(t) == len(o), f"{len(t)} vs {len(o)}")
        s = sizes(r.out)
        if r.fit_mode == C.FIT_FIT:
            want = ori_disp[r.src.name][0]
            check(f"{r.src.name[:28]}：首页尺寸贴合原件（按显示尺寸，含 /Rotate）",
                  s[0] == want, f"{s[0]} vs {want}")
        else:
            check(f"{r.src.name[:28]}：比例不一致时原样保留（有说明）", bool(r.note))

    # ⚠️ 产物就躺在 src 里，所以「原件没被动」只比对**拷进来之前记下的那些名字**
    check("原件一个字节都没动",
          {n: md5(src / n) for n in before} == before)
    products = [p for p in src.glob("*.pdf") if p.name not in before]
    check("产物份数 = 原件份数（都写在原件旁边）",
          len(products) == len(ordered), str(len(products)))
    check("产物名 = 原名 + 后缀",
          all(p.name.endswith(f"{SUF}.pdf") for p in products),
          str(sorted(p.name for p in products)[:3]))
    check("产物名一个都不等于原件名",
          all(p.name not in before for p in products))

    # ── 第二次跑：产物被认出来（不会被当成原件，数量也不会翻倍）──────────
    plan2 = C.plan_batch(opts)
    check("重跑：预览仍是原件那么多份（产物被认出来、没混进原件）",
          len(plan2.items) == len(ordered), str(len(plan2.items)))
    check("重跑：如实报出跳过了几份产物",
          len(plan2.products) == len(products), str(len(plan2.products)))
    will = C.overwrite_list(plan2, opts)
    check("重跑：覆盖清单点名了全部产物（覆盖＝原地更新，不是新建）",
          {t for t, _s in will} == set(products), str(len(will)))

    # 数量对不上必须整批拒绝（真实文件上再验一次）
    w2 = PdfWriter()
    r0 = PdfReader(str(scan))
    w2.add_page(r0.pages[0])
    short = tmp / "少一页.pdf"
    with open(short, "wb") as fh:
        w2.write(fh)
    p2 = C.plan_batch(C.Options(sources=(str(src),), scan_pdf=str(short), suffix=SUF))
    check("真实文件下：少一页就整批拒绝", p2.blocked and "差" in p2.block_reason, p2.block_reason)

    # ── 后缀必填：真实语料上再兜一道（空后缀＝产物会盖掉原件）──────────────
    empty = C.Options(sources=(str(src),), scan_pdf=str(scan), suffix="")
    res_e, notes_e = C.run_batch(C.plan_batch(empty), empty)
    check("真实文件下：后缀为空 → 拒绝执行，一份都不写",
          res_e == [] and notes_e == [C.SUFFIX_REQUIRED_MSG], str(notes_e))
    check("真实文件下：被拒之后原件旁的产物份数没变",
          len([p for p in src.glob("*.pdf") if p.name not in before]) == len(products))
    check("真实文件下：原件仍然一个字节都没动",
          {n: md5(src / n) for n in before} == before)

    # ── 多来源（只列队、一个字都不写）：两个桌面目录各当一路来源 ──────────
    live = [d for d in CORPUS if d.is_dir()]
    if len(live) >= 2:
        o_multi = C.Options(sources=tuple(str(d) for d in live))
        pm = C.plan_batch(o_multi)
        check("真实语料：两路来源一起收，不阻断（只列队，不需要扫描件）",
              not pm.blocked, pm.block_reason)
        check("真实语料：收到的份数 = 两个目录里的 PDF 份数之和",
              len(pm.items) == len(reals), f"{len(pm.items)} vs {len(reals)}")
        check("真实语料：只列队 → paired=False、每行没有扫描页",
              pm.paired is False and all(it.scan_index == 0 for it in pm.items))
        names_multi = [it.src.name for it in pm.items]
        nkeys = [C.natural_key(n) for n in names_multi]
        check("真实语料：队列按文件名统一排（与来自哪个目录无关）",
              nkeys == sorted(nkeys), str(names_multi[:3]))

    # ── 转换模式（v3.0 新增）：拿真实目录实地量一遍「什么会被收、什么会被略过」──
    # ⚠️ 真实语料里**确实混着 .xlsx**（共表 / 数据表），所以这里不能想当然地断言
    #    「一份 Word/Excel 都没有」—— 那是我拍脑袋，不是事实。按实际内容来验。
    # ⚠️ 先塞一个假引擎：不这么做的话，「这台机器有没有 Office」会左右队列有没有
    #    建出来（闸门③在列队**之前**），测试就变成看环境脸色了。
    live_all = [d for d in CORPUS if d.is_dir()]
    if live_all:
        real_engines = C.CONVERT_ENGINES
        C.CONVERT_ENGINES = (C.ConvEngine("probe_only", "只用于回归", "",
                                          lambda: True, lambda s, d, f: "不该被调用"),)
        pc = C.plan_convert(C.Options(sources=tuple(str(d) for d in live_all)))
        check("真实语料：没被阻断（注入假引擎，排除「有没有装 Office」的干扰）",
              not pc.blocked, pc.block_reason)
        want_xlsx = [p.name for d in live_all for p in d.glob("*.xlsx")]
        check("⭐ 收上来的就是那几个 .xlsx（共表 / 数据表），一份不多一份不少",
              sorted(it.src.name for it in pc.items) == sorted(want_xlsx),
              f"{sorted(it.src.name for it in pc.items)} vs {sorted(want_xlsx)}")
        nk = [C.natural_key(it.src.name) for it in pc.items]
        check("⭐ 队列和替换那边同一条规矩：全部按文件名统一排（跨来源）",
              nk == sorted(nk), str([it.src.name for it in pc.items]))
        check("「所在文件夹」列能分辨出这一份是哪家的（跨目录来源）",
              len({it.src.parent for it in pc.items}) > 1,
              str([it.src.parent.name for it in pc.items]))
        check("PDF 一份都不进转换队（PDF 不需要转）",
              all(it.src.suffix.lower() in C.CONVERT_EXTS for it in pc.items),
              str([(it.src.name, it.src.suffix) for it in pc.items]))
        check("PDF 也不会被记成「认不出」（文件夹来源里它们压根不是候选）",
              not pc.unknown, str([p.name for p in pc.unknown]))
        check("格式标签认得出 Excel", all(it.fmt_label == "Excel" for it in pc.items),
              str([it.fmt_label for it in pc.items]))

        # 反过来：只有 PDF 的目录必须**拦住**（不能糊弄出一份空队列）
        p_pdf_only = C.plan_convert(C.Options(sources=(str(src),)))
        check("只有 PDF 的目录 → 阻断，并说清「没找到能转的」",
              p_pdf_only.blocked and "没找到能转的" in p_pdf_only.block_reason,
              p_pdf_only.block_reason)
        check("阻断时队列是空的（不会让用户以为「转完了」）",
              not p_pdf_only.items, f"items={len(p_pdf_only.items)}")
        C.CONVERT_ENGINES = real_engines

    shutil.rmtree(tmp, ignore_errors=True)
    print(f"\n通过 {len(PASS)} 项，失败 {len(FAIL)} 项")
    for n in FAIL:
        print("  失败：" + n)
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
