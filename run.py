"""doc2ppt CLI 入口：python run.py <文档路径> [-t 模板] [-o 输出] [--pages N]
                     python run.py --eval [--eval-dir 目录]   批量评测并输出硬指标

输入 docx / pdf（有文本层）/ md，输出可编辑 .pptx（套用模板版式，非图片拼接）。
断网可用：大纲走规则版（按标题层级切页），模型接口已预留。
"""
import argparse
import sys
import tempfile
import time
from pathlib import Path

from doc2ppt import fix, outline, parse, template, verify

ROOT = Path(__file__).resolve().parent
SUPPORTED = {".docx", ".pdf", ".md"}


def main(argv=None):
    ap = argparse.ArgumentParser(
        prog="run.py",
        description="文档转 PPT：docx / pdf（有文本层）/ md → 可编辑 .pptx",
    )
    ap.add_argument("doc", nargs="?", help="输入文档路径（.docx / .pdf / .md）")
    ap.add_argument("-t", "--template", default=str(ROOT / "templates" / "standard_16x9.pptx"),
                    help="模板 pptx 路径（默认 templates/standard_16x9.pptx）")
    ap.add_argument("-o", "--output", default=None, help="输出 pptx 路径（默认 <文档名>.pptx）")
    ap.add_argument("--pages", type=int, default=8,
                    help="期望总页数（含封面，默认 8；与自然页数差距过大时忽略）")
    ap.add_argument("--llm", action="store_true",
                    help="用 LLM 生成大纲（读 DOC2PPT_API_URL / DOC2PPT_API_KEY / DOC2PPT_MODEL，"
                         "失败自动回退规则版）")
    ap.add_argument("--eval", action="store_true", help="批量评测并打印硬指标（不产出单文件说明）")
    ap.add_argument("--eval-dir", default=None, help="评测用例目录（默认 samples/）")
    args = ap.parse_args(argv)

    if args.eval:
        return _eval(Path(args.template), Path(args.eval_dir) if args.eval_dir else None)

    if not args.doc:
        ap.print_help()
        return 2
    doc = Path(args.doc)
    if not doc.exists():
        print(f"错误: 文件不存在: {doc}", file=sys.stderr)
        return 2
    tpl = Path(args.template)
    if not tpl.exists():
        print(f"错误: 模板不存在: {tpl}", file=sys.stderr)
        return 2
    out = Path(args.output) if args.output else doc.with_suffix(".pptx")

    t0 = time.time()
    stats = {}
    blocks, items, rounds, log, issues = _run_once(doc, tpl, out, args.pages, preview=True,
                                                   stats=stats, use_llm=args.llm)
    elapsed = time.time() - t0

    print(f"输入: {doc}")
    print(f"输出: {out}")
    print(f"预览: {ROOT.parent / 'preview' / doc.stem / 'slide01.png'}")
    print(f"解析块数: {len(blocks)} | 大纲页数: {len(items)} | 修复轮次: {rounds} | 耗时: {elapsed:.1f}s")
    print("修复日志:")
    for line in log:
        print(f"  - {line}")
    verify.print_report(out, issues)
    if args.llm:
        u = outline.LLM_USAGE
        cost = outline.llm_cost_cny()
        msg = (f"LLM 用量: 调用 {u['calls']} 次, 缓存命中 {u['cache_hits']}, "
               f"输入 {u['prompt_tokens']} / 输出 {u['completion_tokens']} tokens")
        if cost is not None:
            msg += f"，估算成本 ¥{cost:.4f}"
        print(msg)
    if stats.get("degraded"):
        print("⚠ 校验降级：本次只跑了估算（未检查标点），请在有 PowerPoint 的环境复验后再交付",
              file=sys.stderr)
        return 1
    return 0 if not issues else 1


def _run_once(doc, tpl, out, pages, preview=False, stats=None, use_llm=False):
    """跑完整管线，返回 (blocks, items, rounds, log, issues)。"""
    try:
        blocks = parse.parse_document(doc)
    except ValueError as e:
        raise SystemExit(f"错误: {e}")
    if not blocks:
        hint = "（该 PDF 可能没有文本层，属扫描件；请先 OCR 再转换）" \
            if doc.suffix.lower() == ".pdf" else ""
        print(f"警告: 未提取到任何内容{hint}", file=sys.stderr)
    _prs, infos = template.load_template(tpl)
    capacity = template.body_capacity(infos)
    body_spec = infos["content"].slots.get("body")
    body_box = (body_spec.w, body_spec.h) if body_spec else None
    items = outline.build_outline(blocks, page_target=pages, doc_name=doc.name,
                                  max_chars=capacity, body_box=body_box, use_llm=use_llm)
    png_dir = (ROOT.parent / "preview" / doc.stem) if preview else None
    rounds, log, issues = fix.run(items, tpl, out, png_dir=png_dir, stats=stats)
    return blocks, items, rounds, log, issues


def _eval(tpl, eval_dir=None):
    """批量评测：输出 6 个硬指标（溢出率/字号合规率/一次通过率/平均轮次/耗时/页数分布）。"""
    tpl = Path(tpl)
    if not tpl.exists():
        print(f"错误: 模板不存在: {tpl}", file=sys.stderr)
        return 2

    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)
        docs = _eval_cases(eval_dir, tmp)
        rows = []
        for doc in docs:
            out = tmp / (doc.stem + ".pptx")
            t0 = time.time()
            stats = {}
            try:
                _blocks, items, rounds, _log, issues = _run_once(doc, tpl, out, 8, stats=stats)
            except SystemExit as e:
                print(f"  跳过 {doc.name}: {e}")
                continue
            elapsed = time.time() - t0
            boxes, bad_font = _deck_stats(out)
            overflows = sum(1 for i in issues if i.issue_type == "overflow")
            rows.append({
                "doc": doc.name, "pages": len(items), "rounds": rounds,
                "secs": elapsed, "boxes": boxes, "overflows": overflows,
                "bad_font": bad_font, "passed": not issues,
                "first": stats.get("first", 0),
            })

    if not rows:
        print("没有可评测的文档")
        return 1

    _print_eval(rows)
    return 0 if all(r["passed"] for r in rows) else 1


def _eval_cases(eval_dir, tmp):
    """评测用例：指定目录，或 samples/ + 三个合成边界文档。"""
    if eval_dir is not None:
        return sorted(p for p in eval_dir.iterdir() if p.suffix.lower() in SUPPORTED)
    docs = sorted(p for p in (ROOT / "samples").iterdir() if p.suffix.lower() in SUPPORTED)
    # 合成边界用例（不落地到项目目录）
    (tmp / "case_plain.md").write_text(
        "# 普通文档\n\n第一段正文内容，用于评测无章节结构时的处理。\n\n第二段正文内容。\n",
        encoding="utf-8")
    (tmp / "case_long.md").write_text(
        "# 长段落文档\n\n## 一、正文\n\n" + "这是一段很长的正文内容，用于评测超长段落切分。" * 40 + "\n",
        encoding="utf-8")
    (tmp / "case_table.md").write_text(
        "# 表格文档\n\n| A | B | C |\n| --- | --- | --- |\n"
        + "".join(f"| a{i} | b{i} | c{i} |\n" for i in range(1, 10)), encoding="utf-8")
    docs += [tmp / "case_plain.md", tmp / "case_long.md", tmp / "case_table.md"]
    return docs


def _deck_stats(path):
    """统计文本框数与字号不合规数（标题槽与图表/表格页小字注豁免）。"""
    from pptx import Presentation

    prs = Presentation(str(path))
    boxes = bad = 0
    for slide in prs.slides:
        # 与 verify 同口径：有表格/图表时，「数据来源」小注允许 12pt
        has_graphic = any(sh.has_table or sh.has_chart for sh in slide.shapes)
        for sh in slide.shapes:
            if sh.has_table or not sh.has_text_frame or not sh.text_frame.text.strip():
                continue
            try:
                is_title = sh.placeholder_format.idx == 0
            except ValueError:
                is_title = False
            sizes = [r.font.size.pt for p in sh.text_frame.paragraphs for r in p.runs
                     if r.font.size is not None]
            if is_title or not sizes:
                continue
            if has_graphic and max(sizes) <= 12.5:  # 数据来源小注
                continue
            boxes += 1
            if min(sizes) < 14:
                bad += 1
    return boxes, bad


def _print_eval(rows):
    n = len(rows)
    boxes = sum(r["boxes"] for r in rows)
    overflows = sum(r["overflows"] for r in rows)
    bad_font = sum(r["bad_font"] for r in rows)
    passed = sum(1 for r in rows if r["passed"])
    rounds = sum(r["rounds"] for r in rows) / n
    secs = sum(r["secs"] for r in rows) / n
    pages = [r["pages"] for r in rows]

    print("==== 评测明细 ====")
    print(f"{'文档':<24}{'页数':>5}{'轮次':>5}{'耗时':>8}{'文本框':>7}{'首轮':>5}{'溢出':>5}{'小字':>5}  结果")
    for r in rows:
        print(f"{r['doc']:<24}{r['pages']:>5}{r['rounds']:>5}{r['secs']:>7.1f}s"
              f"{r['boxes']:>7}{r.get('first', 0):>5}{r['overflows']:>5}{r['bad_font']:>5}  "
              f"{'通过' if r['passed'] else '有问题'}")

    first_total = sum(r.get("first", 0) for r in rows)
    print("\n==== 硬指标 ====")
    print(f"文档数            : {n}")
    print(f"首轮问题率        : {first_total / max(boxes, 1):.1%}  （{first_total} 项 / {boxes} 文本框，修复前）")
    print(f"文字溢出率        : {overflows / max(boxes, 1):.2%}  （{overflows}/{boxes} 文本框，修复后）")
    print(f"字号合规率        : {1 - bad_font / max(boxes, 1):.2%}  （{boxes - bad_font}/{boxes}）")
    print(f"一次通过率        : {passed / n:.0%}  （{passed}/{n}）")
    print(f"平均修复轮次      : {rounds:.2f}")
    print(f"单份平均耗时      : {secs:.1f}s")
    print(f"页数分布          : 最少 {min(pages)} / 中位 {sorted(pages)[len(pages) // 2]} / "
          f"最多 {max(pages)}（均值 {sum(pages) / n:.1f}）")


if __name__ == "__main__":
    sys.exit(main())
