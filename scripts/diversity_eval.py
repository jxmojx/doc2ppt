"""多样性评测集：覆盖真实世界常见文档形态，量化翻车点。

每份文档刻意命中一个「边界」，跑 CLI 全流程，
统计页数可控性、问题数、版式命中，定位真实翻车点。

用法: python scripts/diversity_eval.py
输出: 每份文档 pass/fail + 证据 + 汇总表
"""
import sys
import tempfile
import time
from pathlib import Path

sys.path.insert(0, r"D:\make ppt\doc2ppt")
from doc2ppt import outline, parse, template


def build_docs(tmp):
    docs = []
    def add(name, ext, text):
        p = tmp / f"{name}.{ext}"
        p.write_text(text, encoding="utf-8")
        docs.append((name, p))
    # 1. 深层次结构（多级 h1/h2/h3 混排）
    add("deep_hierarchy", "md", "".join(
        "# 项目A\n## 一、背景\n### 1.1 现状\n" + "现状描述。" * 3 + "\n### 1.2 目标\n目标描述。\n"
        for _ in range(4)))
    # 2. 中英混排（中文为主夹英文术语）
    add("mixed_cn_en", "md",
        "# AI 平台建设\n\n## Data Platform\n\n平台基于 Kubernetes 与 Spark 构建，"
        "支持 batch 与 streaming 两类 workload。\n\n## 基础设施\n\n采用 AWS 与自建机房混合部署。\n")
    # 3. 纯表格大表（窄宽列、多行）
    add("wide_table", "md",
        "# 指标表\n\n| 项 | Q1 | Q2 | Q3 | Q4 | 年累计 |\n| --- | --- | --- | --- | --- | --- |\n"
        + "".join(f"| 指标{i} | {i} | {i+1} | {i+2} | {i+3} | {i*2} |\n" for i in range(1, 9)))
    # 4. 超长单段（2000+ 字无标题）
    add("super_long_para", "md",
        "# 长文\n\n" + "这是很长的一段文字，" * 120 + "\n")
    # 5. 大量列表（>8 条，含嵌套）
    add("many_bullets", "md",
        "# 清单\n\n" + "".join(f"- 事项{i}：描述文字若干\n" for i in range(1, 13)))
    # 6. 代码块 + 引用（md 特有）
    add("code_quote", "md",
        "# 技术文档\n\n## 示例\n\n```python\nprint('hi')\n```\n\n> 这是一条引用说明。\n\n## 结论\n\n代码可运行。\n")
    # 7. 短标题无正文（纯标题目录）
    add("headings_only", "md", "# 目录\n## 一、概述\n## 二、方案\n## 三、实施\n## 四、效果\n")
    # 8. 多表格 + 图表候选（数值列可画图）
    add("chart_bait", "md",
        "# 销售分析\n\n| 月份 | 销售额 | 环比 |\n| --- | --- | --- |\n"
        + "".join(f"| {i}月 | {100+i*10} | {i*5}% |\n" for i in range(1, 8)))
    return docs


def run_doc(path, page_target=8):
    _prs0, infos = template.load_template(template.template_path())
    body = infos["content"].slots["body"]
    blocks = parse.parse_document(path)
    items = outline.build_outline(blocks, page_target=page_target, doc_name=path.name,
                                  max_chars=template.body_capacity(infos),
                                  body_box=(body.w, body.h))
    out = Path(tempfile.mkdtemp()) / f"{path.stem}.pptx"
    t0 = time.time()
    from doc2ppt import fix
    rounds, _log, issues = fix.run(items, template.template_path(), str(out))
    return items, issues, rounds, time.time() - t0, out


def main():
    with tempfile.TemporaryDirectory() as td:
        docs = build_docs(Path(td))
        ok, bad = [], []
        for name, path in docs:
            try:
                _items, issues, rounds, secs, out = run_doc(path)
            except Exception as e:  # noqa: BLE001 评测要兜住一切
                bad.append((name, "CRASH", repr(e)))
                print(f"[{name}] CRASH {e!r}")
                continue
            layouts = {}
            from pptx import Presentation
            prs = Presentation(str(out))
            for sl in prs.slides:
                layouts[sl.slide_layout.name] = layouts.get(sl.slide_layout.name, 0) + 1
            n = len(prs.slides._sldIdLst)
            verdict = "OK" if not issues else "FAIL"
            (ok if not issues else bad).append(
                (name, verdict, f"{rounds}r {secs:.1f}s {n}页 {layouts} issues={len(issues)}"))
            print(f"[{name}] {verdict} rounds={rounds} {secs:.1f}s {n}页 layouts={layouts}")
            for i in issues[:4]:
                print(f"      · p{i.page_no} {i.issue_type}/{i.severity}: {i.detail[:56]}")
    print("\n=== 汇总 ===")
    print(f"通过: {len(ok)} / 失败: {len(bad)}")
    for name, v, extra in bad:
        print(f"  [{name}] {v} {extra}")


if __name__ == "__main__":
    main()
