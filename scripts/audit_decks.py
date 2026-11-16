"""跨 decks 的确定性视觉审计：量化视觉验收代理查的每一项。

用法: python scripts/audit_decks.py
输出每页: 填充至 N% / 单段最宽 / 页码戳是否合规；异常页标红。
"""
import sys

from pptx import Presentation

sys.path.insert(0, r"D:\make ppt\doc2ppt")

DECKS = [
    ("report", r"D:\make ppt\示例输出\示例-进度报告-输出.pptx", 9),
    ("notes", r"D:\make ppt\示例输出\示例-会议纪要-输出.pptx", 10),
    ("plan", r"D:\make ppt\示例输出\示例-技术方案-输出.pptx", 10),
]


def body_shapes(slide):
    return [sh for sh in slide.shapes
            if sh.has_text_frame
            and not sh.name.startswith("doc2ppt_")
            and sh.text_frame.text.strip()]


def audit(name, path, expect):
    prs = Presentation(path)
    ok, bad = [], []
    for i, slide in enumerate(prs.slides, 1):
        stamps = [sh.text_frame.text.strip() for sh in slide.shapes
                  if sh.has_text_frame and sh.name == "doc2ppt_stamp"]
        pageline = stamps[0] if stamps else None
        pno_ok = (i == 1 and "共" in (pageline or "")) or (i > 1 and f"{i} /" in (pageline or ""))
        bodies = body_shapes(slide)
        maxpara = max((len(p.text) for b in bodies
                       for p in b.text_frame.paragraphs if p.text.strip()), default=0)
        slide_h = prs.slide_height
        biggest_h = 0.0
        for b in bodies:
            top = b.top or 0
            hgt = b.height or 0
            biggest_h = max(biggest_h, (top + hgt) / slide_h)
        issues = []
        if not pno_ok:
            issues.append("页码戳缺失/不符")
        if maxpara > 90:
            issues.append(f"文字墙嫌疑(单段{maxpara}字)")
        if biggest_h < 0.50:
            issues.append(f"内容仅到{biggest_h * 100:.0f}%页高(疑似过空)")
        if biggest_h > 0.99:
            issues.append("内容抵到页底(疑似溢出)")
        tag = "[{}]p{}: 填充至{:.0f}% 单段最宽{}字 页码{}".format(
            name, i, biggest_h * 100, maxpara, "OK" if pno_ok else pageline)
        (bad if issues else ok).append(tag + (" " + " ".join(issues) if issues else ""))
    return ok, bad


all_ok, all_bad = [], []
for name, path, expect in DECKS:
    ok, bad = audit(name, path, expect)
    all_ok += ok
    all_bad += bad
    for line in ok:
        print("   ", line)
    for line in bad:
        print("   ", "FAIL " + line)
print(f"summary: {len(all_ok)} pages ok, {len(all_bad)} pages need attention")
