"""OCR 精度评测：把带文本层的 PDF 页面栅格化后过 OCR，与原文对齐打分。

指标：
  - char_acc  字符级精确率（编辑距离对齐后，命中字符 / 原文总字符）
  - line_paras 恢复的段落数 vs 真实段落数（版面还原度）
  - heading_hit 真实标题是否被 OCR 判定为 heading（分类精度）

用法: python scripts/ocr_eval.py [--dpi 200] [--pages 1-3]
"""
import argparse
import sys

import pymupdf

sys.path.insert(0, r"D:\make ppt\doc2ppt")
from doc2ppt import ocr


def align_accuracy(got, want):
    """字符级对齐准确率：对两个字符串做编辑距离对齐，返回命中/总字符比。"""
    got, want = got.replace(" ", ""), want.replace(" ", "")
    n, m = len(got), len(want)
    if n == 0 or m == 0:
        return 0.0
    dp = [[0] * (m + 1) for _ in range(n + 1)]
    path = [[None] * (m + 1) for _ in range(n + 1)]
    for i in range(n + 1):
        dp[i][0] = i; path[i][0] = "del"
    for j in range(m + 1):
        dp[0][j] = j; path[0][j] = "ins"
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            c = 0 if got[i - 1] == want[j - 1] else 1
            opts = [(dp[i - 1][j - 1] + c, "match"),
                    (dp[i - 1][j] + 1, "del"),
                    (dp[i][j - 1] + 1, "ins")]
            best = min(opts, key=lambda t: t[0])
            dp[i][j], path[i][j] = best[0], best[1]
    i, j, hits = n, m, 0
    while i > 0 or j > 0:
        op = path[i][j] or "match"
        if op == "match":
            if got[i - 1] == want[j - 1]:
                hits += 1
            i -= 1; j -= 1
        elif op == "del":
            i -= 1
        else:
            j -= 1
    return hits / m if m else 0.0


def ground_truth(page):
    """从文本层取真实段落（与解析层同口径：按行合并）。"""
    paras, cur, cur_size, prev_y = [], [], None, None
    for b in page.get_text("dict")["blocks"]:
        if b.get("type") != 0:
            continue
        for line in b["lines"]:
            text = "".join(s["text"] for s in line["spans"]).strip()
            if not text:
                continue
            size = max(s["size"] for s in line["spans"])
            y = line["bbox"][1]
            if (cur and abs(size - cur_size) < 0.5 and y - prev_y < size * 1.6 + 2):
                cur.append(text)
            else:
                if cur:
                    paras.append("".join(cur))
                cur, cur_size = [text], size
            prev_y = y
    if cur:
        paras.append("".join(cur))
    return paras


def collect_paragraphs(ocr_lines):
    """把 OCR 行按当前 page_paras 逻辑聚段（与 ocr.page_paras 同口径）。"""
    if not ocr_lines:
        return []
    med_h = sorted(l["h"] for l in ocr_lines)[len(ocr_lines) // 2] or 1.0
    paras = []
    for l in ocr_lines:
        cur = paras[-1] if paras else None
        if (cur is not None and l["y0"] - cur["y1"] < 0.9 * med_h
                and min(cur["x1"], l["x1"]) - max(cur["x0"], l["x0"]) > 0):
            cur["text"] += l["text"]
            cur["y1"] = max(cur["y1"], l["y1"])
        else:
            paras.append(dict(l))
    return [p["text"] for p in paras]


def evaluate(doc_path, dpi, page_nos):
    doc = pymupdf.open(doc_path)
    rows = []
    for pno in page_nos:
        page = doc[pno - 1]
        want = ground_truth(page)
        pix = page.get_pixmap(dpi=dpi)
        result, _elapse = ocr._engine()(pix.tobytes("png"))
        ocr_lines = []
        for box, text, score in result or []:
            try:
                score = float(score)
            except (TypeError, ValueError):
                score = 0.0
            if score < 0.5 or not str(text).strip():
                continue
            xs = [p[0] for p in box]; ys = [p[1] for p in box]
            ocr_lines.append({"y0": min(ys), "y1": max(ys), "x0": min(xs), "x1": max(xs),
                              "h": max(ys) - min(ys), "text": str(text).strip()})
        ocr_lines.sort(key=lambda l: (round(l["y0"], 1), l["x0"]))
        got_paras = collect_paragraphs(ocr_lines)
        got_all = "".join(got_paras)
        want_all = "".join(want)
        acc = align_accuracy(got_all, want_all)
        rows.append({
            "page": pno, "dpi": dpi,
            "char_acc": round(acc * 100, 1),
            "paras": f"{len(got_paras)}/{len(want)}",
            "want_len": len(want_all),
        })
    doc.close()
    return rows


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dpi", type=int, default=200)
    ap.add_argument("--pages", default="1-3")
    args = ap.parse_args()
    lo, hi = (int(x) for x in args.pages.split("-"))
    total = 0.0
    for name, path in [("plan", r"D:\make ppt\doc2ppt\samples\sample_plan.pdf")]:
        for r in evaluate(path, args.dpi, range(lo, hi + 1)):
            print(f"[{name} p{r['page']} @{r['dpi']}dpi] char_acc={r['char_acc']}% "
                  f"paras={r['paras']} want_len={r['want_len']}")
            total += r["char_acc"]
    print(f"\nmean char_acc: {total / (hi - lo + 1):.1f}%")


if __name__ == "__main__":
    main()
