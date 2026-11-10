"""可选 OCR：扫描版 PDF 的兜底解析通道。

仅当安装了 ocr 扩展（`pip install doc2ppt[ocr]`）时生效；未安装时
`page_paras()` 返回空列表，parse 层维持「提示先 OCR」的原行为，不报错。
"""

def available():
    """OCR 扩展是否可用（延迟导入，基础安装不引入 onnxruntime）。"""
    try:
        import rapidocr_onnxruntime  # noqa: F401
    except ImportError:
        return False
    return True


_ENGINE = None


def _engine():
    """进程内共享一个 OCR 实例（模型加载约 1s）。"""
    global _ENGINE
    if _ENGINE is None:
        from rapidocr_onnxruntime import RapidOCR

        _ENGINE = RapidOCR()
    return _ENGINE


def page_paras(page):
    """把无文本层的页面过 OCR，聚合为段落，供 parse 层分类。

    返回 [{"text", "h", "h_ratio"}]（h_ratio = 段高/全页中位行高）；
    OCR 不可用时返回 []。表格无法从 OCR 可靠还原，保持提示人工处理。
    """
    if not available():
        return []
    engine = _engine()
    pm = page.get_pixmap(dpi=200)  # 扫描件常规分辨率，速度/精度折中
    result, _elapse = engine(pm.tobytes("png"))
    lines = []
    for box, text, score in result or []:
        try:
            score = float(score)
        except (TypeError, ValueError):
            score = 0.0
        if score < 0.5 or not str(text).strip():
            continue
        xs = [p[0] for p in box]
        ys = [p[1] for p in box]
        lines.append({"y0": min(ys), "y1": max(ys), "x0": min(xs), "x1": max(xs),
                      "h": max(ys) - min(ys), "text": str(text).strip()})
    if not lines:
        return []
    lines.sort(key=lambda l: (round(l["y0"], 1), l["x0"]))
    med_h = sorted(l["h"] for l in lines)[len(lines) // 2] or 1.0

    # 同段判定：垂直间隙 < 0.9×中位行高 且 水平区间重叠
    paras = []
    for l in lines:
        cur = paras[-1] if paras else None
        if (cur is not None
                and l["y0"] - cur["y1"] < 0.9 * med_h
                and min(cur["x1"], l["x1"]) - max(cur["x0"], l["x0"]) > 0):
            cur["y1"] = max(cur["y1"], l["y1"])
            cur["x0"] = min(cur["x0"], l["x0"])
            cur["x1"] = max(cur["x1"], l["x1"])
            cur["text"] += l["text"]  # 中文行拼接不加空格
        else:
            paras.append({"y0": l["y0"], "y1": l["y1"], "x0": l["x0"],
                          "x1": l["x1"], "text": l["text"], "fh": l["h"]})
    for p in paras:
        # 用首行高做字号代理：多行段整段高会翻倍，不能直接当标题信号
        p["h_ratio"] = p["fh"] / med_h
    return paras
