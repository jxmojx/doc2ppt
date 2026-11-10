"""M1 解析器：docx / pdf（有文本层）/ md → Block 序列。

统一出口 parse_document(path)：按扩展名分派，回填 seq，返回 blocks。
"""
import re
from pathlib import Path

import pymupdf
from docx import Document
from docx.oxml.ns import qn as dqn
from docx.table import Table
from docx.text.paragraph import Paragraph
from markdown_it import MarkdownIt

from .guard import guard_ooxml
from .models import Block

# 中文章节编号模式（PDF 无样式时用于层级判断）
PDF_NUM_H1 = re.compile(r"^(第[一二三四五六七八九十百]+[章节篇部]|[一二三四五六七八九十]+、)")
PDF_NUM_H2 = re.compile(r"^(\d+\.\d+|（[一二三四五六七八九十]+）|\(\d+\)|\d+、)")

# docx 标题：样式名模式（中英文）与内置列表样式
DOCX_TITLE_NAMES = {"title", "标题"}
DOCX_HEADING_RE = re.compile(r"^(?:heading|标题)\s*([1-9])$")
DOCX_LIST_HINT = ("list", "列表")


def parse_document(path):
    """按扩展名分派解析器，回填序号。"""
    ext = Path(path).suffix.lower()
    fn = _SUFFIX_MAP.get(ext)
    if fn is None:
        raise ValueError(f"不支持的文件类型: {ext}（支持 .docx / .pdf / .md）")
    guard_ooxml(path)  # 防 XXE：含 DOCTYPE/ENTITY 的 OOXML 直接拒绝
    blocks = fn(str(path))
    for i, b in enumerate(blocks):
        b.seq = i
    return blocks


# -------------------- docx --------------------

def parse_docx(path):
    """docx → blocks。

    标题识别不依赖英文样式名：① 段落直接 outlineLvl ② 样式链上的样式名
    （heading N / 标题 N / Title / 标题）③ 样式链上的 outlineLvl（自定义样式）。
    列表识别同时认「List/列表」样式与 numPr 自动编号。
    图片不静默丢弃：记为 image 块（含数量），渲染层以备注提示。
    """
    doc = Document(path)
    blocks = []
    for child in doc.element.body.iterchildren():
        if child.tag.endswith("}p"):
            para = Paragraph(child, doc)
            text = para.text.strip()
            n_images = len(child.findall(".//" + dqn("w:drawing")))
            if not text:
                if n_images:
                    blocks.append(Block("image", meta={"count": n_images,
                                                       "images": _docx_images(child, doc)}))
                # 纯文本框段落：Word 里常用来放要点/引文，此前被整体丢弃
                for tb_text in _docx_textboxes(child):
                    blocks.append(Block("para", tb_text))
                continue
            level = _docx_level(para)
            if level is not None:
                blocks.append(Block("heading", text, level=level))
            elif _docx_is_list(para):
                blocks.append(Block("bullet", text, meta={"list_type": "docx"}))
            else:
                blocks.append(Block("para", text))
            if n_images:
                blocks.append(Block("image", meta={"count": n_images,
                                                   "images": _docx_images(child, doc)}))
            for tb_text in _docx_textboxes(child):
                blocks.append(Block("para", tb_text))
        elif child.tag.endswith("}tbl"):
            table = Table(child, doc)
            rows = [[c.text.strip() for c in row.cells] for row in table.rows]
            rows = [r for r in rows if any(r)]
            if len(rows) >= 2:
                blocks.append(Block("table", meta={"headers": rows[0], "rows": rows[1:]}))
    # 页眉常是文档正式名称；正文没有 Title/一级标题时用它当封面标题，
    # 比文件名（如「XX终稿2」）更可读。已有 level-0 标题则不覆盖。
    header = _docx_header_text(doc)
    if header and not any(b.block_type == "heading" and b.level == 0 for b in blocks):
        blocks.insert(0, Block("heading", header, level=0))
    return blocks


_TEXTBOX_P = dqn("w:txbxContent")


def _docx_textboxes(para_el):
    """取段落内文本框的段落文字（按文档顺序）。

    python-docx 的 Paragraph.text 只读 w:p 直属 run，文本框内容在
    w:drawing → wps:txbx → w:txbxContent 里，不单独提取就会被静默丢掉。
    """
    out = []
    for box in para_el.findall(".//" + _TEXTBOX_P):
        for p in box.findall(dqn("w:p")):
            text = "".join(t.text or "" for t in p.findall(".//" + dqn("w:t"))).strip()
            if text:
                out.append(text)
    return out


_HEADER_NOISE = re.compile(r"^[\d\s\-—–·.]+$|^第\s*\d+\s*页|^page\s*\d+|^\d{4}[-/.年]\d{1,2}")


def _docx_header_text(doc):
    """取可当标题的页眉文字（跳过页码、日期一类噪音；最多 40 字）。"""
    for section in doc.sections:
        try:
            paras = section.header.paragraphs
        except (KeyError, AttributeError, ValueError):  # 非标准 header 部件不影响正文
            continue
        for p in paras:
            text = p.text.strip()
            if 2 <= len(text) <= 40 and not _HEADER_NOISE.search(text):
                return text
    return ""


def _docx_images(para_el, doc):
    """取段落内嵌图片的二进制与扩展名（用于真正嵌入 PPT，而不是只记数量）。"""
    out = []
    for blip in para_el.findall(".//" + dqn("a:blip")):
        rid = blip.get(dqn("r:embed"))
        if not rid:
            continue
        try:
            part = doc.part.related_parts[rid]
        except KeyError:
            continue
        ext = str(part.partname).rsplit(".", 1)[-1].lower()
        if ext in ("png", "jpg", "jpeg", "gif", "bmp", "tiff", "tif"):
            out.append({"blob": part.blob, "ext": ext})
    return out


def _outline_lvl(el):
    """读元素 w:pPr/w:outlineLvl，返回 0-8（9 表示正文级，视为无层级）。"""
    ppr = el.find(dqn("w:pPr"))
    if ppr is None:
        return None
    ol = ppr.find(dqn("w:outlineLvl"))
    if ol is None:
        return None
    try:
        val = int(ol.get(dqn("w:val")))
    except (TypeError, ValueError):
        return None
    return None if val >= 9 else val


def _docx_level(para):
    """段落标题层级：0=文档标题，1..9=层级，None=正文。"""
    lvl = _outline_lvl(para._p)  # 段落直接设置优先
    if lvl is not None:
        return lvl + 1
    style, seen = para.style, set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        name = (style.name or "").strip()
        low = name.lower()
        if low in DOCX_TITLE_NAMES or name in DOCX_TITLE_NAMES:
            return 0
        m = DOCX_HEADING_RE.match(low) or DOCX_HEADING_RE.match(name)
        if m:
            return int(m.group(1))
        lvl = _outline_lvl(style.element)  # 自定义样式靠 outlineLvl
        if lvl is not None:
            return lvl + 1
        style = style.base_style
    return None


def _docx_is_list(para):
    """列表判定：段落或样式链上有 numPr，或样式名含 List/列表。"""
    ppr = para._p.find(dqn("w:pPr"))
    if ppr is not None and ppr.find(dqn("w:numPr")) is not None:
        return True
    style, seen = para.style, set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        name = (style.name or "").strip()
        if any(h in name.lower() for h in DOCX_LIST_HINT):
            return True
        sppr = style.element.find(dqn("w:pPr"))
        if sppr is not None and sppr.find(dqn("w:numPr")) is not None:
            return True
        style = style.base_style
    return False


# -------------------- pdf --------------------

def parse_pdf(path):
    """pdf（有文本层）→ blocks。

    无样式信息，标题判定用「相对字号 + 粗体 + 编号模式」三信号：
      1. 先统计正文字号（按字符数加权的众数），得到相对倍率；
      2. 倍率 ≥1.15 视为标题；粗体且短行（≤40 字）且不矮于正文亦视为标题；
      3. 层级优先按编号模式（一、/第X章 → 1；1.1/（一）→ 2），否则按倍率。
    同字号加粗标题（真实 PDF 常见）因此可被识别。
    """
    doc = pymupdf.open(path)
    raw = []  # [(kind, payload)] 按阅读顺序；line 段已合并为段落
    for page in doc:
        # 扫描页判定：无文本层。整页大图只是扫描容器，不作为插图输出
        scanned_page = not page.get_text("text").strip()
        tables = _page_tables(page)
        tab_rects = [t.bbox for t in tables]

        def in_table(bb, tab_rects=tab_rects):
            cx, cy = (bb[0] + bb[2]) / 2, (bb[1] + bb[3]) / 2
            return any(r[0] <= cx <= r[2] and r[1] <= cy <= r[3] for r in tab_rects)

        entries = []
        for t in tables:
            data = [[_clean_cell(c) for c in row] for row in t.extract()]
            data = [r for r in data if any(r)]
            if len(data) >= 2:
                entries.append((t.bbox[1], t.bbox[0], "table", data))
        # 无边框表格兜底：列对齐 + 多空格分列的行聚成表（真实报告常见）
        soft_tables = _borderless_tables(page, tab_rects)
        for bbox, data in soft_tables:
            entries.append((bbox[1], bbox[0], "table", data))
            tab_rects.append(bbox)
        # 图片：按 bbox 插入阅读顺序（与 docx 一致，供渲染层真嵌入）
        if not scanned_page:
            for info in page.get_image_info(xrefs=True):
                img = _pdf_image(doc, info.get("xref"))
                if img:
                    bb = info["bbox"]
                    entries.append((bb[1], bb[0], "image", img))
        for b in page.get_text("dict")["blocks"]:
            if b.get("type") != 0:
                continue
            for line in b["lines"]:
                text = "".join(s["text"] for s in line["spans"]).strip().replace("\xa0", " ")
                if not text or in_table(line["bbox"]):
                    continue
                size = max((s["size"] for s in line["spans"]), default=11.0)
                bold = _line_bold(line["spans"])
                bb = line["bbox"]
                entries.append((bb[1], bb[0], "line", (text, size, bold)))
        entries.sort(key=lambda e: (e[0], e[1]))
        # 行级合并成段落：同字号、相对字号+粗体+编号三信号分类
        cur = []
        for y0, x0, kind, payload in entries:
            if kind == "table":
                _push_merged(raw, cur)
                cur = []
                raw.append(("table", payload))
                continue
            if kind == "image":
                _push_merged(raw, cur)
                cur = []
                raw.append(("image", payload))
                continue
            text, size, bold = payload
            same_line = (
                cur
                and abs(size - cur[-1][1]) < 0.5
                and abs(bold - cur[-1][4]) < 0.3  # 粗体状态不同不合并（同字号加粗标题）
                and y0 - cur[-1][2] < size * 1.6 + 2
                and abs(x0 - cur[0][3]) < 3
            )
            if same_line:
                cur.append((text, size, y0, x0, bold))
            else:
                _push_merged(raw, cur)
                cur = [(text, size, y0, x0, bold)]
        _push_merged(raw, cur)
        # OCR 兜底：无文本层的页（扫描件）且有 ocr 扩展时启用；
        # OCR 不可用则本页产出 0 块（维持「先 OCR」提示行为）
        if scanned_page:
            from . import ocr
            for i, p in enumerate(ocr.page_paras(page), 1):
                raw.append(("line", (p["text"], 11.0 * p["h_ratio"],
                                     float(i) * 20, 0.0, False)))
    doc.close()

    body_size = _pdf_body_size(raw)
    blocks = []
    for kind, payload in raw:
        if kind == "table":
            blocks.append(Block("table", meta={"headers": payload[0], "rows": payload[1:]}))
        elif kind == "image":
            blocks.append(Block("image", meta={"count": 1, "images": [payload]}))
        else:
            blocks.append(_pdf_classify(*payload, body_size=body_size))
    return blocks


def _pdf_image(doc, xref):
    """取 PDF 内嵌图片的二进制与扩展名（渲染层按比例嵌入）。"""
    if not xref:
        return None
    try:
        img = doc.extract_image(int(xref))
    except (ValueError, RuntimeError):
        return None
    ext = (img.get("ext") or "").lower()
    if ext in ("png", "jpg", "jpeg", "gif", "bmp", "tiff", "tif"):
        return {"blob": img["image"], "ext": ext}
    return None


def _page_tables(page):
    """安全取表格：pymupdf 的表格识别依赖 ONNX，资源紧张时会异常或返回 None。

    失败时按「本页无表格」处理（正文仍可正常提取），不让整份文档解析崩掉。
    """
    try:
        found = page.find_tables()
    except Exception:  # noqa: BLE001 - 第三方识别器可能抛任意异常
        return []
    tables = getattr(found, "tables", None)
    return tables or []


def _borderless_tables(page, occupied):
    """识别无边框表格：连续若干行的「列数相同、首列左边界对齐」即视为表格。

    判据保守（要求 ≥3 行、≥2 列、首列 x 对齐误差 <6pt），宁可漏检也不误判正文。
    返回 [(bbox, [[cell,...], ...])]。
    """
    rows = []
    for b in page.get_text("dict")["blocks"]:
        if b.get("type") != 0:
            continue
        for line in b["lines"]:
            bb = line["bbox"]
            if _inside(bb, occupied):
                continue
            text = "".join(s["text"] for s in line["spans"]).strip().replace("\xa0", " ")
            parts = [p for p in re.split(r"\s{2,}", text) if p.strip()]
            if len(parts) >= 2:
                rows.append((bb[1], bb[0], parts, bb))
    rows.sort()
    out, i = [], 0
    while i < len(rows):
        j = i + 1
        base_x, ncol = rows[i][1], len(rows[i][2])
        while j < len(rows) and abs(rows[j][1] - base_x) < 6 and len(rows[j][2]) >= ncol:
            j += 1
        if j - i >= 3:  # 至少 3 行同构才算表格
            data = [rows[k][2] for k in range(i, j)]
            xs = [rows[k][3][0] for k in range(i, j)]
            x1 = max(rows[k][3][2] for k in range(i, j))
            y0, y1 = rows[i][3][1], rows[j - 1][3][3]
            if len(data) >= 2 and any(row for row in data[1:]):
                out.append(((min(xs), y0, x1, y1), data))
            i = j
        else:
            i += 1
    return out


def _inside(bbox, rects):
    cx, cy = (bbox[0] + bbox[2]) / 2, (bbox[1] + bbox[3]) / 2
    return any(r[0] <= cx <= r[2] and r[1] <= cy <= r[3] for r in rects)


def _line_bold(spans):
    """行内粗体字符占比（flags bit4 或字体名含 Bold/Black/Heavy/Semibold）。"""
    total = sum(len(s["text"]) for s in spans) or 1
    bold = 0
    for s in spans:
        name = s.get("font", "")
        is_bold = bool(s.get("flags", 0) & 16) or any(
            k in name for k in ("Bold", "Black", "Heavy", "Semibold"))
        if is_bold:
            bold += len(s["text"])
    return bold / total


def _pdf_body_size(raw):
    """正文字号：按字符数加权的众数（同票取较小者）。"""
    weight = {}
    for kind, payload in raw:
        if kind != "line":
            continue
        text, size, _bold, _y, _x = payload
        weight[size] = weight.get(size, 0) + len(text)
    if not weight:
        return 11.0
    return max(weight.items(), key=lambda kv: (kv[1], -kv[0]))[0]


def _pdf_classify(text, size, bold, y0, x0, body_size=11.0):
    """按相对字号 / 粗体 / 编号模式判定 heading 层级或正文。"""
    ratio = size / body_size if body_size else 1.0
    is_heading = ratio >= 1.15 or (bold >= 0.6 and len(text) <= 40 and ratio >= 1.0)
    if not is_heading:
        return Block("para", text)
    if PDF_NUM_H2.match(text):
        level = 2
    elif PDF_NUM_H1.match(text):
        level = 1
    elif ratio >= 1.5:
        level = 0 if _looks_title(text) else 1
    else:
        level = 2
    return Block("heading", text, level=level)


def _looks_title(text):
    """文档标题特征：无章节编号、较短。"""
    return not (PDF_NUM_H1.match(text) or PDF_NUM_H2.match(text)) and len(text) <= 40


def _clean_cell(text):
    """归一化 PDF 表格单元格：去 nbsp、合并折行、剔除折行产生的纯标点碎片。"""
    if not text:
        return ""
    parts = [p.strip() for p in text.replace("\xa0", " ").split("\n")]
    kept = [p for p in parts if p and not all(ch in "、，。；：,.;:" for ch in p)]
    return " ".join(kept).strip()


def _push_merged(out, cur):
    """把行级合并缓冲转成一个段落条目：line 段为 (text, size, bold, y, x)。"""
    if not cur:
        return
    text = "".join(c[0] for c in cur)
    size = cur[0][1]
    bold = sum(c[4] * len(c[0]) for c in cur) / max(len(text), 1)
    out.append(("line", (text, size, bold, cur[0][2], cur[0][3])))


# -------------------- md --------------------

def parse_md(path):
    md = MarkdownIt("commonmark").enable("table")
    tokens = md.parse(Path(path).read_text(encoding="utf-8"))
    return _md_blocks(tokens)


def _md_blocks(tokens):
    blocks = []
    table_ctx = None
    i, n = 0, len(tokens)
    while i < n:
        t = tokens[i]
        if t.type == "heading_open":
            content = tokens[i + 1].content if tokens[i + 1].type == "inline" else ""
            if content.strip():
                blocks.append(Block("heading", _md_inline(content), level=int(t.tag[1])))
            i += 3
        elif t.type == "paragraph_open":
            if table_ctx is None:
                content = tokens[i + 1].content if tokens[i + 1].type == "inline" else ""
                if content.strip():
                    blocks.append(Block("para", _md_inline(content)))
            i += 3
        elif t.type in ("fence", "code_block"):
            # 代码块：保留原有换行（渲染层转 <a:br/>），不再静默丢弃
            code = t.content.rstrip("\n")
            if code.strip():
                blocks.append(Block("para", code))
            i += 1
        elif t.type in ("bullet_list_open", "ordered_list_open"):
            i = _md_list(tokens, i, t.type == "ordered_list_open", blocks)
        elif t.type == "table_open":
            table_ctx = {"headers": [], "rows": []}
            i += 1
        elif t.type == "table_close":
            if table_ctx and table_ctx["rows"]:
                blocks.append(Block("table", meta=dict(table_ctx)))
            table_ctx = None
            i += 1
        elif t.type == "tr_open" and table_ctx is not None:
            row = []
            i += 1
            while i < n and tokens[i].type != "tr_close":
                if tokens[i].type in ("th_open", "td_open"):
                    cell = tokens[i + 1].content if tokens[i + 1].type == "inline" else ""
                    row.append(_md_cell(cell))
                    i += 3
                else:
                    i += 1
            i += 1
            if any(row):
                if not table_ctx["headers"]:
                    table_ctx["headers"] = row
                else:
                    table_ctx["rows"].append(row)
        else:
            i += 1
    return blocks


def _md_list(tokens, i, ordered, blocks):
    """列表块 → 独立 bullet。嵌套列表递归展开（本项目版式无多级列表，
    子项作为同级条目紧随父项），避免把父项与子项拼成一条。"""
    close = "ordered_list_close" if ordered else "bullet_list_close"
    i += 1  # 跳过 list_open
    while i < len(tokens) and tokens[i].type != close:
        if tokens[i].type == "list_item_open":
            i += 1
            parts, nested = [], []
            while i < len(tokens) and tokens[i].type != "list_item_close":
                if tokens[i].type in ("bullet_list_open", "ordered_list_open"):
                    sub_ordered = tokens[i].type == "ordered_list_open"
                    i = _md_list(tokens, i, sub_ordered, nested)
                    continue
                if tokens[i].type == "inline":
                    parts.append(tokens[i].content.strip())
                i += 1
            content = _md_inline(" ".join(p for p in parts if p).strip())
            if content:
                blocks.append(Block("bullet", content,
                                    meta={"list_type": "ordered" if ordered else "bullet"}))
            blocks.extend(nested)
            i += 1  # 跳过 list_item_close
        else:
            i += 1
    return i + 1


def _md_cell(content):
    """表格单元格：还原转义竖线，折叠软换行，去行内反引号。"""
    return _md_inline(content)


def _md_inline(text):
    """行内文本归一：折叠软换行、还原转义竖线、去行内代码反引号。"""
    text = _soft_join(text).replace("\\|", "|")
    return re.sub(r"`([^`]*)`", r"\1", text).strip()


def _soft_join(text):
    """markdown 软换行（段落内 \\n）拼接为一行，避免渲染成硬换行。"""
    return text.replace("\n", "")


_SUFFIX_MAP = {".docx": parse_docx, ".pdf": parse_pdf, ".md": parse_md}


# -------------------- 打印（M1 验收用） --------------------

def print_blocks(blocks):
    """调试工具（非管线用）：按标题层级打印 blocks 树，排查解析结果。"""
    for b in blocks:
        if b.block_type == "heading":
            print(f"[{b.seq:>3}] {'  ' * b.level}▸ {b.text}")
        elif b.block_type == "table":
            h = b.meta.get("headers", [])
            print(f"[{b.seq:>3}]     ┌ table {len(b.meta.get('rows', [])) + 1}行x{len(h)}列"
                  f" | 表头: {'|'.join(h)}")
        elif b.block_type == "image":
            print(f"[{b.seq:>3}]     ▣ 图片 ×{b.meta.get('count', 1)}（未嵌入，仅记录）")
        else:
            label = "•" if b.block_type == "bullet" else "·"
            text = b.text.replace("\n", " ")
            print(f"[{b.seq:>3}]     {label} {text[:60]}{'…' if len(text) > 60 else ''}")
