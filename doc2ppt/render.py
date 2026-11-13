"""M3 渲染器：大纲 → 模板槽位填充。

每次渲染从模板副本开始（幂等：不在上次输出上叠加），正文超长时按
item.font_size 缩字号（下限 14pt）；缩到下限仍放不下由 M5 校验环处理。
"""
import math
import re
from pathlib import Path

from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

from .template import load_template

FONT_PRIMARY = "思源黑体 CN"
FONT_FILES = {
    "思源黑体 CN": ["SourceHanSansCN-Regular.otf", "SourceHanSansSC-Regular.otf",
                    "SourceHanSansCN-Normal.otf", "SourceHanSansCN-Regular.ttf"],
    "微软雅黑": ["msyh.ttc", "msyh.ttf"],
    "黑体": ["simhei.ttf"],
    "宋体": ["simsun.ttc", "simsun.ttf"],
}
C_TITLE, C_BODY, C_SUB, C_CAPT = "1F4E79", "333333", "595959", "7F7F7F"
C_HEAD, C_ZEBRA = "DEEBF7", "F4F8FC"
TABLE_ROW_IN = 0.33  # 表格单行高度（英寸，≈1.6 倍 14pt 字号）
INTRO_PT = 16        # 表格页导语字号
DEFAULT_BODY = 20
MIN_BODY = 14
TITLE_SIZE = 32
COVER_TITLE_SIZE = 40
TABLE_MAX_COLS, TABLE_MAX_ROWS = 6, 8
EMU = 914400


def resolve_font():
    """白名单选本机可用中文字体，返回 (字体名, 字体文件路径或 None, 降级警告或 None)。"""
    fonts_dir = Path("C:/Windows/Fonts")
    try:
        installed = {f.name.lower(): f for f in fonts_dir.iterdir() if f.is_file()}
    except OSError:
        installed = {}
    for name, files in FONT_FILES.items():
        for f in files:
            if f.lower() in installed:
                warn = None if name == FONT_PRIMARY else f"字体缺失: {FONT_PRIMARY} 不可用，回退 {name}"
                return name, installed[f.lower()], warn
    return "宋体", None, "字体缺失: 白名单字体均不可用，回退 宋体"


def pick_font():
    """返回 (字体名, 降级警告或 None)。"""
    name, _path, warn = resolve_font()
    return name, warn


SEMANTICS = ("title", "content", "two_col", "chart")
TWO_COL_THRESHOLD = 6  # 列表项 ≥ 该值走两栏


def pick_layout(item) -> str:
    """版式选型规则：封面→title；含表格→chart；纯列表页（≥6 列表项且
    页内 ≤1 个小节标题）→two_col；其余（含多节合并页）→content。"""
    if item.layout_hint in SEMANTICS:
        return item.layout_hint
    if item.page_no == 1:
        return "title"
    if any(b.block_type == "table" for b in item.blocks):
        return "chart"
    bullets = sum(1 for b in item.blocks if b.block_type == "bullet")
    headings = sum(1 for b in item.blocks if b.block_type == "heading")
    if bullets >= TWO_COL_THRESHOLD and headings <= 1:
        return "two_col"
    return "content"


def render(items, template_path, out_path, warnings=None):
    """渲染整份 deck。每次从模板副本开始；返回 warnings 列表。"""
    warnings = warnings if warnings is not None else []
    # 封面页数按最终页数校正（修复环可能拆页/删空页，outline 生成时还不知道）
    if items:
        count_text = f"共 {len(items)} 页"
        items[0].thesis = re.sub(r"共 \d+ 页", count_text, items[0].thesis or "")
        if items[0].notes:
            items[0].notes = re.sub(r"共 \d+ 页", count_text, items[0].notes)
    prs, infos = load_template(template_path)
    font, warn = pick_font()
    if warn:
        warnings.append(warn)
    total = len(items)
    for i, item in enumerate(items, 1):
        sem = pick_layout(item)
        info = infos.get(sem) or infos["content"]
        slide = prs.slides.add_slide(info.layout)
        _fill_slide(slide, item, sem, info, font, warnings)
        _fill_notes(slide, item)
        # 页脚信息条：封面为「日期 · 共 N 页」，内容页为页码——交付即用，无需手补
        stamp = f"{_today_cn()} · 共 {total} 页" if i == 1 else f"{i} / {total}"
        _add_stamp(slide, prs.slide_width, prs.slide_height, stamp, font)
    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out)
    return warnings


def _today_cn():
    import datetime

    d = datetime.date.today()
    return f"{d.year} 年 {d.month} 月"


def _add_stamp(slide, slide_w, slide_h, text, font):
    """右下角信息条（10.5pt 灰字，右对齐；放在正文槽之下，不与内容重叠）。"""
    w_in, h_in = 1.6, 0.3
    box = slide.shapes.add_textbox(
        Emu(int((slide_w / 914400 - 0.8 - w_in) * 914400)),
        Emu(int((slide_h / 914400 - 0.40) * 914400)),
        Emu(int(w_in * 914400)), Emu(int(h_in * 914400)))
    tf = box.text_frame
    tf.word_wrap = False
    box.name = "doc2ppt_stamp"  # verify 对 doc2ppt_* 自绘形状豁免字号/扫描
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    para = tf.paragraphs[0]
    run = para.add_run()
    _set_run_text(para, run, text, 10.5, False, C_CAPT, font)
    para._p.get_or_add_pPr().set("algn", "r")


def _ph(slide, idx):
    """按占位符 idx 取形状（模板无关），不存在返回 None。"""
    if idx is None:
        return None
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == idx:
            return ph
    return None


def _fill_slide(slide, item, sem, info, font, warnings):
    title_ph = _ph(slide, info.idx("title"))
    if sem == "title":
        if title_ph is not None:
            _set_para(title_ph, item.title, COVER_TITLE_SIZE, C_TITLE, bold=True, font=font)
        sub = _ph(slide, info.idx("subtitle"))
        if sub is not None:
            _set_para(sub, item.thesis, 20, C_SUB, font=font)
        return
    if title_ph is not None:
        tsize = item.title_size or TITLE_SIZE
        _set_para(title_ph, item.title, tsize, C_TITLE, bold=True, font=font, center=True)
    if sem == "chart" and any(b.block_type == "table" for b in item.blocks):
        _fill_chart_slide(slide, item, info, font, warnings)
    else:
        _fill_content_slide(slide, item, sem, info, font)


def _fill_content_slide(slide, item, sem, info, font):
    size = item.font_size or DEFAULT_BODY
    body = _ph(slide, info.idx("body"))
    if body is None:
        body = _ph(slide, info.idx("body2"))
    if body is None:
        return
    body2 = _ph(slide, info.idx("body2"))
    if sem == "two_col" and body2 is not None:
        spec = info.slots.get("body") or info.slots.get("body2")
        left, right = _split_blocks(item.blocks, size, spec.w if spec else None,
                                    title=item.title)
        _fill_body(body, left, size, font, title=item.title)
        _fill_body(body2, right, size, font, title=item.title)
        # 双栏各自按本栏内容做垂直居中/段距平衡（用整页块会高估、导致不触发）
        _balance_blocks(body, left, size, info)
        _balance_blocks(body2, right, size, info)
    else:
        # 无两栏版式时降级为单栏顺序填充
        _fill_body(body, item.blocks, size, font, title=item.title)
        _balance_short_page(body, item, size, info)
        _place_images(slide, item, info, size)


def _balance_short_page(body, item, size, info):
    """内容明显不足时：段距填充 /  statement 页居中（见 _balance_blocks）。

    与页标题重复的小标题渲染时会被跳过，估算前先剔除，否则高度虚高
    会把该居中的页顶在阈值之上（视觉验收实测 slide02 顶置 fail）。
    """
    blocks = [b for b in item.blocks
              if not (b.block_type == "heading" and _same_title(b.text, item.title))]
    _balance_blocks(body, blocks, size, info)


def _place_images(slide, item, info, size):
    """把原文档图片按比例嵌入正文下方剩余空间（空间不足则只留备注提示）。"""
    from . import verify as verify_mod

    imgs = [im for b in item.blocks if b.block_type == "image"
            for im in (b.meta.get("images") or [])]
    spec = info.slots.get("body")
    if not imgs or spec is None:
        return
    need_pt = verify_mod.estimate_body_need(item.blocks, size, spec.w, spec.h)
    free_in = spec.h - need_pt / 72 - 0.15
    if free_in < 1.0:  # 正文已占满，图片留给人工补充
        return
    from io import BytesIO

    import pymupdf

    x_in, y_in, w_in = spec.x, spec.y + need_pt / 72 + 0.12, spec.w
    for im in imgs[:2]:  # 最多两张，避免堆叠溢出
        try:
            pix = pymupdf.Pixmap(im["blob"])
        except (ValueError, RuntimeError):
            continue
        ratio = pix.width / max(pix.height, 1)
        h_in = min(free_in, w_in / ratio)
        w_use = h_in * ratio
        left = x_in + (w_in - w_use) / 2
        try:
            slide.shapes.add_picture(BytesIO(im["blob"]), Emu(int(left * 914400)),
                                     Emu(int(y_in * 914400)),
                                     Emu(int(w_use * 914400)), Emu(int(h_in * 914400)))
        except (ValueError, TypeError):
            continue
        y_in += h_in + 0.1
        free_in -= h_in + 0.1
        if free_in < 1.0:
            break


def _balance_blocks(body, blocks, size, info):
    """按给定块集合做短内容平衡（双栏时传本栏块）。

    顶对齐 + 段间距填充：不用垂直居中——居中会让正文悬在页中、
    标题下出现大片空白带（视觉验收实测被判定 fail）。
    """
    from . import verify as verify_mod

    spec = info.slots.get("body")
    if spec is None or not blocks:
        return
    need = verify_mod.estimate_body_need(blocks, size, spec.w, spec.h)
    if need >= 0.6 * spec.h * 72:
        return
    paras = [p for p in body.text_frame.paragraphs if "".join(r.text for r in p.runs).strip()]
    if len(paras) < 2:
        return
    # 目标：段距填充到约 72% 版心；填充后仍 <55% 的极稀疏页改垂直居中（statement 页），
    # 避免底部整片空白（视觉验收对「下坠悬中」与「底部空洞」都判 fail）。
    extra_pt = (0.72 * spec.h * 72 - need) / max(len(paras) - 1, 1)
    extra_pt = max(0.0, min(extra_pt, size * 2.0))  # 段距上限 2 倍字号
    if extra_pt >= 2:
        for p in paras:
            p.space_after = Pt(extra_pt)
    fill_after = need + extra_pt * (len(paras) - 1)
    if fill_after < 0.55 * spec.h * 72:
        body.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE


def _split_blocks(blocks, size=None, width=None, title=None):
    """两栏页分栏：按**渲染高度**对半切（保序），标题不落单在栏底。

    旧实现按累计字数切，长短条目混排时两栏高度差可达数行。给定字号与栏宽时，
    用解析式估算每块所需高度并取「左右高度差最小」的切点；否则退回字数切分。
    与页标题重复的小标题渲染时本就不会出现，分栏时先剔除，否则高度会算多。
    """
    content = [b for b in blocks if b.block_type != "table"
               and not (b.block_type == "heading" and title
                        and _same_title(b.text, title))]
    if not content:
        return [], []
    if size and width:
        from . import verify as verify_mod

        needs = [verify_mod.estimate_body_need([b], size, width, 0) for b in content]
        prefix = [0.0]
        for n in needs:
            prefix.append(prefix[-1] + n)
        total = prefix[-1]
        best_i, best_diff = 1, None
        for i in range(1, len(content)):
            diff = abs(total - 2 * prefix[i])
            if best_diff is None or diff < best_diff:
                best_i, best_diff = i, diff
        # 左栏以标题收尾会成为孤行标题 → 把标题带到右栏开头
        if content[best_i - 1].block_type == "heading" and best_i < len(content):
            best_i += 1
        if 0 < best_i < len(content):
            return content[:best_i], content[best_i:]
    total = sum(len(b.text) for b in content)
    acc, split_at = 0, 0
    for i, b in enumerate(content):
        acc += len(b.text)
        if acc >= total / 2:
            split_at = i + 1
            break
    # 切点前一块是标题 → 后移一格，保持标题与内容同栏
    if split_at < len(content) and content[split_at - 1].block_type == "heading":
        split_at += 1
    left, right = content[:split_at], content[split_at:]
    if not left or not right:  # 兜底：均分数量
        half = (len(content) + 1) // 2
        left, right = content[:half], content[half:]
    return left, right


def _num(cell):
    """从单元格里取数值：'87.2%'→87.2、'≥85%'→85、'18秒'→18；无数字返回 None。"""
    m = _NUM_RE.search(str(cell))
    if not m:
        return None
    try:
        return float(m.group())
    except ValueError:
        return None


_NUM_RE = re.compile(r"[-+]?\d+(?:\.\d+)?")


def _unit(cell):
    """取单元格数值后的单位（'≥85%'→'%'、'18秒'→'秒'、'8周'→'周'）；无数字返回 None。"""
    m = _NUM_RE.search(str(cell))
    if not m:
        return None
    return str(cell)[m.end():].strip().lower()


def _chart_spec(headers, rows):
    """判断表格能否转成原生图表，返回 spec 或 None。

    条件：≥2 行、1 个分类列 + 至少 1 个全数值的数据列；最多取 2 个数据列；
    **所有数值单位必须一致**（% 与 秒 混在同一坐标轴会误导，宁可不画）。
    类型：分类值像时间（年/月/季/第N期）→ 折线；单列且都在 0-100 且合计≈100 → 饼图；
    其余 → 柱状图（对比）。
    """
    if len(headers) < 2 or len(rows) < 2 or len(rows) > 8:
        return None
    cats = [str(r[0]).strip() for r in rows]
    if not all(cats) or max(len(c) for c in cats) > 20:
        return None
    series, units = [], set()
    for j in range(1, len(headers)):
        cells = [r[j] if j < len(r) else "" for r in rows]
        vals = [_num(c) for c in cells]
        if not all(v is not None for v in vals):
            continue
        col_units = {u for u in (_unit(c) for c in cells) if u is not None}
        series.append((str(headers[j]).strip() or f"系列{j}", vals))
        units |= col_units
    if not series or len(units) > 1:
        return None  # 无可用数值列，或单位不一致
    series = series[:2]
    if _looks_temporal(cats):
        kind = "line"
    elif len(series) == 1 and all(0 <= v <= 100 for v in series[0][1]) \
            and 95 <= sum(series[0][1]) <= 105:
        kind = "pie"
    else:
        kind = "column"
    return {"kind": kind, "categories": cats, "series": series,
            "unit": next(iter(units)) if units else ""}


_TEMPORAL_RE = re.compile(r"(\d{4}\s*年|^\d{1,2}\s*月|^[Qq][1-4]|第[一二三四1-4]季度|"
                          r"^\d{4}[-/]\d{1,2}|^\d{1,2}月$|上旬|中旬|下旬|上半年|下半年)")


def _looks_temporal(cats):
    return sum(1 for c in cats if _TEMPORAL_RE.search(c)) >= max(2, len(cats) // 2)


def _draw_chart(slide, spec, box, font_size=14):
    """插入原生图表（可在 PowerPoint 里「编辑数据」）。"""
    data = CategoryChartData()
    data.categories = spec["categories"]
    for name, vals in spec["series"]:
        data.add_series(name, vals)
    kinds = {"column": XL_CHART_TYPE.COLUMN_CLUSTERED,
             "pie": XL_CHART_TYPE.PIE,
             "line": XL_CHART_TYPE.LINE_MARKERS}
    gf = slide.shapes.add_chart(kinds[spec["kind"]], Emu(int(box["left"])),
                                Emu(int(box["top"])), Emu(int(box["width"])),
                                Emu(int(box["height"])), data)
    chart = gf.chart
    chart.font.size = Pt(font_size)
    chart.font.name = "微软雅黑"
    # 单系列时不显示图例（legend 为 None，直接设 position 会崩）
    chart.has_legend = len(spec["series"]) > 1 or spec["kind"] == "pie"
    if chart.has_legend and chart.legend is not None:
        chart.legend.position = (XL_LEGEND_POSITION.RIGHT if spec["kind"] == "pie"
                                 else XL_LEGEND_POSITION.BOTTOM)
        chart.legend.include_in_layout = False
    plot = chart.plots[0]
    plot.has_data_labels = True
    if spec["kind"] == "pie":
        plot.data_labels.number_format = '0.0"%"'
        plot.data_labels.number_format_is_linked = False
        plot.data_labels.position = XL_LABEL_POSITION.OUTSIDE_END
    else:
        plot.data_labels.font.size = Pt(10)
        if spec["kind"] == "column":
            plot.gap_width = 60
    return chart


def _fill_chart_slide(slide, item, info, font, warnings):
    content_ph = _ph(slide, info.idx("body"))
    table_block = next((b for b in item.blocks if b.block_type == "table"), None)
    if content_ph is None or table_block is None:
        _fill_content_slide(slide, item, "content", info, font)
        return
    headers = list(table_block.meta.get("headers", []))
    rows = [list(r) for r in table_block.meta.get("rows", [])]
    if len(headers) > TABLE_MAX_COLS:
        warnings.append(f"第{item.page_no}页: 表格 {len(headers)} 列超上限，截断为 "
                        f"{TABLE_MAX_COLS} 列")
        headers = headers[:TABLE_MAX_COLS]
        rows = [r[:TABLE_MAX_COLS] for r in rows]
    if len(rows) + 1 > TABLE_MAX_ROWS:
        warnings.append(f"第{item.page_no}页: 表格 {len(rows) + 1} 行超上限，截断为 "
                        f"{TABLE_MAX_ROWS} 行")
        rows = rows[:TABLE_MAX_ROWS - 1]
    # 引导语：放在表格上方按正文排版（不进脚注，避免关键信息降级为小字）
    # 同页含正文块时（小表并入），正文直接作为表格上方的文字
    title_key = item.title.rstrip("（续0123456789）").strip() or item.title
    text_blocks = [b for b in item.blocks
                   if b.block_type not in ("table", "image") and b.text.strip()
                   and not (b.block_type == "heading"
                            and b.text.strip() in (item.title, title_key))]
    _thesis_intro, source = _split_thesis(item.thesis)
    intro = "\n".join(b.text for b in text_blocks) if text_blocks else _thesis_intro
    # 表格字号随版心宽度自适应（窄版心用 12pt，减少单元格换行与标点悬挂）
    avail_in = content_ph.width / 914400
    cell_pt = 14 if avail_in >= 8.0 else 12
    cap_h = int(0.45 * 914400)
    intro_h = 0
    if intro:
        # 行数按模板实际版心宽度估算（4:3 模板只有 5.6in 宽，不能按 12in 算）
        w_pt = content_ph.width / 914400 * 72 - 12
        cpl = max(1, int(w_pt / INTRO_PT))
        lines = sum(max(1, math.ceil(len(t) / cpl)) for t in intro.split("\n"))
        line_h = int(INTRO_PT * 1.34 / 72 * 914400)  # 实际行高，避免高估留下空带
        intro_h = min(int(3.2 * 914400), lines * line_h) + int(0.1 * 914400)
    # 先算列宽与估算行高，再决定整块位置（行高会因单元格换行而增加）
    widths_in, table_w_in = _col_widths(headers, rows, content_ph.width / 914400, cell_pt)
    table_w = int(table_w_in * 914400)
    est_h = int(_table_height(headers, rows, widths_in, cell_pt) * 914400)
    avail_h = int(content_ph.height) - intro_h - cap_h - int(0.2 * 914400)
    # 目标：整块占版心约 85%；表格不足时拉伸行高填满，避免大片留白
    desired = int(0.85 * content_ph.height) - intro_h - cap_h - int(0.08 * 914400)
    table_h = min(avail_h, max(est_h, desired))
    if table_h < int(1.5 * 914400):
        # 表格过矮时加大字号一档，改善投影可读性
        cell_pt = min(18, cell_pt + 2)
        widths_in, table_w_in = _col_widths(headers, rows, content_ph.width / 914400, cell_pt)
        table_w = int(table_w_in * 914400)
        est_h = int(_table_height(headers, rows, widths_in, cell_pt) * 914400)
        table_h = min(avail_h, max(est_h, desired))
    table_h = max(table_h, int(TABLE_ROW_IN * 914400))
    block_h = intro_h + table_h + int(0.08 * 914400) + cap_h  # noqa: F841 说明性合计
    # 表格页顶对齐：表格已按需拉伸，居中反而在标题下留出空带
    gap = 0
    top = Emu(int(content_ph.top + gap))
    if intro:
        box = slide.shapes.add_textbox(Emu(int(content_ph.left)), Emu(int(top)),
                                       Emu(int(content_ph.width)), Emu(intro_h))
        box.name = "doc2ppt_intro"  # 固定名，便于修复环定位该段的标点问题
        box.line.fill.background()  # 显式无边框：否则继承主题默认线，可能压出横线
        tf = box.text_frame
        tf.word_wrap = True
        # 内容优先取本页正文块（保留各自字号覆盖），否则用 thesis 里的引导语
        segs = [(b.text, b.meta.get("size") or INTRO_PT) for b in text_blocks] or \
               [(t, INTRO_PT) for t in intro.split("\n") if t.strip()]
        first = True
        for seg, seg_pt in segs:
            para = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            run = para.add_run()
            _set_run_text(para, run, seg, seg_pt, False, C_BODY, font)
        top = Emu(int(top + intro_h + int(0.12 * 914400)))  # 与表格留出呼吸区，避免边缘叠压
    # 数据页且可数值化 → 出原生图表或表格；记录内容块左缘/宽度，供注脚随块对齐
    block_left, block_w = int(content_ph.left), int(content_ph.width)  # 图表占满版心
    chart_spec = _chart_spec(headers, rows) if len(intro) <= 150 else None
    if chart_spec is not None:
        chart_h = max(int(2.6 * 914400),
                      min(avail_h, int(0.85 * content_ph.height) - intro_h))
        content_ph._element.getparent().remove(content_ph._element)
        _draw_chart(slide, chart_spec,
                    {"left": int(content_ph.left), "top": int(top),
                     "width": int(content_ph.width), "height": chart_h},
                    cell_pt)
        table_h = chart_h
    else:
        # 移除内容占位符，原位放原生表格（避免幽灵空框压在表格下）
        content_ph._element.getparent().remove(content_ph._element)
        n_rows, n_cols = len(rows) + 1, len(headers)
        left = content_ph.left + Emu(int((content_ph.width - table_w) / 2))
        block_left, block_w = int(left), int(table_w)  # 表格居中时注脚随块对齐
        gf = slide.shapes.add_table(n_rows, n_cols, Emu(int(left)), Emu(int(top)),
                                    Emu(table_w), Emu(table_h))
        table = gf.table
        for j in range(n_cols):
            table.columns[j].width = Emu(int(widths_in[j] * 914400))
            cell = table.cell(0, j)
            _set_cell(cell, headers[j], font, bold=True, size=cell_pt)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor.from_string(C_HEAD)
        for i, row in enumerate(rows):
            zebra = i % 2 == 1  # 隔行浅底，长表格页更易横向读数
            for j in range(min(len(row), n_cols)):
                cell = table.cell(i + 1, j)
                _set_cell(cell, row[j], font, size=cell_pt)
                if zebra:
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = RGBColor.from_string(C_ZEBRA)
        # graphic frame 的高度不会拉伸行；逐行等比放大到目标高度，消除表格页底部空白带
        total_h = sum(r.height for r in table.rows) or 1
        for r_ in table.rows:
            r_.height = Emu(max(int(r_.height * table_h / total_h), 0.18 * 914400))
    cap = _ph(slide, info.idx("caption"))
    if cap is None:
        # 模板没有脚注槽：来源注写进演讲者备注，不阻塞出片
        warnings.append(f"第{item.page_no}页: 模板无脚注槽，数据来源写入备注")
        return
    table_bottom = int(top) + table_h
    # 注脚与内容块（表格/图表）左缘对齐：表格居中时注脚不再孤悬最左；四边全设防继承丢失
    cap.left = Emu(int(block_left))
    cap.width = Emu(int(block_w))
    cap.top = Emu(table_bottom + int(0.08 * 914400))  # 来源注紧随内容块，不留大段空白
    cap.height = Emu(int(0.45 * 914400))
    _set_para(cap, source or "数据来源：原文表格", 12, C_CAPT, font=font)


def _split_thesis(thesis):
    """表格页 thesis = 引导语 + "\\n" + 数据来源；拆开分别排版。"""
    if not thesis:
        return "", ""
    parts = [p.strip() for p in thesis.split("\n") if p.strip()]
    if not parts:
        return "", ""
    if parts[-1].startswith("数据来源"):
        return " ".join(parts[:-1]), parts[-1]
    return " ".join(parts), ""


def _same_title(heading, title):
    """判断小标题是否与页标题重复：忽略编号前缀（一、/1.1/（一））与空白差异。"""
    def norm(s):
        return re.sub(r"^[\s0-9]+([.．][0-9]+)*[、.．]?\s*|^[一二三四五六七八九十〇]+、\s*|^（[一二三四五六七八九十]+）\s*",
                      "", (s or "").strip())

    a = norm(heading)
    b = norm(re.sub(r"（续\d*）$", "", title or ""))
    return bool(a) and (a == b or a.startswith(b) or b.startswith(a))


def _fill_notes(slide, item):
    notes = item.notes or f"{item.title}。{item.thesis}"
    img_blocks = [b for b in item.blocks if b.block_type == "image"]
    if img_blocks:
        total = sum(b.meta.get("count", 1) for b in img_blocks)
        placed = sum(1 for sh in slide.shapes
                     if getattr(sh, "shape_type", None) == MSO_SHAPE_TYPE.PICTURE)
        if placed:
            rest = f"，另有 {total - placed} 张未嵌入" if total > placed else ""
            notes += f"（已嵌入原文档图片 {placed} 张{rest}）"
        else:
            notes += f"（原文档此处含 {total} 张图片，版面空间不足未嵌入，如需请人工补充）"
    slide.notes_slide.notes_text_frame.text = notes


def _col_widths(headers, rows, available, pt=14):
    """列宽按内容估算分配（CJK 1em / ASCII 0.55em）。

    总宽 = 自然宽夹在「可用宽 55% ~ 100%」之间：内容少时不成一团，内容多时
    等比压缩且单列不低于 0.7in。返回 (宽度列表, 表格总宽)，单位英寸。
    """
    n_cols = len(headers)
    need = []
    char_in = pt / 72  # 单字宽度（英寸）
    for j in range(n_cols):
        cells = [headers[j]] + [r[j] if j < len(r) else "" for r in rows]
        widest = max((sum(1.0 if ord(ch) >= 0x2E80 else 0.55 for ch in str(c)) for c in cells),
                     default=4.0)
        need.append(max(widest * char_in + 0.4, 1.1))
    natural = sum(need)
    total = min(max(natural, available * 0.55), available)
    k = total / natural
    widths = [w * k for w in need]
    if min(widths) < 0.7:  # 极窄列兜底：抬高到 0.7in 后重新等比压缩
        widths = [max(w, 0.7) for w in widths]
        k = total / sum(widths)
        widths = [w * k for w in widths]
    return widths, sum(widths)


def _set_run_text(para, first_run, text, size, bold, color, font):
    """写段落文本：\n → 段级 <a:br/>（官方 add_line_break，br 是 a:p 的子元素而非 run 的子元素）。

    后续行写入独立 run 并应用相同样式。
    """
    parts = text.split("\n")
    first_run.text = parts[0]
    _style_run(first_run, size, bold, color, font)
    for seg in parts[1:]:
        para.add_line_break()
        r = para.add_run()
        r.text = seg
        _style_run(r, size, bold, color, font)


def _center(para):
    """标题段落居中（图表版式继承默认左对齐，跨页需统一）。"""
    pPr = para._p.get_or_add_pPr()
    pPr.set("algn", "ctr")


def _set_para(ph, text, size, color, font, bold=False, center=False):
    para = ph.text_frame.paragraphs[0]
    para.text = ""
    run = para.add_run()
    _set_run_text(para, run, text, size, bold, color, font)
    if center:
        _center(para)


def _set_cell(cell, text, font, bold=False, size=14):
    cell.text = ""
    run = cell.text_frame.paragraphs[0].add_run()
    run.text = text
    _style_run(run, size, bold, C_TITLE if bold else C_BODY, font)


def _table_height(headers, rows, widths_in, pt):
    """按单元格换行行数估算表格总高（英寸）：行高 = 行数×行距 + 内边距。"""
    def lines_in(text, w_in):
        cpl = max(1.0, (w_in * 72 - 12) / pt)  # CJK 约 1em/字
        return max(1, math.ceil(len(str(text)) / cpl))

    def row_h(cells):
        n = max((lines_in(c, w) for c, w in zip(cells, widths_in)), default=1)
        return n * pt * 1.3 / 72 + 0.08

    total = row_h(headers)
    for r in rows:
        total += row_h(list(r) + [""] * (len(headers) - len(r)))
    return total

PARA_SPLIT_LEN = 60  # 普通段落超过该长度且有 ≥2 句即按句拆行（幻灯片化，逐句原样保留）


def _sentences(text):
    """按句读切分（。！？；），碎句并入前句，避免孤字行。"""
    parts = re.findall(r"[^。！？；]*[。！？；]|[^。！？；]+$", text)
    out = []
    for p in parts:
        p = p.strip()
        if not p:
            continue
        if out and len(p) < 8:
            out[-1] += p
        else:
            out.append(p)
    return out


def display_texts(block):
    """块 → 渲染段落文本列表。长普通段按句拆成多行，让正文页呈现条目感
    而不是大段文字墙（逐句原样保留，不改写不删字）。
    fix._body_blocks 与 verify.estimate_body_need 必须与此函数同口径，
    否则标点修复会打错段落、容量估算会低估。"""
    if block.block_type == "para" and len(block.text) > PARA_SPLIT_LEN:
        parts = _sentences(block.text)
        if len(parts) >= 2:
            return parts
    return [block.text]


def _fill_body(ph, blocks, size, font, title=None):
    """填充正文：与页标题重复的小标题不渲染；小节标题去项目符号、加前间距
    （buNone + spcBef），与列表项形成层级；列表项不加「•」前缀（PowerPoint
    内容占位符自带项目符号，加前缀会出现双重圆点）。"""
    tf = ph.text_frame
    tf.word_wrap = True
    first = True
    for b in blocks:
        if b.block_type == "table":
            continue
        if b.block_type == "image" or not b.text.strip():
            continue  # 图片不嵌入（解析层已记录数量），空文本跳过
        if b.block_type == "heading":
            if title and _same_title(b.text, title):
                continue  # 与页标题重复的小标题不重复渲染（忽略编号差异）
            _text, bold, color = b.text, True, C_TITLE
        elif b.block_type == "bullet":
            _text, bold, color = b.text, False, C_BODY
        else:
            _text, bold, color = b.text, False, C_BODY
        segs = display_texts(b)
        for seg in segs:
            para = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            para.text = ""
            run = para.add_run()
            para_size = b.meta.get("size") or size  # 块级字号覆盖（修复环针对个别段微调）
            _set_run_text(para, run, seg, para_size, bold, color, font)
            if b.block_type != "heading":
                # 顶格：首行与续行同容量（消除悬挂缩进导致的孤标点折行），模拟器方可精确复现
                pPr = para._p.get_or_add_pPr()
                pPr.set("marL", "0")
                pPr.set("indent", "0")
                mar_r = b.meta.get("marR")
                if mar_r:
                    pPr.set("marR", str(int(mar_r)))  # 标点修复的右缩进微调
            if b.block_type == "heading":
                pPr = para._p.get_or_add_pPr()
                if pPr.find(qn("a:buNone")) is None:
                    pPr.append(parse_xml(
                        '<a:buNone xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'
                    ))
                if pPr.find(qn("a:spcBef")) is None:
                    pPr.insert(0, parse_xml(
                        '<a:spcBef xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                        '<a:spcPts val="600"/></a:spcBef>'
                    ))
            elif b.block_type == "para" and len(segs) == 1:
                # 单段普通文字不带项目符号；拆成多条的段落改用与列表一致的圆点
                pPr = para._p.get_or_add_pPr()
                if pPr.find(qn("a:buNone")) is None:
                    pPr.append(parse_xml(
                        '<a:buNone xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'
                    ))


def _style_run(run, size, bold, color, font):
    """run 级样式：latin + ea 双写（仅 latin 时中文字形不生效）。"""
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = font
    run.font.color.rgb = RGBColor.from_string(color)
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = rPr.makeelement(qn("a:ea"), {"typeface": font})
        latin = rPr.find(qn("a:latin"))
        latin.addnext(ea) if latin is not None else rPr.append(ea)
    else:
        ea.set("typeface", font)
