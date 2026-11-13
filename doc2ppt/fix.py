"""M5 修复：按问题清单在 OutlineItem 层定点修复，重渲染幂等出片。

缩字号 / 拆页 / 精简文本都落在对应页的 item 上（页面级定点）；每次修复后
整份从模板副本重渲染（呼应"不在上次输出上叠加"）。最多 2 轮，issue 数
不下降即提前终止；仍有问题打印人工提示。
"""
from . import render as render_mod
from . import verify as verify_mod
from .models import Block, OutlineItem

MAX_ROUNDS = 5  # 标点重排逐处微调，会轻微改变折行；问题集合不变即提前停止
MIN_SIZE = 14
MIN_TITLE_SIZE = 20
TITLE_SIZE = 32  # 与 render 默认标题字号一致
SIZE_STEP = 2
INTRO_SIZE = 16  # 表格页导语字号（与 render 一致）
MAR_R_STEP = 109728   # 右缩进微调步长（0.12 英寸，EMU）
MAX_MAR_R = 548640    # 右缩进上限（0.6 英寸）
DEFAULT_SIZE = 20
MAX_TITLE_CHARS = 20
CLOSING_PUNCT = "，。、；：！？）」』】》〉”’"
SENTENCE_END = "。！？；"
CLAUSE_END = "，、："


def _renderable(blocks, title=None):
    """该页是否有可渲染内容。

    表格、正文/列表都算；标题块只有「不与页标题重复」时才算内容
    （与页标题相同的标题会被渲染层跳过，不能作为页面存在的理由）。
    """
    for b in blocks:
        if b.block_type == "table":
            return True
        if b.block_type in ("para", "bullet") and b.text.strip():
            return True
        if b.block_type == "heading" and b.text.strip() and b.text.strip() != title:
            return True
    return False


def _splittable(blocks):
    """能否在不产生空页的前提下二分（两半都要有可渲染内容）。"""
    if len(blocks) < 3:
        return False
    half = _split_index(blocks)
    return _renderable(blocks[:half]) and _renderable(blocks[half:])


def _split_index(blocks):
    """二分位置：优先让两半都有正文，避免把标题单独甩到一页。"""
    half = (len(blocks) + 1) // 2
    if blocks[half - 1].block_type == "heading" and half < len(blocks):
        half += 1
    return half


def _sentence_cut(text):
    """在文本中部附近找句读断点，返回切分下标（0 表示找不到合适位置）。"""
    mid = len(text) // 2
    for marks in (SENTENCE_END, CLAUSE_END):
        for i in range(len(text) - 1, 0, -1):
            if text[i] in marks and abs(i - mid) <= max(mid // 2, 40):
                return i + 1
    for i in range(len(text) - 1, 0, -1):
        if text[i] in SENTENCE_END:
            return i + 1
    return 0


def run(items, template_path, out_path, png_dir=None, stats=None):
    """渲染→校验→修复循环。返回 (轮次, 日志, 最终问题清单)。

    png_dir 给定且非空时，最终校验同时导出每页 PNG 预览。
    stats 可选：传入 dict 会写入 first（首轮问题数）、issues_by_round（每轮问题数）。
    """
    from .template import load_template

    _prs, infos = load_template(template_path)
    box = {sem: (info.slots["body"].w, info.slots["body"].h)
           for sem, info in infos.items() if "body" in info.slots}
    box2 = {sem: (info.slots["body2"].w, info.slots["body2"].h)
            for sem, info in infos.items() if "body2" in info.slots}
    log = []
    prev_sig = None
    issues = []
    rnd = 0
    per_round = []
    status = {}
    _escalate_sparse_pages(items, box.get("content", (12.0, 5.6)))
    for rnd in range(1, MAX_ROUNDS + 1):
        render_mod.render(items, template_path, out_path, warnings=log)
        issues = verify_mod.check(out_path, log=log, status=status)
        per_round.append(len(issues))
        if not issues:
            log.append(f"第{rnd}轮: 校验通过，无问题")
            break
        sig = {(it.page_no, it.element, it.issue_type) for it in issues}
        if prev_sig is not None and sig == prev_sig:
            log.append(f"第{rnd}轮: 问题集合不变（{len(issues)} 项），提前终止")
            break
        prev_sig = sig
        applied = _apply(items, issues, box, box2)
        if not applied:
            log.append(f"第{rnd}轮: {len(issues)} 项问题无法自动修复，转人工")
            break
        log.append(f"第{rnd}轮: {len(issues)} 项问题 → 修复 {len(applied)} 项"
                   f"（{'；'.join(applied[:6])}{'…' if len(applied) > 6 else ''}）")
    else:
        # 跑满轮次且最后一轮有修复：磁盘落后于最终 items，补一次渲染
        render_mod.render(items, template_path, out_path, warnings=log)
        issues = verify_mod.check(out_path, log=log, status=status)
    if png_dir:
        issues = verify_mod.check(out_path, png_dir=png_dir, status=status)
    # 清掉修复过程中可能产生的空页（只有标题、无正文/表格），避免成品出现空白页
    dropped = _drop_empty_pages(items)
    if dropped:
        log.append(f"清理空页 {dropped} 页")
        render_mod.render(items, template_path, out_path, warnings=log)
        issues = verify_mod.check(out_path, png_dir=png_dir, status=status)
    if status.get("degraded"):
        log.append("⚠ 本次校验降级为估算（未检查标点），结论不可作为验收依据")
    if issues:
        log.append(f"最终仍有 {len(issues)} 项问题，需人工处理")
    if stats is not None:
        stats["first"] = per_round[0] if per_round else 0
        stats["issues_by_round"] = per_round
        stats["degraded"] = bool(status.get("degraded"))
    return rnd, log, issues


def _drop_empty_pages(items):
    """移除没有可渲染内容的页（保留封面），返回删除数量。"""
    keep, dropped = [], 0
    for idx, it in enumerate(items):
        if idx == 0 or _renderable(it.blocks, it.title):
            keep.append(it)
        else:
            dropped += 1
    if dropped:
        items[:] = keep
        for i, it in enumerate(items, 1):
            it.page_no = i
    return dropped


def _apply(items, issues, box=None, box2=None):
    """按页分组应用修复，返回已执行动作描述列表。

    box/box2：各语义版式正文槽的实际尺寸（英寸），用于按模板真实版心估算字号。
    """
    applied = []
    by_page = {}
    for it in issues:
        by_page.setdefault(it.page_no, []).append(it)
    # 从后往前处理：拆页会插入新页导致后续页下标位移，倒序可避免修复错对象
    for page_no, its in sorted(by_page.items(), reverse=True):
        if page_no < 1 or page_no > len(items):
            continue
        item = items[page_no - 1]
        kinds = {it.issue_type for it in its}
        if "empty" in kinds and not item.blocks:
            item.blocks = [Block("para", "（本页内容待补充，请人工完善）")]
            applied.append(f"p{page_no} 空页填充兜底文本")
            continue
        if "overflow" in kinds:
            # 缩字号会重排整页折行：当轮只缩字号，标点问题下轮在最终字号上处理
            if page_no == 1:
                item.thesis = item.thesis[:MAX_TITLE_CHARS] + "…"
                applied.append(f"p{page_no} 封面副标题精简")
                continue
            if any(it.element and (it.element.startswith("Title") or "#title" in it.element)
                   for it in its):
                cur_t = item.title_size or TITLE_SIZE
                if cur_t - SIZE_STEP >= MIN_TITLE_SIZE:
                    item.title_size = cur_t - SIZE_STEP
                    applied.append(f"p{page_no} 标题字号 {cur_t} → {cur_t - SIZE_STEP}pt")
                else:
                    item.title = item.title[:MAX_TITLE_CHARS] + (
                        "…" if len(item.title) > MAX_TITLE_CHARS else "")
                    applied.append(f"p{page_no} 标题精简")
                continue
            cur = item.font_size or DEFAULT_SIZE
            target = _needed_size(item, box, box2)
            if target is not None:
                if target < cur:
                    item.font_size = target
                    applied.append(f"p{page_no} 缩字号 {cur} → {target}pt")
                    continue
                if cur > MIN_SIZE:
                    # 估算认为放得下但实测溢出：降一档，下轮实测复核
                    item.font_size = cur - SIZE_STEP
                    applied.append(f"p{page_no} 实测溢出但估算通过，缩字号 {cur} → {cur - SIZE_STEP}pt")
                    continue
                # target == cur == MIN_SIZE：缩到下限仍溢出 → 同轮降级拆页/精简
            if len(item.blocks) > 2 and _splittable(item.blocks):
                half = _split_index(item.blocks)
                items.insert(page_no, OutlineItem(
                    page_no=page_no + 1, title=f"{item.title}（续）",
                    blocks=item.blocks[half:], notes=item.notes,
                    font_size=item.font_size,
                ))
                item.blocks = item.blocks[:half]
                applied.append(f"p{page_no} 拆页为 2 页")
            elif all(b.block_type == "table" for b in item.blocks):
                # 表格页溢出 = 脚注过长：裁剪 thesis（引导语）
                item.thesis = item.thesis[:60] + "…"
                applied.append(f"p{page_no} 脚注精简")
            else:
                # 块数少但内容超长：按句读二分该块拆成两页（不截断、不丢字）
                longest = max(item.blocks, key=lambda x: len(x.text))
                cut = _sentence_cut(longest.text)
                if cut > 0:
                    tail = Block(longest.block_type, longest.text[cut:],
                                 level=longest.level, meta=dict(longest.meta))
                    longest.text = longest.text[:cut]
                    items.insert(page_no, OutlineItem(
                        page_no=page_no + 1, title=f"{item.title}（续）",
                        blocks=[tail], notes=item.notes, font_size=item.font_size,
                    ))
                    applied.append(f"p{page_no} 超长段落按句读拆分")
                else:
                    applied.append(f"p{page_no} 无法自动拆分，需人工处理")
            continue
        if "punct_orphan" in kinds:
            # 行首悬挂/末行孤标点：按段落号与断点下标批量注入硬换行
            _fix_punct_orphan(item, [it for it in its if it.issue_type == "punct_orphan"],
                              applied, page_no, box, box2)
            continue
        elif "font_small" in kinds:
            item.font_size = MIN_SIZE
            applied.append(f"p{page_no} 正文字号统一为 {MIN_SIZE}pt")
    for i, it in enumerate(items, 1):
        it.page_no = i
    return applied


def _fix_punct_orphan(item, issues, applied, page_no, box=None, box2=None):
    """行首悬挂/末行孤标点：对该段**单独缩小 1pt**，改变折行以避开标点。

    物理约束：上一行若已排满，无法把标点塞回行内，只能改变折行边界；
    缩字号是最小侵入的做法（物化整段会把满行撑出新的单字行）。
    已到最小字号仍悬挂则保留不动（PowerPoint 本身不做避头尾）。
    """
    sem = render_mod.pick_layout(item)
    if sem == "two_col":
        # 必须与渲染层同口径分栏（高度均衡 + 剔除重复标题），否则段落索引错位
        w = (box2 or {}).get("two_col", (box or {}).get("two_col", (5.9, 5.6)))[0]
        left, right = render_mod._split_blocks(
            item.blocks, item.font_size or DEFAULT_SIZE, w, title=item.title)
        cols = {1: left, 2: right}
    else:
        cols = {1: _body_blocks(item), 2: _body_blocks(item)}
    fixed = 0
    for it in issues:
        name, _, rest = it.element.partition("#para")
        para_s = rest.partition("@")[0]
        if "doc2ppt_intro" in name:
            # 表格页导语文本框：整块缩小 1pt（其内容来自该页正文块）
            blocks = [b for b in item.blocks
                      if b.block_type not in ("table", "image") and b.text.strip()
                      and b.block_type != "heading"]
            for b in blocks:
                cur = b.meta.get("size") or INTRO_SIZE
                if cur > MIN_SIZE:
                    b.meta["size"] = cur - 1
                    fixed += 1
            continue
        try:
            ph_idx = int(name[2:]) if name.startswith("ph") else int(name.rsplit(" ", 1)[1])
            para_idx = int(para_s)
        except (ValueError, IndexError):
            continue
        blocks = cols.get(ph_idx)
        if blocks is None or para_idx >= len(blocks):
            continue
        b = blocks[para_idx]
        if b.block_type not in ("para", "bullet"):
            continue
        cur = b.meta.get("size") or item.font_size or DEFAULT_SIZE
        if cur > MIN_SIZE:
            b.meta["size"] = cur - 1
            fixed += 1
        else:
            # 字号已到下限：改用右缩进微调折行（每次 0.12in，上限 0.6in）
            mar_r = int(b.meta.get("marR") or 0)
            if mar_r < MAX_MAR_R:
                b.meta["marR"] = min(mar_r + MAR_R_STEP, MAX_MAR_R)
                fixed += 1
    if fixed:
        applied.append(f"p{page_no} 标点避头尾（该段字号 -1pt）×{fixed}")


def _escalate_sparse_pages(items, body_box):
    """稀疏页字号抬升：内容不足 50% 版心的页逐级加大字号（上限 28pt）。

    三句话的页面无论怎么加段距都填不满 16:9 展示型页面（100pt 段距是设计灾难），
    专业做法是 statement 页：字号抬升 + 居中。字号写在 item.font_size 上，
    渲染/估算/标点修复自然同口径。**优先选「每个条目单行放下」的字号**——
    折行悬挂短行（如「过 100%。」）是视觉验收明确判 fail 的机器痕迹。
    """
    from . import verify as verify_mod

    w, h = body_box or (12.0, 5.6)
    for it in items[1:]:  # 封面不抬
        if any(b.block_type in ("table", "image") for b in it.blocks):
            continue
        content_blocks = [b for b in it.blocks if b.text.strip()]
        if not content_blocks:
            continue
        base = it.font_size or 16
        need = verify_mod.estimate_body_need(content_blocks, base, w, h)
        if need > 0.50 * h * 72:
            continue
        texts = [t for b in content_blocks for t in render_mod.display_texts(b)]
        chosen, fallback = None, None
        for s in (28, 26, 24, 22):  # 自大而小：优先保住「每条目单行」，避免折行悬挂短行
            if s <= base:
                continue
            est = verify_mod.estimate_body_need(content_blocks, s, w, h)
            if est > 0.70 * h * 72:
                continue
            if all(verify_mod.fits_one_line(t, s, w) for t in texts):
                chosen = s
                break
            fallback = fallback or s
        if chosen or fallback:
            it.font_size = chosen or fallback


def _body_blocks(item):
    """与 render._fill_body 同口径的正文块（跳过表格、图片、与页标题重复的小标题）。

    必须复用 render 的标题等价判断（忽略编号），否则段落索引会错位、修复打偏。
    长段落按 render.display_texts 拆成多个渲染段，索引按渲染段对齐
    （同一块重复出现，meta 修改作用于整块，与渲染逐段共用 meta 一致）。
    """
    out = []
    for b in item.blocks:
        if b.block_type in ("table", "image") or not b.text.strip():
            continue
        if b.block_type == "heading" and render_mod._same_title(b.text, item.title):
            continue
        out.extend([b] * len(render_mod.display_texts(b)))
    return out


def _needed_size(item, box=None, box2=None):
    """解析式求不溢出的正文字号（与 verify 同口径的字体度量估算），下限 14pt。

    版心尺寸取自模板实际槽位（box/box2），未提供时回退内置默认模板的尺寸。
    """
    sem = render_mod.pick_layout(item)
    if sem not in ("content", "two_col"):
        return None
    w_in, h_in = (box or {}).get("content", (12.0, 5.6))
    if sem == "two_col":
        w_in, h_in = (box2 or {}).get("two_col", (box or {}).get("two_col", (5.9, 5.6)))
    cur = item.font_size or DEFAULT_SIZE
    while cur > MIN_SIZE:
        left, right = render_mod._split_blocks(item.blocks, cur, w_in,
                                               title=item.title)
        cols = (left + right,) if sem == "content" else (left, right)
        worst = max((verify_mod.estimate_body_need(col, cur, w_in, h_in)
                     for col in cols if col), default=0)
        if worst <= h_in * 72 + 8:
            return cur
        cur -= SIZE_STEP
    return MIN_SIZE
