"""冒烟测试：覆盖解析 / 大纲 / 模板探测 / 安全防护 / 渲染闭环。

运行：
    python -m pytest test_smoke.py -q                # 快测（约 10s，不依赖办公软件）
    DOC2PPT_SLOW=1 python -m pytest test_smoke.py -q  # 含端到端（启动 PowerPoint，约 30s）

设计原则：
- 只断言「可判定的行为」，不断言具体像素/字体，避免一处样式调整就全红；
- 每个断言都对应一条 README 决策记录或一次真实回归缺陷（括号内注明）。
"""
import os
import sys
import zipfile
from pathlib import Path

import pytest
from pptx import Presentation

from doc2ppt import fix, guard, outline, parse, render, template, verify

ROOT = Path(__file__).resolve().parent
SAMPLES = ROOT / "samples"
TPL = ROOT / "doc2ppt" / "templates" / "standard_16x9.pptx"
SLOW = pytest.mark.skipif(not os.environ.get("DOC2PPT_SLOW"),
                          reason="端到端用例需 PowerPoint，设置 DOC2PPT_SLOW=1 启用")


# -------------------- 夹具 --------------------

@pytest.fixture
def docx_blocks():
    return parse.parse_document(SAMPLES / "sample_report.docx")


@pytest.fixture
def pdf_blocks():
    return parse.parse_document(SAMPLES / "sample_plan.pdf")


@pytest.fixture
def md_blocks():
    return parse.parse_document(SAMPLES / "sample_notes.md")


def _write_md(tmp_path, name, text):
    p = tmp_path / name
    p.write_text(text, encoding="utf-8")
    return p


# -------------------- 解析层 --------------------

def test_parse_docx_headings_and_table(docx_blocks):
    """标题层级与表格识别（M1 验收项）"""
    heads = [b for b in docx_blocks if b.block_type == "heading"]
    levels = {b.level for b in heads}
    assert {0, 1, 2} <= levels, f"应含文档标题/章/节三级，实际 {levels}"
    tables = [b for b in docx_blocks if b.block_type == "table"]
    assert len(tables) == 1
    assert tables[0].meta["headers"] == ["指标", "目标值", "实际值"]


def test_parse_pdf_text_and_table(pdf_blocks):
    """PDF 文本层与表格（决策 #9/#10：字号相对化 + 表格兜底）"""
    assert sum(1 for b in pdf_blocks if b.block_type == "heading") >= 10
    assert sum(1 for b in pdf_blocks if b.block_type == "table") == 1
    # 段落没有被拆成一堆超短块
    short = [b for b in pdf_blocks if b.block_type == "para" and len(b.text) < 30]
    assert not short, f"PDF 段落切分异常: {[b.text[:10] for b in short]}"


def test_parse_md_code_block_kept(tmp_path):
    """md 代码块不再丢内容（决策 #69）"""
    p = _write_md(tmp_path, "code.md", "# 标题\n\n## 一、示例\n\n```python\nprint(1)\nx = 2\n```\n")
    blocks = parse.parse_document(p)
    texts = "".join(b.text for b in blocks)
    assert "print(1)" in texts and "x = 2" in texts


def test_parse_md_table_escaped_pipe(tmp_path):
    """表格单元格转义竖线（决策 #61）"""
    p = _write_md(tmp_path, "pipe.md",
                  "# T\n\n| 命令 | 说明 |\n| --- | --- |\n| a \\| b | 转义 |\n")
    tbl = next(b for b in parse.parse_document(p) if b.block_type == "table")
    assert tbl.meta["rows"][0][0] == "a | b"


def test_parse_md_nested_list_not_merged(tmp_path):
    """嵌套列表展开为独立条目（决策 #61）"""
    p = _write_md(tmp_path, "nested.md",
                  "# T\n\n## 一、清单\n\n- 一级\n  - 二级 A\n  - 二级 B\n")
    bullets = [b.text for b in parse.parse_document(p) if b.block_type == "bullet"]
    assert bullets == ["一级", "二级 A", "二级 B"], bullets


# -------------------- 大纲层 --------------------

def test_outline_page_range_and_no_none_title(docx_blocks):
    """页数落在验收区间、页标题不为空（决策 #67：None 标题崩溃修复）"""
    items = outline.build_outline(docx_blocks, page_target=8, doc_name="r.docx",
                                  max_chars=575, body_box=(12.0, 5.6))
    assert 6 <= len(items) <= 10, f"页数 {len(items)} 超出 6-10"
    assert all(it.title and it.title.strip() for it in items)


def test_outline_headings_only_makes_toc(tmp_path):
    """纯大纲文档汇总成目录页，不丢标题（决策 #100）"""
    p = _write_md(tmp_path, "toc.md", "# 标题\n\n## 一、第一章\n\n## 二、第二章\n")
    items = outline.build_outline(parse.parse_document(p), page_target=8, doc_name="t.md")
    assert len(items) >= 2
    joined = "".join(b.text for it in items for b in it.blocks)
    assert "一、第一章" in joined and "二、第二章" in joined


def test_outline_long_paragraph_split(tmp_path):
    """超长单段按句读切分，不产生超大块（决策 #70）"""
    long_text = "这是一段用于测试超长段落切分行为的正文内容。" * 60
    p = _write_md(tmp_path, "long.md", f"# T\n\n## 一、章\n\n{long_text}\n")
    items = outline.build_outline(parse.parse_document(p), page_target=8, doc_name="l.md",
                                  max_chars=575, body_box=(12.0, 5.6))
    biggest = max(len(b.text) for it in items for b in it.blocks)
    assert biggest <= 600, f"最大块 {biggest} 字，切分未生效"


def test_outline_empty_doc_only_cover(tmp_path):
    """空文档只出封面且不崩（决策 #68）"""
    p = _write_md(tmp_path, "empty.md", "")
    items = outline.build_outline(parse.parse_document(p), page_target=8, doc_name="e.md")
    assert len(items) == 1


# -------------------- 模板 --------------------

def test_template_probe_own():
    """自家模板探测出四类版式与角色槽位（决策 #80）"""
    _prs, infos = template.load_template(TPL)
    assert set(infos) >= {"title", "content", "two_col", "chart"}
    for sem in ("title", "content", "two_col", "chart"):
        assert "title" in infos[sem].slots, f"{sem} 缺 title 槽"
    assert "body" in infos["content"].slots
    assert "caption" in infos["chart"].slots
    assert template.body_capacity(infos) > 300


def test_template_probe_external_degraded(tmp_path):
    """任意 .pptx 都能加载；缺版式时降级而非报错（决策 #81/#82）"""
    deck = tmp_path / "external.pptx"
    Presentation().save(deck)  # python-pptx 默认 4:3 模板
    _prs, infos = template.load_template(deck)
    assert set(infos) >= {"title", "content", "two_col", "chart"}


# -------------------- 安全 --------------------

def test_guard_rejects_xxe(tmp_path):
    """含 DOCTYPE 的 OOXML 被拒绝（决策 #63）"""
    src = SAMPLES / "sample_report.docx"
    bad = tmp_path / "xxe.docx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(bad, "w") as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                data = b'<!DOCTYPE x [<!ENTITY e SYSTEM "file:///c:/win.ini">]>' + data
            zout.writestr(item, data)
    with pytest.raises(ValueError):
        guard.guard_ooxml(bad)
    # 正常文件不受影响
    guard.guard_ooxml(src)


# -------------------- 渲染闭环（需 PowerPoint COM） --------------------

@SLOW
def test_end_to_end_docx(docx_blocks, tmp_path):
    """docx 端到端：8 页左右、每页可编辑有备注、校验 0 问题"""
    items = outline.build_outline(docx_blocks, page_target=8, doc_name="r.docx",
                                  max_chars=575, body_box=(12.0, 5.6))
    out = tmp_path / "out.pptx"
    _rounds, _log, issues = fix.run(items, TPL, out)
    assert issues == [], f"仍有问题: {[(i.page_no, i.issue_type) for i in issues]}"

    prs = Presentation(str(out))
    assert 6 <= len(prs.slides) <= 12
    for i, slide in enumerate(prs.slides, 1):
        has_content = any(sh.has_text_frame and sh.text_frame.text.strip() for sh in slide.shapes) \
            or any(sh.has_table for sh in slide.shapes)
        assert has_content, f"第 {i} 页没有任何内容"
        assert slide.notes_slide.notes_text_frame.text.strip(), f"第 {i} 页缺演讲者备注"


@SLOW
def test_verify_detects_overflow(tmp_path):
    """溢出检测未退化（决策 #91/#93）"""
    from doc2ppt.models import Block, OutlineItem

    items = [
        OutlineItem(page_no=1, title="自检", thesis="溢出自检"),
        OutlineItem(page_no=2, title="超长",
                    blocks=[Block("para", "这是一段故意写得很长的正文内容，用于验证溢出检测。" * 45)]),
    ]
    out = tmp_path / "ovf.pptx"
    render.render(items, TPL, out)  # 跳过修复环，保留溢出
    issues = verify.check(out)
    overflows = [i for i in issues if i.issue_type == "overflow"]
    assert overflows, "构造的溢出页未被检出"
    assert overflows[0].detail, "问题描述为空"


# -------------------- 讲稿 / 无边框表格 / 图片 / 可复现 --------------------

def test_notes_are_structured(docx_blocks):
    """讲稿为「开场→第一/第二/第三→过渡」结构（B5）"""
    items = outline.build_outline(docx_blocks, page_target=8, doc_name="r.docx",
                                  max_chars=575, body_box=(12.0, 5.6))
    assert items[0].notes.startswith("各位好")
    assert f"共 {len(items)} 页" in items[0].notes
    body_notes = items[1].notes
    assert "这一页讲" in body_notes
    assert "第一，" in body_notes
    assert "接下来看" in body_notes or "谢谢" in body_notes
    assert all(len(it.notes) >= 40 for it in items[1:]), "讲稿过短"


def test_pdf_scanned_ocr(tmp_path):
    """扫描版 PDF（无文本层）：有 ocr 扩展则解析出标题与正文；否则跳过"""
    import pymupdf

    from doc2ppt import ocr

    if not ocr.available():
        pytest.skip("未安装 ocr 扩展（pip install doc2ppt[ocr]）")
    src = pymupdf.open(SAMPLES / "sample_plan.pdf")
    pix = src[0].get_pixmap(dpi=150)
    out = pymupdf.open()
    page = out.new_page(width=src[0].rect.width, height=src[0].rect.height)
    page.insert_image(page.rect, pixmap=pix)
    scanned = tmp_path / "scanned.pdf"
    out.save(str(scanned))

    blocks = parse.parse_document(scanned)
    assert blocks, "扫描件未解析出任何内容"
    assert any(b.block_type == "heading" for b in blocks), "OCR 后标题未被识别"
    joined = "".join(b.text or "" for b in blocks)
    assert "企业数据中台" in joined, "OCR 文本与原文不符"


def test_pdf_borderless_table(tmp_path):
    """无边框表格能被识别（B6：列对齐 + 多空格分列）"""
    import pymupdf

    pdf = tmp_path / "borderless.pdf"
    doc = pymupdf.open()
    page = doc.new_page(width=595, height=842)
    msyh = "C:/Windows/Fonts/msyh.ttc"
    font = pymupdf.Font(fontfile=msyh) if os.path.exists(msyh) else pymupdf.Font("china-s")
    tw = pymupdf.TextWriter(page.rect)
    rows = [("指标", "目标值", "实际值"), ("首响准确率", "85%", "87.2%"),
            ("人工转接率", "25%", "21.6%"), ("平均响应时长", "30秒", "18秒")]
    y = 120
    for a, b_, c in rows:
        tw.append((72, y), a.ljust(14) + b_.ljust(12) + c, font=font, fontsize=11)
        y += 22
    tw.write_text(page)
    doc.save(pdf)
    doc.close()

    blocks = parse.parse_document(pdf)
    tables = [b for b in blocks if b.block_type == "table"]
    assert len(tables) == 1, f"应识别出 1 张无边框表，实际 {len(tables)}"
    assert tables[0].meta["headers"][0] == "指标"
    assert len(tables[0].meta["rows"]) == 3


def test_docx_image_embedded(tmp_path):
    """docx 图片真正嵌入到 PPT（B7），且备注说明已嵌入"""
    import pymupdf
    from docx import Document
    from docx.shared import Inches

    img = tmp_path / "pic.png"
    pix = pymupdf.Pixmap(pymupdf.csRGB, pymupdf.IRect(0, 0, 400, 300), False)
    pix.clear_with(180)
    pix.save(str(img))

    doc = Document()
    doc.add_heading("图片嵌入测试", level=0)
    doc.add_heading("一、配图说明", level=1)
    doc.add_paragraph("本页正文较短，下方应出现原文档中的示意图。")
    doc.add_picture(str(img), width=Inches(3))
    src = tmp_path / "with_image.docx"
    doc.save(src)

    items = outline.build_outline(parse.parse_document(src), page_target=8,
                                  doc_name="img.docx", max_chars=575, body_box=(12.0, 5.6))
    out = tmp_path / "img.pptx"
    render.render(items, TPL, out)
    prs = Presentation(str(out))
    pics = [sh for sl in prs.slides for sh in sl.shapes
            if getattr(sh, "shape_type", None) == 13]
    assert pics, "图片未嵌入"
    assert any("已嵌入原文档图片" in sl.notes_slide.notes_text_frame.text
               for sl in prs.slides), "备注未说明图片已嵌入"


def test_docx_textbox_and_header(tmp_path):
    """文本框内容不丢；正文无标题时用页眉当封面标题（而非文件名）"""
    from docx import Document
    from docx.oxml import parse_xml

    doc = Document()
    doc.add_paragraph("正文第一段，没有用标题样式。")
    holder = doc.add_paragraph("")
    tb = parse_xml(
        '<w:txbxContent xmlns:w="http://schemas.openxmlformats.org/'
        'wordprocessingml/2006/main"><w:p><w:r><w:t>文本框里的关键结论</w:t>'
        "</w:r></w:p></w:txbxContent>"
    )
    holder._p.append(tb)
    doc.sections[0].header.is_linked_to_previous = False
    doc.sections[0].header.paragraphs[0].text = "2026 年度项目总结报告"
    src = tmp_path / "tb.docx"
    doc.save(src)

    blocks = parse.parse_document(src)
    texts = [b.text for b in blocks]
    assert "文本框里的关键结论" in texts, "文本框内容被丢弃"
    assert blocks[0].block_type == "heading" and blocks[0].level == 0
    assert blocks[0].text == "2026 年度项目总结报告", "页眉未作为封面标题"


def test_two_col_split_balances_rendered_height():
    """两栏按渲染高度切分：纯字数切分会把长块全堆在左栏"""
    from doc2ppt.models import Block

    blocks = [Block("bullet", "要点内容示例") for _ in range(6)]
    blocks.append(Block("para", "A" * 120))
    blocks += [Block("bullet", "补充说明") for _ in range(2)]
    left, right = render._split_blocks(blocks, size=20, width=5.5)
    assert left and right
    ln = verify.estimate_body_need(left, 20, 5.5, 0)
    rn = verify.estimate_body_need(right, 20, 5.5, 0)
    assert abs(ln - rn) <= 0.15 * (ln + rn), f"两栏高度失衡: {ln:.0f} vs {rn:.0f}"


def test_render_is_reproducible(tmp_path, docx_blocks):
    """同一输入渲染两次，各 XML 部件内容一致（A3：可复现）"""
    items = outline.build_outline(docx_blocks, page_target=8, doc_name="r.docx",
                                  max_chars=575, body_box=(12.0, 5.6))
    out1, out2 = tmp_path / "a.pptx", tmp_path / "b.pptx"
    render.render(items, TPL, out1)
    render.render(items, TPL, out2)

    def parts(path):
        with zipfile.ZipFile(path) as z:
            return {n: z.read(n) for n in z.namelist()
                    if n.endswith((".xml", ".rels")) and "docProps" not in n}

    p1, p2 = parts(out1), parts(out2)
    assert set(p1) == set(p2), "两次输出的部件集合不同"
    diff = [n for n in p1 if p1[n] != p2[n]]
    assert not diff, f"以下部件内容不一致: {diff[:5]}"


# -------------------- 视觉基线回归 --------------------

BASELINE = Path(os.environ.get("DOC2PPT_BASELINE",
                               str(ROOT.parent / "baseline")))
DIFF_RATIO_LIMIT = 0.005  # 同机渲染差异实测为 0%，留 0.5% 容差；超阈值即视为版式漂移


def _pixel_diff_ratio(a, b, stride=3, tol=30):
    """两张 PNG 的差异像素比例。

    逐通道比较、取最大差：抗锯齿会让三通道各差十几，求和会越过阈值，
    把「同一版式的字体平滑差异」误判成版式漂移（实测误报 2–4%）。
    单通道 >30 仍能捕捉真实位移/字号/内容变化。
    """
    import pymupdf

    pa = pymupdf.open(a)[0].get_pixmap(dpi=96)
    pb = pymupdf.open(b)[0].get_pixmap(dpi=96)
    if (pa.width, pa.height) != (pb.width, pb.height):
        return 1.0
    sa, sb, n = pa.samples, pb.samples, pa.n
    total = diff = 0
    for y in range(0, pa.height, stride):
        row = y * pa.stride
        for x in range(0, pa.width, stride):
            i = row + x * n
            total += 1
            if max(abs(sa[i] - sb[i]), abs(sa[i + 1] - sb[i + 1]),
                   abs(sa[i + 2] - sb[i + 2])) > tol:
                diff += 1
    return diff / max(total, 1)


@SLOW
@pytest.mark.skipif(sys.platform != "win32",
                    reason="视觉基线由 PowerPoint 生成，仅在有 COM 的 Windows 上比对")
@pytest.mark.parametrize("doc_name", ["sample_report.docx", "sample_notes.md", "sample_plan.pdf"])
def test_visual_baseline(tmp_path, doc_name):
    """三份样例的渲染结果与基线图的像素差异率不超阈值（A1：替代主观人工看图）。

    首次运行或版式有意调整后，用 DOC2PPT_UPDATE_BASELINE=1 重建基线。
    """
    blocks = parse.parse_document(SAMPLES / doc_name)
    _prs, infos = template.load_template(TPL)
    spec = infos["content"].slots["body"]
    items = outline.build_outline(blocks, page_target=8, doc_name=doc_name,
                                  max_chars=template.body_capacity(infos),
                                  body_box=(spec.w, spec.h))
    out = tmp_path / "deck.pptx"
    png_dir = tmp_path / "png"
    _rounds, _log, _issues = fix.run(items, TPL, out, png_dir=png_dir)
    pngs = sorted(png_dir.glob("slide*.png"))
    assert pngs, "未导出预览图"

    base_dir = BASELINE / Path(doc_name).stem
    if os.environ.get("DOC2PPT_UPDATE_BASELINE"):
        base_dir.mkdir(parents=True, exist_ok=True)
        for p in base_dir.glob("*.png"):
            p.unlink()
        for p in pngs:
            (base_dir / p.name).write_bytes(p.read_bytes())
        pytest.skip(f"已重建基线: {base_dir}（{len(pngs)} 张）")

    if not base_dir.exists():
        pytest.skip(f"基线不存在，先运行 DOC2PPT_UPDATE_BASELINE=1 生成: {base_dir}")

    bad = []
    for p in pngs:
        base = base_dir / p.name
        if not base.exists():
            bad.append(f"{p.name}: 基线缺失")
            continue
        ratio = _pixel_diff_ratio(base, p)
        if ratio > DIFF_RATIO_LIMIT:
            bad.append(f"{p.name}: 差异 {ratio:.1%} > {DIFF_RATIO_LIMIT:.0%}")
    assert not bad, f"{doc_name} 视觉基线不符：\n" + "\n".join(bad)


# -------------------- LLM 大纲（用 MockTransport，不需要真实 Key） --------------------

def _mock_llm(monkeypatch, content, status=200):
    """把 httpx 传输换成固定响应，模拟 LLM 接口（每个用例独立缓存目录）。"""
    import tempfile

    import httpx

    payload = {"choices": [{"message": {"content": content}}]}
    transport = httpx.MockTransport(
        lambda request: httpx.Response(status, json=payload))
    monkeypatch.setattr(outline, "_TRANSPORT", transport)
    monkeypatch.setenv("DOC2PPT_API_URL", "https://api.example.com/v1/chat/completions")
    monkeypatch.setenv("DOC2PPT_API_KEY", "test-key")
    monkeypatch.setenv("DOC2PPT_CACHE_DIR", tempfile.mkdtemp(prefix="d2p_cache_"))


def _llm_refs(blocks):
    """与 outline._llm_payload 同口径的块编号（跳过图片与文档标题块）。"""
    return [i for i, b in enumerate(blocks)
            if b.block_type != "image"
            and not (b.block_type == "heading" and b.level == 0)]


def test_llm_outline_maps_blocks(monkeypatch, docx_blocks):
    """LLM 返回合法分页时：按块编号映射原文、块不重不漏、标题生效"""
    llm_index = _llm_refs(docx_blocks)
    n = len(llm_index)
    table_pos = next(k for k, i in enumerate(llm_index)
                     if docx_blocks[i].block_type == "table")
    half = n // 2
    page1 = [k for k in range(half) if k != table_pos]
    page3 = [k for k in range(half, n) if k != table_pos]
    content = (
        '{"pages":['
        '{"title":"概览","thesis":"概述项目背景与目标","notes":"本页先交代项目背景与目标。","blocks":' + str(page1) + '},'
        '{"title":"关键指标","thesis":"列出核心指标","blocks":[' + str(table_pos) + ']},'
        '{"title":"进展","thesis":"汇报建设进展","blocks":' + str(page3) + '}'
        ']}'
    )
    _mock_llm(monkeypatch, content)

    items = outline.build_outline(docx_blocks, page_target=8, doc_name="r.docx",
                                  max_chars=575, body_box=(12.0, 5.6), use_llm=True)
    # 模型给的页可能过密，会被按版心容量拆成续页（这是设计行为）
    assert len(items) >= 4, f"至少封面+3页，实际 {len(items)}"
    assert items[1].title.startswith("概览")
    assert any("关键指标" in it.title for it in items)
    assert any(it.thesis == "列出核心指标" for it in items)
    assert "LLM 大纲" in items[0].thesis
    # 续页标题应带序号，避免同名页
    titles = [it.title for it in items]
    assert len(titles) == len(set(titles)), f"页标题重复: {titles}"
    # 内容必须来自原文：页内块与原文块「对象级」完全一致（顺序可因表格独立成页而调整）
    got_ids = [id(b) for it in items[1:] for b in it.blocks]
    want_ids = [id(docx_blocks[i]) for i in _llm_refs(docx_blocks)]
    assert sorted(got_ids) == sorted(want_ids), (
        f"LLM 页未完整引用原文块：got={len(got_ids)} want={len(want_ids)} "
        f"dup={len(got_ids) - len(set(got_ids))} "
        f"missing={len(set(want_ids) - set(got_ids))} extra={len(set(got_ids) - set(want_ids))}")
    assert len(got_ids) == len(set(got_ids)), "有块被重复引用"
    assert all(it.notes for it in items), "讲稿缺失"
    # 混合讲稿：模型给的开场句在最前面，后面接规则补的支撑点
    assert items[1].notes.startswith("本页先交代项目背景与目标。"), items[1].notes[:40]


def test_llm_outline_falls_back_on_bad_json(monkeypatch, docx_blocks):
    """模型返回非 JSON 时回退规则版，并在封面备注说明原因"""
    _mock_llm(monkeypatch, "抱歉，我无法完成这个请求。")
    items = outline.build_outline(docx_blocks, page_target=8, doc_name="r.docx",
                                  max_chars=575, body_box=(12.0, 5.6), use_llm=True)
    assert "LLM 大纲未启用" in items[0].notes
    rule_items = outline.build_outline(docx_blocks, page_target=8, doc_name="r.docx",
                                       max_chars=575, body_box=(12.0, 5.6))
    assert len(items) == len(rule_items), "回退后应与规则版页数一致"


def test_llm_outline_attaches_missing_blocks(monkeypatch, docx_blocks):
    """模型漏引用块时自动补挂，内容不丢（而不是整体回退）"""
    _mock_llm(monkeypatch, '{"pages":[{"title":"只有一块","thesis":"测试","blocks":[0]}]}')
    items = outline.build_outline(docx_blocks, page_target=8, doc_name="r.docx",
                                  max_chars=575, body_box=(12.0, 5.6), use_llm=True)
    assert "LLM 大纲未启用" not in items[0].notes, "漏引用不应导致整体回退"
    got_ids = sorted(id(b) for it in items[1:] for b in it.blocks)
    want_ids = sorted(id(docx_blocks[i]) for i in _llm_refs(docx_blocks))
    assert got_ids == want_ids, "补挂后必须覆盖全部原文块"


def test_llm_retries_without_json_mode(monkeypatch):
    """平台不认识 response_format（返回 400）时自动去掉重试（LLM 大纲第 5 条兼容性）"""
    import httpx

    seen = []

    def handler(request):
        body = request.read().decode()
        seen.append("response_format" in body)
        if "response_format" in body:
            return httpx.Response(400, json={"error": "unknown field"})
        return httpx.Response(200, json={"choices": [{"message": {"content": "{}"}}]})

    monkeypatch.setattr(outline, "_TRANSPORT", httpx.MockTransport(handler))
    out = outline._llm_request("https://api.example.com/v1/chat/completions", "k", "", "p")
    assert out == "{}"
    assert seen == [True, False], "应先带 JSON 模式、400 后去掉重试"


def test_llm_retries_on_server_error(monkeypatch):
    """5xx 会退避重试一次（P1-1：避免网络抖动直接整体回退）"""
    import httpx

    calls = []

    def handler(request):
        calls.append(1)
        if len(calls) == 1:
            return httpx.Response(502, json={"error": "bad gateway"})
        return httpx.Response(200, json={"choices": [{"message": {"content": "{}"}}]})

    monkeypatch.setattr(outline, "_TRANSPORT", httpx.MockTransport(handler))
    monkeypatch.setattr("time.sleep", lambda _s: None)
    out = outline._llm_request("https://api.example.com/v1/chat/completions", "k", "", "p")
    assert out == "{}"
    assert len(calls) == 2, f"应重试一次，实际请求 {len(calls)} 次"


# -------------------- 长文档分批 / token 计量 / 缓存 --------------------

def test_llm_chunks_split_and_cover():
    """超长文档按 12k 字符分批，且分批后块不丢不重（P2-5）"""
    from doc2ppt.models import Block

    blocks = [Block("heading", "超长文档", level=0)]
    for i in range(300):
        blocks.append(Block("heading", f"第{i + 1}节", level=1))
        blocks.append(Block("para", f"第{i + 1}节的正文内容，描述业务要点。" * 2))
    chunks = outline._llm_chunks(blocks)
    assert len(chunks) >= 2, "300 节文档应被分成多批"
    flat = [id(b) for c in chunks for b in c]
    want = [id(b) for b in blocks
            if not (b.block_type == "heading" and b.level == 0)]
    assert sorted(flat) == sorted(want), "分批后块集合必须与原文一致"
    assert len(flat) == len(set(flat)), "分批导致块重复"


def test_llm_usage_is_recorded(monkeypatch):
    """调用后记录 token 用量（P2-7）"""
    import httpx

    def handler(request):
        return httpx.Response(200, json={
            "choices": [{"message": {"content": "{}"}}],
            "usage": {"prompt_tokens": 100, "completion_tokens": 50},
        })

    monkeypatch.setattr(outline, "_TRANSPORT", httpx.MockTransport(handler))
    monkeypatch.setenv("DOC2PPT_PRICE_IN", "0.001")
    monkeypatch.setenv("DOC2PPT_PRICE_OUT", "0.002")
    outline.LLM_USAGE.update({"calls": 0, "prompt_tokens": 0,
                              "completion_tokens": 0, "cache_hits": 0})
    outline._llm_request("https://api.example.com/v1/chat/completions", "k", "", "p")
    assert outline.LLM_USAGE["prompt_tokens"] == 100
    assert outline.LLM_USAGE["completion_tokens"] == 50
    cost = outline.llm_cost_cny()
    assert cost is not None and abs(cost - (100 * 0.001 + 50 * 0.002) / 1000) < 1e-9


def test_llm_cache_hit_avoids_second_call(monkeypatch, docx_blocks, tmp_path):
    """相同文档第二次走缓存，不再请求接口（P2-8）"""
    import httpx

    monkeypatch.setenv("DOC2PPT_CACHE_DIR", str(tmp_path / "cache"))
    monkeypatch.delenv("DOC2PPT_NO_CACHE", raising=False)
    calls = []

    llm_index = _llm_refs(docx_blocks)
    n = len(llm_index)
    table_pos = next(k for k, i in enumerate(llm_index)
                     if docx_blocks[i].block_type == "table")
    page1 = [k for k in range(n) if k != table_pos]
    content = ('{"pages":[{"title":"全部","thesis":"概览","blocks":' + str(page1) + '},'
               '{"title":"数据","thesis":"指标","blocks":[' + str(table_pos) + ']}]}')

    def handler(request):
        calls.append(1)
        return httpx.Response(200, json={
            "choices": [{"message": {"content": content}}],
            "usage": {"prompt_tokens": 10, "completion_tokens": 5},
        })

    monkeypatch.setattr(outline, "_TRANSPORT", httpx.MockTransport(handler))
    monkeypatch.setenv("DOC2PPT_API_URL", "https://api.example.com/v1/chat/completions")
    monkeypatch.setenv("DOC2PPT_API_KEY", "k")

    outline.build_outline(docx_blocks, page_target=8, doc_name="r.docx",
                          max_chars=575, body_box=(12.0, 5.6), use_llm=True)
    assert len(calls) == 1
    assert outline.LLM_USAGE["cache_hits"] == 0

    outline.build_outline(docx_blocks, page_target=8, doc_name="r.docx",
                          max_chars=575, body_box=(12.0, 5.6), use_llm=True)
    assert len(calls) == 1, "第二次不应再请求接口"
    assert outline.LLM_USAGE["cache_hits"] == 1
    assert outline.LLM_USAGE["calls"] == 0


# -------------------- 原生图表 --------------------

def test_chart_spec_rules():
    """图表选型：统一单位才转；饼/柱/折线按数据特征选择（P3-9）"""
    # 混合单位（% / 秒 / 分）不转，避免同一坐标轴误导
    assert render._chart_spec(
        ["指标", "目标值", "实际值"],
        [["准确率", "85%", "87.2%"], ["响应", "30秒", "18秒"], ["满意度", "4.2分", "4.4分"]],
    ) is None
    # 纯文本表不转
    assert render._chart_spec(["项目", "说明"], [["A", "很好"], ["B", "一般"]]) is None
    # 单列占比合计≈100 → 饼图
    pie = render._chart_spec(["渠道", "占比"],
                             [["Web", "45%"], ["App", "35%"], ["电话", "20%"]])
    assert pie and pie["kind"] == "pie" and pie["unit"] == "%"
    # 分类像时间 → 折线
    line = render._chart_spec(["月份", "销量"], [["1月", "120"], ["2月", "135"], ["3月", "98"]])
    assert line and line["kind"] == "line"
    # 其余数值对比 → 柱状
    col = render._chart_spec(["阶段", "周期"], [["搭建", "8周"], ["深化", "8周"], ["推广", "6周"]])
    assert col and col["kind"] == "column" and col["series"][0][1] == [8.0, 8.0, 6.0]


def test_chart_rendered_with_source_values(tmp_path):
    """数值表会渲染成原生图表，且图表数据与原文一致（P3-9）"""
    from doc2ppt.models import Block, OutlineItem

    table = Block("table", meta={
        "headers": ["阶段", "周期"],
        "rows": [["搭建", "8周"], ["深化", "8周"], ["推广", "6周"]],
    })
    items = [
        OutlineItem(page_no=1, title="图表自检", thesis="图表页测试"),
        OutlineItem(page_no=2, title="各阶段周期", thesis="数据来源：自检",
                    blocks=[table]),
    ]
    out = tmp_path / "chart.pptx"
    render.render(items, TPL, out)
    prs = Presentation(str(out))
    charts = [sh for sl in prs.slides for sh in sl.shapes
              if getattr(sh, "has_chart", False)]
    assert charts, "数值表未渲染成图表"
    chart = charts[0].chart
    assert list(chart.plots[0].categories) == ["搭建", "深化", "推广"]
    assert [round(v, 1) for v in chart.series[0].values] == [8.0, 8.0, 6.0]


# -------------------- Web 服务（P3-10） --------------------

def _multipart(fields, files):
    """构造 multipart/form-data 请求体，返回 (body, content_type)。"""
    boundary = b"----doc2pptTEST"
    out = b""
    for k, v in fields.items():
        out += (b"--" + boundary + b"\r\nContent-Disposition: form-data; name=\""
                + k.encode() + b"\"\r\n\r\n" + str(v).encode() + b"\r\n")
    for k, (fn, data) in files.items():
        out += (b"--" + boundary + b"\r\nContent-Disposition: form-data; name=\""
                + k.encode() + b"\"; filename=\"" + fn.encode() + b"\"\r\n"
                + b"Content-Type: application/octet-stream\r\n\r\n" + data + b"\r\n")
    out += b"--" + boundary + b"--\r\n"
    return out, "multipart/form-data; boundary=----doc2pptTEST"


def test_web_multipart_and_safe_name():
    """上传解析不依赖 cgi（3.13 已移除）；文件名去路径、限扩展名"""
    from server import parse_multipart, safe_name

    body, ctype = _multipart({"pages": "12", "llm": "1"},
                             {"file": ("报告 v2.docx", b"PK\x03\x04payload")})
    fields, files = parse_multipart(ctype, body)
    assert fields == {"pages": "12", "llm": "1"}
    assert files["file"][0] == "报告 v2.docx"
    assert files["file"][1] == b"PK\x03\x04payload"

    assert safe_name("../../etc/passwd.docx", {".docx"}) == "passwd.docx"
    assert safe_name(r"C:\tmp\方案.pdf", {".pdf"}) == "方案.pdf"
    assert safe_name("CON.md", {".md"}) == "_CON.md"
    with pytest.raises(ValueError):
        safe_name("evil.exe", {".docx"})


@SLOW
def test_web_convert_over_http(tmp_path, monkeypatch):
    """真实 HTTP：上传 docx → 返回可打开的 pptx（同步转换，串行锁）"""
    import http.client
    import threading

    import server

    monkeypatch.setattr(server, "LOG_PATH", tmp_path / "runs.jsonl")
    src = (SAMPLES / "sample_report.docx").read_bytes()
    httpd = server.make_server("127.0.0.1", 0)
    th = threading.Thread(target=httpd.serve_forever, daemon=True)
    th.start()
    port = httpd.server_address[1]
    try:
        body, ctype = _multipart({"pages": "8"}, {"file": ("sample_report.docx", src)})
        conn = http.client.HTTPConnection("127.0.0.1", port, timeout=300)
        conn.request("POST", "/convert", body=body, headers={"Content-Type": ctype})
        resp = conn.getresponse()
        data = resp.read()
        assert resp.status == 200, data[:300]
        assert "presentationml" in resp.getheader("Content-Type", "")
        assert resp.getheader("X-Output-Name"), "缺少下载文件名响应头"
        out = tmp_path / "web.pptx"
        out.write_bytes(data)
        assert len(Presentation(str(out)).slides) >= 1
        assert (tmp_path / "runs.jsonl").exists(), "未写运行日志"
        conn.close()

        conn = http.client.HTTPConnection("127.0.0.1", port, timeout=30)
        conn.request("GET", "/health")
        health = conn.getresponse()
        assert health.status == 200
        conn.close()

        body, ctype = _multipart({}, {"file": ("bad.exe", b"MZ")})
        conn = http.client.HTTPConnection("127.0.0.1", port, timeout=30)
        conn.request("POST", "/convert", body=body, headers={"Content-Type": ctype})
        bad = conn.getresponse()
        assert bad.status == 400
        assert "不支持的文件类型" in bad.read().decode()
        conn.close()
    finally:
        httpd.shutdown()
        httpd.server_close()
        th.join(5)


def test_web_port_fallback():
    """默认端口被占用时自动顺延，而不是抛 WinError 10013 traceback"""
    import server

    first = server.make_server("127.0.0.1", 0)
    busy = first.server_address[1]
    second, actual, shifted = server.bind_server("127.0.0.1", busy, tries=5)
    try:
        assert shifted and actual != busy, f"未顺延: busy={busy} actual={actual}"
        assert second.server_address[1] == actual
    finally:
        second.server_close()
        first.server_close()


def test_web_startup_preflight(tmp_path, monkeypatch):
    """非法端口/缺模板在启动阶段就报错退出，不抛 traceback"""
    import server

    assert server.main(["--port", "70000"]) == 2
    # 模板路径已改为包内资源，patch template_path 才能模拟「缺模板」
    monkeypatch.setattr(server, "template_path",
                        lambda: tmp_path / "no_such_template.pptx")
    assert server.main(["--port", "8123"]) == 2
