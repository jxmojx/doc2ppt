"""新版式（指标/流程/目录）渲染冒烟测试：命中语义、生成不崩、对象合法。"""
import tempfile
from pathlib import Path

from pptx import Presentation

from doc2ppt import render as R
from doc2ppt import template as T
from doc2ppt.models import Block, OutlineItem


def _render(items):
    out = Path(tempfile.mkdtemp()) / "layouts.pptx"
    R.render(items, T.template_path(), str(out))
    prs = Presentation(str(out))
    return prs


def test_layout_stats_short_phrases():
    """纯短句（3-5 条、无编号）命中指标页，卡片为装饰形状"""
    items = [
        OutlineItem(1, "封面", thesis="x"),
        OutlineItem(2, "核心指标",
                    blocks=[Block("bullet", "73%"), Block("bullet", "自动解决率"),
                            Block("bullet", "18秒"), Block("bullet", "平均等待"), Block("bullet", "2.1x")]),
    ]
    prs = _render(items)
    sl = prs.slides[1]
    assert sl.slide_layout.name == "指标页", sl.slide_layout.name
    decs = [sh.name for sh in sl.shapes if sh.name.startswith("doc2ppt_dec")]
    assert len(decs) >= 3, "指标卡未生成"


def test_layout_steps_sequence():
    """半数条目带「第X步/阶段」命中流程页，节点圆点为装饰"""
    items = [
        OutlineItem(1, "封面", thesis="x"),
        OutlineItem(2, "实施节奏",
                    blocks=[Block("bullet", "第一步 接入数据"), Block("bullet", "第二步 口径治理"),
                            Block("bullet", "第三步 对外开放")]),
    ]
    prs = _render(items)
    sl = prs.slides[1]
    assert sl.slide_layout.name == "流程页", sl.slide_layout.name
    dots = [sh.name for sh in sl.shapes if sh.name.startswith("doc2ppt_dec")]
    assert dots, "流程节点未生成"


def test_layout_agenda_headings():
    """纯标题条目（≥4 章）命中目录页（渲染层按条目画目录行）"""
    items = [
        OutlineItem(1, "封面", thesis="x"),
        OutlineItem(2, "目录",
                    blocks=[Block("heading", "一、方案背景"), Block("heading", "二、总体架构"),
                            Block("heading", "三、数据治理"), Block("heading", "四、实施路线")]),
    ]
    prs = _render(items)
    sl = prs.slides[1]
    assert any(sh.name.startswith("doc2ppt_dec") for sh in sl.shapes)
    joined = "".join(sh.text_frame.text for sh in sl.shapes if sh.has_text_frame)
    assert "一、方案背景" in joined


def test_layout_fallback_to_content():
    """较长正文段落（每段 >16 字）不误触发新版式，回落内容页"""
    items = [
        OutlineItem(1, "封面", thesis="x"),
        OutlineItem(2, "普通页", blocks=[Block("para", "这是一段较长普通正文，用于自然回退到内容版式。")] * 3),
    ]
    prs = _render(items)
    assert prs.slides[1].slide_layout.name == "内容页"
