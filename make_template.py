"""M0 构建脚本：生成标准 16:9 母版 templates/standard_16x9.pptx

方案：基于 python-pptx 默认模板改造（其占位符是真实占位符，非散装文本框）。
实验结论（详见 README 决策记录）：
  - 版式占位符的位置/尺寸/字体可改，保存读回一致；
  - 由改过的版式新建幻灯片会继承槽位几何；
  - python-pptx 无新建母版 API，也无 LayoutShapes.add_shape，装饰形状用受控 XML 直写。

注册四类语义版式：
  标题页 <- 'Title Slide'(0)           : 主标题 + 副标题
  内容页 <- 'Title and Content'(1)     : 标题 + 内容区
  两栏页 <- 'Two Content'(3)           : 标题 + 左右内容区
  图表页 <- 'Content with Caption'(7)  : 标题 + 图表/表格区 + 数据来源 caption
用法：python make_template.py
"""
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "doc2ppt" / "templates" / "standard_16x9.pptx"

FONT_CN = "思源黑体 CN"
C_DARK = "1F4E79"   # 主色：深蓝（标题/色带）
C_BODY = "333333"   # 正文
C_SUB = "595959"    # 副标题
C_CAPT = "7F7F7F"   # 数据来源 caption

EMU = 914400.0
SLIDE_W, SLIDE_H = 13.333, 7.5  # 英寸，16:9

# (语义名, 默认模板布局序号, 新版式名)
SPECS = [
    ("title", 0, "标题页"),
    ("content", 1, "内容页"),
    ("two_col", 3, "两栏页"),
    ("chart", 7, "图表页"),
]

PP_DATE, PP_FOOTER = 16, 15  # 需移除的占位符类型


def ph_by_idx(layout, idx):
    return next(p for p in layout.placeholders if p.placeholder_format.idx == idx)


def set_geom(ph, x, y, w, h):
    ph.left = Emu(int(x * EMU))
    ph.top = Emu(int(y * EMU))
    ph.width = Emu(int(w * EMU))
    ph.height = Emu(int(h * EMU))


def style_para(paragraph, size, rgb, bold=False):
    """设置段落字体：latin + ea 双写（仅写 a:latin 时中文字形不生效）。"""
    f = paragraph.font
    f.name = FONT_CN
    f.size = Pt(size)
    f.bold = bold
    f.color.rgb = RGBColor.from_string(rgb)
    rPr = paragraph._p.get_or_add_pPr().get_or_add_defRPr()
    latin = rPr.find(qn("a:latin"))
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = rPr.makeelement(qn("a:ea"), {"typeface": FONT_CN})
        if latin is not None:
            latin.addnext(ea)
        else:
            rPr.append(ea)
    else:
        ea.set("typeface", FONT_CN)


def add_rect(layout, x, y, cx, cy, rgb, name):
    """往版式 spTree 直写一个矩形（python-pptx 无 layout.add_shape API）。"""
    sp = parse_xml(
        f'<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        f'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        f'<p:nvSpPr><p:cNvPr id="{100 + len(layout.shapes)}" name="{name}"/>'
        f'<p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{rgb}"/></a:solidFill>'
        f'<a:ln><a:noFill/></a:ln></p:spPr>'
        f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
    )
    sp_tree = layout._element.spTree
    first_sp = sp_tree.find(qn("p:sp"))
    if first_sp is not None:
        first_sp.addprevious(sp)  # 插到最前，避免遮挡占位符
    else:
        sp_tree.append(sp)


def build():
    prs = Presentation()
    prs.slide_width = Emu(int(SLIDE_W * EMU))
    prs.slide_height = Emu(int(SLIDE_H * EMU))
    bar_w = int(0.28 * EMU)

    for sem, idx, name in SPECS:
        layout = prs.slide_layouts[idx]
        layout.name = name
        # 清理：移除日期/页脚占位符，保留页码占位符
        for ph in list(layout.placeholders):
            if ph.placeholder_format.type in (PP_DATE, PP_FOOTER):
                ph._element.getparent().remove(ph._element)
        # 左侧深蓝 accent 色带
        add_rect(layout, 0, 0, bar_w, int(SLIDE_H * EMU), C_DARK, f"AccentBar_{sem}")

        if sem == "title":
            t = ph_by_idx(layout, 0)
            set_geom(t, 0.9, 2.6, 11.6, 1.4)
            style_para(t.text_frame.paragraphs[0], 40, C_DARK, bold=True)
            s = ph_by_idx(layout, 1)
            set_geom(s, 0.9, 4.3, 11.6, 1.0)
            style_para(s.text_frame.paragraphs[0], 20, C_SUB)
        elif sem == "content":
            t = ph_by_idx(layout, 0)
            set_geom(t, 0.8, 0.35, 12.0, 0.9)
            style_para(t.text_frame.paragraphs[0], 32, C_DARK, bold=True)
            b = ph_by_idx(layout, 1)
            set_geom(b, 0.8, 1.5, 12.0, 5.6)
            style_para(b.text_frame.paragraphs[0], 20, C_BODY)
        elif sem == "two_col":
            t = ph_by_idx(layout, 0)
            set_geom(t, 0.8, 0.35, 12.0, 0.9)
            style_para(t.text_frame.paragraphs[0], 32, C_DARK, bold=True)
            left = ph_by_idx(layout, 1)
            set_geom(left, 0.8, 1.5, 5.9, 5.6)
            style_para(left.text_frame.paragraphs[0], 20, C_BODY)
            right = ph_by_idx(layout, 2)
            set_geom(right, 6.85, 1.5, 5.95, 5.6)
            style_para(right.text_frame.paragraphs[0], 20, C_BODY)
        elif sem == "chart":
            t = ph_by_idx(layout, 0)
            set_geom(t, 0.8, 0.35, 12.0, 0.9)
            style_para(t.text_frame.paragraphs[0], 32, C_DARK, bold=True)
            c = ph_by_idx(layout, 1)
            set_geom(c, 0.8, 1.5, 12.0, 4.5)
            cap = ph_by_idx(layout, 2)
            set_geom(cap, 0.8, 6.15, 12.0, 0.9)  # 脚注槽：引导语 2-3 行 + 数据来源
            style_para(cap.text_frame.paragraphs[0], 12, C_CAPT)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    prs.save(OUT)
    return prs


def report():
    prs = Presentation(OUT)
    print(f"模板: {OUT}")
    print(f"尺寸: {prs.slide_width / EMU:.3f} x {prs.slide_height / EMU:.3f} 英寸 (16:9)")
    for sem, _, name in SPECS:
        layout = next(l for l in prs.slide_layouts if l.name == name)
        deco = len(layout.shapes) - len(layout.placeholders)
        print(f"\n[{name}] 槽位清单（装饰形状 {deco} 个）:")
        for ph in layout.placeholders:
            p = ph.text_frame.paragraphs[0]
            try:
                size, color = f"{p.font.size.pt}pt", str(p.font.color.rgb)
            except (AttributeError, TypeError):
                size, color = "-", "-"
            print(f"  idx={ph.placeholder_format.idx:<3} type={ph.placeholder_format.type:<3} "
                  f"{ph.name:<30} x={ph.left / EMU:5.2f} y={ph.top / EMU:5.2f} "
                  f"w={ph.width / EMU:5.2f} h={ph.height / EMU:5.2f} "
                  f"font={p.font.name} {size} #{color}")


if __name__ == "__main__":
    build()
    report()
