"""M2 模板解析：读模板 pptx → 语义版式 + 槽位角色。

支持任意 .pptx 模板：先按版式名（中英文）匹配四类语义版式，匹配不到再按
占位符结构打分（标题槽 + 正文槽的数量、位置、面积）。缺失的版式按降级规则
补齐（如无两栏版式则两栏内容回落单栏），保证任何模板都能出片。

四类语义版式与角色：
  title    封面：title(主标题) + subtitle
  content  内容页：title + body
  two_col  两栏页：title + body + body2（无则降级为单栏）
  chart    图表页：title + body(表格区) + caption(数据来源，可缺)
"""
from dataclasses import dataclass, field
from pathlib import Path

from pptx import Presentation

from .guard import guard_ooxml
from .models import SlotSpec

LAYOUT_MAP = {"title": "标题页", "content": "内容页", "two_col": "两栏页", "chart": "图表页"}

# 不参与填充的占位符（日期/页脚/页码）
SKIP_TYPES = {13, 15, 16}
# 图表页 caption 槽位的默认占位符 idx（自家模板约定）
CAPTION_IDX = 2

TITLE_TYPES = {1, 3}      # TITLE / CENTER_TITLE
BODY_TYPES = {2, 7}       # BODY / OBJECT（图片、表格占位符不能填文字，排除）
SUB_TYPE = 4

NAME_RULES = (
    ("title", ("title slide", "标题幻灯片", "封面", "cover")),
    ("two_col", ("two content", "两栏", "双栏", "对比", "comparison")),
    ("chart", ("content with caption", "图表", "caption", "数据页", "表格页")),
    ("content", ("title and content", "标题和内容", "正文", "content")),
)


@dataclass
class LayoutInfo:
    """一个语义版式：真实 layout 对象 + 角色→槽位。"""

    sem: str
    layout: object
    slots: dict = field(default_factory=dict)  # role -> SlotSpec

    def idx(self, role):
        spec = self.slots.get(role)
        return None if spec is None else spec.idx


def load_template(path):
    """读取模板，返回 (Presentation, {sem: LayoutInfo})。"""
    guard_ooxml(path)  # 模板同样是用户输入，防 XXE
    prs = Presentation(str(path))
    sw, sh = prs.slide_width, prs.slide_height
    infos = _probe(prs, sw, sh)
    return prs, infos


# -------------------- 版式探测 --------------------

def _probe(prs, sw, sh):
    layouts = list(prs.slide_layouts)
    scored = {sem: [] for sem in LAYOUT_MAP}
    for layout in layouts:
        feats = _features(layout, sw, sh)
        if not feats["has_title"]:
            continue
        for sem in LAYOUT_MAP:
            s = _score(sem, layout.name, feats)
            if s > 0:
                scored[sem].append((s, layout, feats))
    infos = {}
    used = set()
    for sem in ("title", "content", "two_col", "chart"):
        cand = sorted(scored[sem], key=lambda x: -x[0])
        pick = next((c for c in cand if id(c[1]) not in used), None) or (cand[0] if cand else None)
        if pick is None:
            continue
        _s, layout, feats = pick
        used.add(id(layout))
        infos[sem] = LayoutInfo(sem, layout, _slots_for(sem, layout, feats))
    # 降级补齐：任何模板都要能出片
    content = infos.get("content") or infos.get("title") or next(
        (LayoutInfo("content", l, _slots_for("content", l, _features(l, sw, sh))) for l in layouts),
        None,
    )
    if content is None:
        raise ValueError("模板中没有任何版式可用")
    infos.setdefault("content", content)
    if "title" not in infos:
        infos["title"] = LayoutInfo("title", content.layout,
                                    _slots_for("title", content.layout,
                                               _features(content.layout, sw, sh)))
    if "two_col" not in infos:
        # 无两栏版式：两栏内容回落单栏（渲染层检测 body2 缺失后顺序填充）
        infos["two_col"] = LayoutInfo("two_col", content.layout,
                                      dict(content.slots))
    if "chart" not in infos:
        infos["chart"] = LayoutInfo("chart", content.layout, dict(content.slots))
    return infos


def _features(layout, sw, sh):
    """提炼版式结构特征（相对化，适配任意尺寸/宽高比）。"""
    phs = [p for p in layout.placeholders if p.placeholder_format.type not in SKIP_TYPES]
    title = next((p for p in phs if p.placeholder_format.type in TITLE_TYPES), None)
    sub = next((p for p in phs if p.placeholder_format.type == SUB_TYPE), None)
    bodies = [p for p in phs if p.placeholder_format.type in BODY_TYPES]
    body_area = [(p.width * p.height) / max(sw * sh, 1) for p in bodies]
    big = [p for p, a in zip(bodies, body_area) if a >= 0.25]
    smalls = [p for p, a in zip(bodies, body_area) if a < 0.12]
    return {
        "has_title": title is not None,
        "title": title,
        "sub": sub,
        "bodies": bodies,
        "big": big,
        "smalls": smalls,
        "two_side": _side_by_side(big),
        "n_bodies": len(bodies),
    }


def _side_by_side(bodies):
    """两个正文槽是否左右并排（y 接近且 x 不重叠）。"""
    for i in range(len(bodies)):
        for j in range(i + 1, len(bodies)):
            a, b = bodies[i], bodies[j]
            same_row = abs(a.top - b.top) < 0.2 * max(a.height, b.height)
            disjoint = a.left + a.width <= b.left + 10000 or b.left + b.width <= a.left + 10000
            if same_row and disjoint:
                return (a, b)
    return None


def _score(sem, name, f):
    """版式名优先，其次结构特征。返回 0 表示不适配。"""
    lname = (name or "").lower()
    for label, keys in NAME_RULES:
        if any(k in lname for k in keys):
            if label != sem:
                return 0
            return 10  # 名称命中，直接采用
    if sem == "title":
        if f["sub"] is not None:
            return 6 if f["title"].placeholder_format.type == 3 else 4
        return 0
    if sem == "content":
        return 6 if len(f["big"]) == 1 and f["sub"] is None else 0
    if sem == "two_col":
        return 6 if f["two_side"] else 0
    if sem == "chart":
        if f["big"] and f["smalls"] and f["sub"] is None:
            return 6
        return 0
    return 0


def _slots_for(sem, layout, feats):
    """按语义角色取槽位：title / subtitle / body / body2 / caption。"""
    slots = {}
    title = feats.get("title")
    if title is not None:
        slots["title"] = _spec(sem, title)
    sub = feats.get("sub")
    if sub is not None:
        slots["subtitle"] = _spec(sem, sub)
    bodies = feats.get("bodies") or []
    if sem == "two_col" and feats.get("two_side"):
        a, b = feats["two_side"]
        slots["body"] = _spec(sem, a)
        slots["body2"] = _spec(sem, b)
    elif bodies:
        slots["body"] = _spec(sem, bodies[0])
        if sem == "chart" and len(bodies) > 1:
            slots["caption"] = _spec(sem, bodies[-1])
        if sem == "two_col" and len(bodies) > 1:
            slots["body2"] = _spec(sem, bodies[1])
    return slots


def _spec(sem, ph):
    return SlotSpec(
        layout=sem,
        idx=ph.placeholder_format.idx,
        ph_type=str(ph.placeholder_format.type),
        name=ph.name,
        x=ph.left / 914400,
        y=ph.top / 914400,
        w=ph.width / 914400,
        h=ph.height / 914400,
    )


def body_capacity(infos, size=20.0):
    """按模板正文槽面积估算单页可容纳字符数（CJK 按 1em 计）。

    用于把大纲的每页字数预算收敛到模板真实版心：4:3 模板的版心明显小于
    16:9，沿用固定预算会导致每页都溢出、修复环反复拆页。
    """
    info = infos.get("content") or next(iter(infos.values()))
    spec = info.slots.get("body")
    if spec is None:
        return None
    w_pt = spec.w * 72 - 12          # 扣左右内边距
    h_pt = spec.h * 72 - 6
    chars_per_line = max(1.0, w_pt / size)
    lines = max(1.0, h_pt / (size * 1.25))
    return int(chars_per_line * lines * 0.85)  # 0.85：留出段落间隙与断行余量


def print_slots(infos):
    """调试工具（非管线用）：打印模板探测出的版式与角色槽位。"""
    for sem in LAYOUT_MAP:
        info = infos.get(sem)
        if info is None:
            print(f"[{sem}] 缺失")
            continue
        print(f"[{sem}] {info.layout.name}")
        for role, s in info.slots.items():
            print(f"    {role:<8} idx={s.idx:<3} {s.name:<28} "
                  f"x={s.x:5.2f} y={s.y:5.2f} w={s.w:5.2f} h={s.h:5.2f} ({s.ph_type})")


def template_path():
    """默认模板路径（包内资源；源码仓库与 pip 安装均适用）。"""
    return Path(__file__).resolve().parent / "templates" / "standard_16x9.pptx"
