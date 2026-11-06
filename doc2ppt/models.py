"""数据模型：管线各环节的契约结构。"""
from dataclasses import dataclass, field


@dataclass
class Block:
    """解析层输出的内容块（标题/段落/列表项/表格）。"""

    block_type: str                       # heading | para | bullet | table
    text: str = ""
    level: int | None = None              # heading 层级（0=文档标题），非标题为 None
    seq: int = -1                         # 全文档序号，parse 出口统一回填
    meta: dict = field(default_factory=dict)  # 表格: headers/rows；列表: list_type


@dataclass
class OutlineItem:
    """大纲中的一页：标题 + 唯一论点 + 内容块 + 版式/字号提示。"""

    page_no: int
    title: str
    thesis: str = ""
    layout_hint: str | None = None        # None=渲染时自动选型
    blocks: list = field(default_factory=list)
    font_size: int | None = None          # fix 阶段的页面级字号覆盖（None=默认）
    title_size: int | None = None         # fix 阶段的标题字号覆盖（适配小标题框模板）
    chapter: str = ""                     # 所属章标题（同章统一字号用）
    notes: str = ""


@dataclass
class SlotSpec:
    """模板中的一个槽位（英寸计）。"""

    layout: str                           # 语义版式名: title/content/two_col/chart
    idx: int
    ph_type: str
    name: str
    x: float
    y: float
    w: float
    h: float


@dataclass
class Issue:
    """verify 输出的一条问题。"""

    page_no: int
    element: str
    issue_type: str                       # overflow | font_small | empty
    severity: str                         # high | mid | low
    detail: str
    fix_hint: str = ""
