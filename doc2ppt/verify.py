"""M5 校验：读回 pptx 做几何/字号/标点校验，输出问题清单。

溢出判定优先级（真实渲染 > 渲染兜底 > 度量估算）：
  1. PowerPoint COM 渲染每页 PNG → 按形状区域像素扫描文本行底，越出框底即溢出；
  2. LibreOffice 无头渲染成 PDF 兜底检查文本越界；
  3. 字体度量估算（pymupdf 实测字宽 + 实测校准行高系数）。
标点排版（行首悬挂 / 末行孤标点）直接用 PowerPoint 的 TextRange.Lines() 真值判定，
不做折行模拟（实测 PowerPoint 每行 53 字、且不做避头尾，模拟器 ±1 字误差会误判）。
字号合规与空槽位用 python-pptx 读回检查。
"""
import json
import math
import shutil
import subprocess
import tempfile
import time
from pathlib import Path

import pymupdf
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from . import render as render_mod
from .models import Issue
from .template import LAYOUT_MAP

BODY_MIN = 14.0
TITLE_MIN = 24.0
LINE_H = 1.25     # 实测校准：微软雅黑换行渲染行距 ≈1.22em，取 1.25 留余量
PARA_EXTRA_PT = 6.0  # 实测校准（COM 墨迹底反推）：模板段落的段间距是绝对值 ~6pt（spcBef 600），与字号无关；用 em 表示会在小字号时低估
MARGIN_PT = 12.0  # 文本框左右内边距合计
BOTTOM_SLACK = 3.0
CLOSING_PUNCT = "，。、；：！？）」』】》〉”’"
REV_LAYOUT = {v: k for k, v in LAYOUT_MAP.items()}
_FONT_CACHE = {}

# 导出每页 PNG + 转储 PowerPoint 逐段逐行真值（首字/偏移/字数）
_PW_SCRIPT = r"""
param([string]$Pptx, [string]$PngDir, [string]$JsonOut, [int]$PngW = 1280, [int]$PngH = 720)
$ErrorActionPreference = 'Stop'
$slides = @()
try {
  $app = New-Object -ComObject PowerPoint.Application
  try {
    $pres = $app.Presentations.Open($Pptx, -1, 0, 0)
    New-Item -ItemType Directory -Force -Path $PngDir | Out-Null
    $page = 1
    foreach ($slide in $pres.Slides) {
      $shapes = @()
      foreach ($shape in $slide.Shapes) {
        try {
          if ($shape.HasTextFrame -ne -1) { continue }
          $tr = $shape.TextFrame.TextRange
          if ($tr.Text.Trim().Length -eq 0) { continue }
          $paras = @()
          $phIdx = -1
          try { $phIdx = $shape.PlaceholderFormat.Idx } catch { $phIdx = -1 }
          $pc = $tr.Paragraphs().Count
          for ($pi = 1; $pi -le $pc; $pi++) {
            $pr = $tr.Paragraphs($pi)
            $lines = @()
            try {
              $lc = $pr.Lines().Count
              for ($li = 1; $li -le $lc; $li++) {
                $lr = $pr.Lines($li)
                $t = $lr.Text -replace "\r", ""
                if ($t.Trim().Length -eq 0) { continue }
                $lines += [PSCustomObject]@{
                  first = $t.Substring(0, 1)
                  last  = $t.Substring($t.Length - 1)
                  len   = $t.Trim().Length
                  rel   = ($lr.Start - $pr.Start)
                }
              }
            } catch { }
            $paras += [PSCustomObject]@{ lines = $lines }
          }
          $shapes += [PSCustomObject]@{ name = [string]$shape.Name; idx = $phIdx; paras = $paras }
        } catch { }
      }
      $slides += [PSCustomObject]@{ page = $page; shapes = $shapes }
      $png = Join-Path $PngDir ('slide{0:D2}.png' -f $page)
      $slide.Export($png, 'PNG', $PngW, $PngH)
      $page += 1
    }
    $pres.Close()
  } finally {
    $app.Quit()
  }
} catch {
  Write-Error $_.Exception.Message
  exit 1
}
$slides | ConvertTo-Json -Depth 8 | Out-File -FilePath $JsonOut -Encoding utf8
"""


def has_libreoffice():
    return shutil.which("soffice") is not None


def check(path, png_dir=None, log=None, status=None):
    """校验 pptx，返回问题清单。

    log 可选：回传校验通道状态；status 可选：写入 degraded=True 表示
    COM/LibreOffice 均不可用、只跑了估算（此时**标点问题未被检查**，
    不能据此声称「0 问题」）。
    """
    issues = _check_static(path)
    if png_dir:
        out_dir, cleanup = Path(png_dir), False
    else:
        out_dir, cleanup = Path(tempfile.mkdtemp()), True
    try:
        com_data, com_err = _com_check(path, out_dir)
        if com_data is not None:
            lines_by_page, pngs = com_data
            for m in _pixel_overflow(path, pngs):
                issues.append(Issue(m["page"], m["name"], "overflow", "high",
                                    f"渲染实测文本底 {m['max_row_pt']:.0f}pt > 框底 "
                                    f"{m['bottom_pt']:.0f}pt（溢出 {m['spill_pt']:.0f}pt）",
                                    "缩字号 / 拆页 / 精简文本"))
            for m in _line_punct(lines_by_page, _idx_maps(path)):
                issues.append(Issue(m["page"], m["name"], "punct_orphan", "mid",
                                    m["detail"], "硬换行重排"))
            if log is not None:
                log.append("溢出校验: COM 渲染像素实测；标点校验: COM 折行真值")
        elif has_libreoffice():
            issues.extend(_libreoffice_check(path, png_dir=out_dir))
            if log is not None:
                log.append("溢出校验: LibreOffice 渲染像素实测（标点未检查）")
            if status is not None:
                status["degraded"] = True
        else:
            issues.extend(_check_estimate(path))
            if log is not None:
                log.append(f"⚠ 校验降级为估算，标点未检查（COM 不可用: {com_err}；无 LibreOffice）")
            if status is not None:
                status["degraded"] = True
    finally:
        if cleanup:
            shutil.rmtree(out_dir, ignore_errors=True)
    return issues


def _check_static(path):
    """python-pptx 读回：字号合规 + 空槽位。"""
    prs = Presentation(str(path))
    issues = []
    for i, slide in enumerate(prs.slides, 1):
        layout_name = slide.slide_layout.name
        # 数据页（表格或原生图表）的小字注不受正文最小字号约束
        has_graphic = any(sh.has_table or getattr(sh, "has_chart", False)
                          for sh in slide.shapes)
        for sh in slide.shapes:
            if sh.name.startswith("doc2ppt_"):  # 页脚戳/导语等自绘形状，不属于版式内容
                continue
            if sh.has_table or not sh.has_text_frame:
                continue
            if not sh.text_frame.text.strip():
                if _is_content_ph(layout_name, sh):
                    issues.append(Issue(i, sh.name, "empty", "mid",
                                        "内容槽位为空", "补充内容或转人工"))
                continue
            sizes = [r.font.size.pt for p in sh.text_frame.paragraphs for r in p.runs
                     if r.font.size is not None]
            if not sizes:
                continue
            # 表格页的小字注（数据来源）不受正文最小字号约束；标题槽按标题口径检查
            is_caption = has_graphic and max(sizes) <= 12.5
            is_title = _ph_idx(sh) == 0
            if not is_caption:
                if is_title and min(sizes) < TITLE_MIN:
                    issues.append(Issue(i, sh.name, "font_small", "mid",
                                        f"标题字号 {min(sizes):.0f}pt < {TITLE_MIN:.0f}pt",
                                        "标题字号提升"))
                elif not is_title and any(s < BODY_MIN for s in sizes):
                    issues.append(Issue(i, sh.name, "font_small", "mid",
                                        f"正文存在字号 < {BODY_MIN:.0f}pt",
                                        "正文字号提升至 14pt"))
    return issues


def _idx_maps(path):
    """每页的「形状名 → 占位符 idx」映射（COM 的 Idx 不稳定，用 python-pptx 取）。"""
    prs = Presentation(str(path))
    maps = []
    for slide in prs.slides:
        m = {}
        for sh in slide.shapes:
            try:
                m[sh.name] = sh.placeholder_format.idx
            except ValueError:
                continue
        maps.append(m)
    return maps


def _line_punct(lines_by_page, idx_maps=None):
    """用 PowerPoint 折行真值判定标点排版问题。

    行首悬挂：某非首行的首字符是闭合标点。
    末行孤标点：末行 ≤2 字且以闭合标点结尾。
    单字孤行：正文段里某行只剩 1 个字。
    issue.element 携带该段**全部行首偏移**（逗号分隔），供 fix 按真实折行定位。
    """
    out = []
    for si, slide in enumerate(lines_by_page):
        page_map = (idx_maps[si] if idx_maps and si < len(idx_maps) else {})
        for sh in slide["shapes"]:
            if sh.get("idx") == 0:
                continue  # 标题页签不做标点检查（短文本，且与正文排版规则不同）
            for pi, para in enumerate(sh["paras"]):
                lines = para["lines"]
                if len(lines) < 1:
                    continue
                hang = next((k for k in range(1, len(lines))
                             if lines[k]["first"] in CLOSING_PUNCT), None)
                tail = (len(lines) > 1 and lines[-1]["len"] <= 2
                        and lines[-1]["last"] in CLOSING_PUNCT)
                # 单字孤行：正文段里某一行只剩 1 个字（换行断得很丑）
                total = sum(l["len"] for l in lines)
                widow = next((k for k, l in enumerate(lines)
                              if l["len"] == 1 and total > 20), None)
                if hang is None and not tail and widow is None:
                    continue
                rels = ",".join(str(l["rel"]) for l in lines)
                if hang is not None:
                    desc = f"第{hang + 1}行行首悬挂标点「{lines[hang]['first']}」"
                elif tail:
                    desc = f"末行仅 {lines[-1]['len']} 字孤标点「{lines[-1]['last']}」"
                else:
                    desc = f"第{widow + 1}行单字孤行"
                out.append({"page": slide["page"],
                            "name": f"{_shape_key(sh, page_map)}#para{pi}@{rels}",
                            "detail": desc})
    return out


def _check_estimate(path):
    """度量估算兜底（COM 不可用时）：pymupdf 实测字宽 + 实测校准行高。"""
    _name, font_file, _warn = render_mod.resolve_font()
    font = _load_font(font_file)
    prs = Presentation(str(path))
    issues = []
    for i, slide in enumerate(prs.slides, 1):
        for sh in slide.shapes:
            if sh.name.startswith("doc2ppt_"):  # 页脚戳/导语等自绘形状，不属于版式内容
                continue
            if sh.has_table or not sh.has_text_frame or not sh.text_frame.text.strip():
                continue
            sizes = [r.font.size.pt for p in sh.text_frame.paragraphs for r in p.runs
                     if r.font.size is not None]
            if not sizes:
                continue
            w_pt = sh.width / 914400 * 72 - MARGIN_PT
            h_pt = sh.height / 914400 * 72
            if w_pt <= 10 or h_pt <= 10:
                continue
            need = _metrics_need(sh.text_frame, max(sizes), w_pt, font)
            if need > h_pt + 6:
                issues.append(Issue(i, sh.name, "overflow", "high",
                                    f"估算需 {need:.0f}pt > 框高 {h_pt:.0f}pt"
                                    f"（超 {need - h_pt:.0f}pt）",
                                    "缩字号 / 拆页 / 精简文本"))
    return issues


def fits_one_line(text, size, w_in):
    """该文本在给定字号与版心宽度下是否单行放下（用于稀疏页字号抬升的折行预防）。"""
    _name, font_file, _warn = render_mod.resolve_font()
    font = _load_font(font_file)
    w_pt = max(w_in * 72 - MARGIN_PT, 10)
    width = (font.text_length(text, fontsize=size) if font
             else sum(size if ord(ch) >= 0x2E80 else 0.55 * size for ch in text))
    return width <= w_pt


def estimate_body_need(blocks, size, w_in, h_in):
    """给 fix / outline 用的解析式估算：返回所需高度(pt)。

    与渲染同口径：长普通段按 display_texts 拆成多个渲染段（每段独立段距），
    否则会低估高度、导致首页溢出后修复环频繁缩字号。
    """
    _name, font_file, _warn = render_mod.resolve_font()
    font = _load_font(font_file)
    w_pt = max(w_in * 72 - MARGIN_PT, 10)
    lines, paras = 0, 0
    for b in blocks:
        for text in render_mod.display_texts(b):
            if not text:
                continue
            width = (font.text_length(text, fontsize=size) if font
                     else sum(size if ord(ch) >= 0x2E80 else 0.55 * size for ch in text))
            lines += max(1, math.ceil(width / w_pt))
            paras += 1
    return lines * size * LINE_H + max(paras - 1, 0) * PARA_EXTRA_PT


def _load_font(font_file):
    if font_file is None:
        return None
    key = str(font_file)
    if key not in _FONT_CACHE:
        try:
            _FONT_CACHE[key] = pymupdf.Font(fontfile=str(font_file))
        except (ValueError, RuntimeError):
            _FONT_CACHE[key] = None
    return _FONT_CACHE[key]


def _metrics_need(text_frame, size, w_pt, font):
    lines, paras = 0, 0
    for p in text_frame.paragraphs:
        text = "".join(r.text for r in p.runs)
        if not text:
            continue
        width = (font.text_length(text, fontsize=size) if font
                 else sum(size if ord(ch) >= 0x2E80 else 0.55 * size for ch in text))
        lines += max(1, math.ceil(width / max(w_pt, 10)))
        paras += 1
    return lines * size * LINE_H + max(paras - 1, 0) * PARA_EXTRA_PT


def _png_size(path):
    """按幻灯片实际宽高比算导出像素尺寸（正方形像素，保证 pt↔px 单一比例）。"""
    prs = Presentation(str(path))
    sw, sh = prs.slide_width, prs.slide_height
    w = 1280
    h = max(1, round(w * sh / sw))
    return w, h


def _com_check(path, out_dir):
    """COM 渲染每页 PNG + 转储折行真值。返回 ((lines_by_page, pngs), 错误)。"""
    with tempfile.TemporaryDirectory() as td:
        out_dir.mkdir(parents=True, exist_ok=True)
        jout = Path(td) / "lines.json"
        ps1 = Path(td) / "export.ps1"
        ps1.write_text(_PW_SCRIPT, encoding="utf-8")
        pw, ph = _png_size(path)
        proc = None
        try:
            for _attempt in range(3):  # COM 退出是异步的，偶发 Open 失败时重试
                proc = subprocess.run(
                    ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(ps1),
                     "-Pptx", str(Path(path).resolve()), "-PngDir", str(out_dir),
                     "-JsonOut", str(jout), "-PngW", str(pw), "-PngH", str(ph)],
                    capture_output=True, timeout=300, check=False,
                )
                if proc.returncode == 0:
                    break
                time.sleep(5)
        except (OSError, subprocess.TimeoutExpired) as e:
            # Linux/macOS 没有 powershell（或超时）——走 LibreOffice/估算通道
            return None, f"PowerShell 不可用: {type(e).__name__}"
        if proc is None or proc.returncode != 0:
            err = ((proc.stderr or b"") if proc else b"") + ((proc.stdout or b"") if proc else b"")
            return None, err.decode(errors="replace")[:500] or "未知错误"
        pngs = sorted(out_dir.glob("slide*.png"))
        if not pngs:
            return None, "COM 导出无产物"
        data = json.loads(jout.read_text(encoding="utf-8-sig"))
        if isinstance(data, dict):
            data = [data]
        return (data, pngs), None


def _pixel_overflow(path, pngs):
    """按形状区域扫描 PNG 暗像素行：文本行底越过形状底边（留 3pt 余量）即溢出。

    扫描上限 = 框底 + 24pt，且不超过下方相邻重叠形状的顶边，避免探入其他形状的文本。
    """
    prs = Presentation(str(path))
    slide_w_pt = prs.slide_width / 914400 * 72
    out = []
    for page_no, (slide, png) in enumerate(zip(prs.slides, pngs), 1):
        doc = pymupdf.open(png)
        pm = doc[0].get_pixmap(dpi=96)
        scale = pm.width / slide_w_pt  # px per pt（正方形像素，纵横同一比例）
        rects = []
        stamp_lefts = []  # doc2ppt_stamp（页脚戳）的左缘：扫描列需避开，防戳墨迹混入
        for sh in slide.shapes:
            if sh.name == "doc2ppt_stamp":
                stamp_lefts.append(sh.left / 914400 * 72)
                continue  # 页脚戳不参与扫描、不作天花板（其墨迹在框底+24pt 窗内）
            # doc2ppt_intro（表格页导语）保持普通形状：参与扫描并充当上方标题的天花板
            # 表格/图片/图表计入 ceiling（其墨迹会干扰上方文字扫描），但自身不参与扫描
            is_pic = getattr(sh, "shape_type", None) == MSO_SHAPE_TYPE.PICTURE
            is_graphic = sh.has_table or is_pic or getattr(sh, "has_chart", False)
            if not (is_graphic
                    or (sh.has_text_frame and sh.text_frame.text.strip())):
                continue
            left = sh.left / 914400 * 72
            top = sh.top / 914400 * 72
            rects.append({"name": sh.name, "scan": not is_graphic,
                          "left": left, "top": top,
                          "right": left + sh.width / 914400 * 72,
                          "bottom": top + sh.height / 914400 * 72})
        for r in rects:
            if not r["scan"]:
                continue
            ceiling = r["bottom"] + 24
            for o in rects:
                if o is r:
                    continue
                # 水平重叠、且顶边在当前形状内部或下方 → 以它的顶边作为扫描上限
                if (o["top"] > r["top"] + 2 and o["left"] < r["right"] - 2
                        and o["right"] > r["left"] + 2):
                    ceiling = min(ceiling, o["top"] - 2)
            x0 = int((r["left"] + 7) * scale)
            x1 = min(int((r["right"] - 7) * scale), pm.width - 1)
            # 页脚戳墨迹在框底+24pt 的扫描窗内，按 x 列排除其覆盖区间
            for sl_pt in stamp_lefts:
                if r["left"] + 7 < sl_pt - 4 < r["right"] - 7:
                    x1 = min(x1, int((sl_pt - 4) * scale))
            y0 = int(r["top"] * scale)
            y1 = min(int(ceiling * scale), pm.height - 1)
            if x1 <= x0:
                continue
            max_row = _last_ink_row(pm, x0, x1, y0, y1)
            if max_row is not None and max_row / scale > r["bottom"] - BOTTOM_SLACK:
                out.append({"page": page_no, "name": r["name"],
                            "max_row_pt": round(max_row / scale, 1),
                            "bottom_pt": round(r["bottom"] - BOTTOM_SLACK, 1),
                            "spill_pt": round(max_row / scale - (r["bottom"] - BOTTOM_SLACK), 1)})
        doc.close()
    return out


def _last_ink_row(pm, x0, x1, y0, y1, thresh=420, need=2):
    """区域内最后一个含文字墨迹的行（自下而上扫，命中即停）。

    直接用 Pixmap.samples 字节缓冲（逐像素 pm.pixel 调用慢一个数量级），
    且自下而上找到第一行墨迹就返回——绝大多数情况下几行内命中。
    """
    n = pm.n  # 每像素通道数（3 或 4）
    stride = pm.stride
    data = pm.samples
    xs = range(x0, x1, 3)
    for y in range(y1, y0 - 1, -1):
        base = y * stride
        dark = 0
        for x in xs:
            i = base + x * n
            if data[i] + data[i + 1] + data[i + 2] < thresh:
                dark += 1
                if dark >= need:
                    return y
    return None


def _shape_key(sh, page_map=None):
    """问题定位键：占位符优先用 python-pptx 的 idx，退路用 COM 的 idx，再退用形状名。"""
    name = sh["name"]
    if page_map and name in page_map:
        return f"ph{page_map[name]}"
    idx = sh.get("idx", -1)
    if idx is not None and idx >= 0:
        return f"ph{idx}"
    return name


def _ph_idx(sh):
    try:
        return sh.placeholder_format.idx
    except ValueError:
        return None


def _is_content_ph(layout_name, sh):
    """正文槽判定：占位符 idx 1（跨模板通用）。表格页会移除该槽，故不在此列为空。"""
    return _ph_idx(sh) == 1


def _libreoffice_check(path, png_dir=None):
    """LibreOffice 无头渲染兜底：pdf 页光栅化后走与 COM 相同的像素溢出扫描。

    soffice 输出文件名跟随输入名（此前误找 out.pdf，兜底一直静默空转）。
    标点检查依赖 COM 折行真值，此通道不提供（check 会标 degraded）。
    """
    soffice = shutil.which("soffice")
    if soffice is None:
        return []
    issues = []
    with tempfile.TemporaryDirectory() as td:
        proc = subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", td, str(path)],
            capture_output=True, timeout=180, check=False,
        )
        pdf = Path(td) / (Path(path).stem + ".pdf")
        if proc.returncode != 0 or not pdf.exists():
            return issues
        doc = pymupdf.open(pdf)
        target_w, _ph = _png_size(path)
        out_dir = Path(png_dir) if png_dir else None
        if out_dir is not None:
            out_dir.mkdir(parents=True, exist_ok=True)
        pngs = []
        for i, page in enumerate(doc, 1):
            zoom = target_w / page.rect.x1
            pm = page.get_pixmap(matrix=pymupdf.Matrix(zoom, zoom))
            png = (out_dir / f"slide{i:02d}.png") if out_dir is not None \
                else Path(td) / f"slide{i:02d}.png"
            pm.save(png)
            pngs.append(png)
            pr = page.rect
            for b in page.get_text("dict")["blocks"]:
                if b.get("type") != 0:
                    continue
                bb = pymupdf.Rect(b["bbox"])
                if bb.x1 > pr.x1 + 1 or bb.y1 > pr.y1 + 1 or bb.x0 < -1 or bb.y0 < -1:
                    issues.append(Issue(i, "渲染图", "overflow", "high",
                                        f"文本越出页面边界: {tuple(round(v, 1) for v in b['bbox'])}",
                                        "缩字号 / 拆页"))
        doc.close()
        for m in _pixel_overflow(path, pngs):
            issues.append(Issue(m["page"], m["name"], "overflow", "high",
                                f"渲染实测文本底 {m['max_row_pt']:.0f}pt > 框底 "
                                f"{m['bottom_pt']:.0f}pt（溢出 {m['spill_pt']:.0f}pt）",
                                "缩字号 / 拆页 / 精简文本"))
    return issues


def print_report(path, issues):
    """校验报告（打印到终端，不生成报告文件）。"""
    prs = Presentation(str(path))
    print("==== verify 校验报告 ====")
    print(f"总页数: {len(prs.slides)}")
    for i, slide in enumerate(prs.slides, 1):
        print(f"  p{i:<3} [{slide.slide_layout.name}]")
    if issues:
        print(f"问题数: {len(issues)}")
        for it in issues:
            print(f"  p{it.page_no} [{it.issue_type}/{it.severity}] {it.element}: "
                  f"{it.detail} → {it.fix_hint}")
    else:
        print("问题数: 0")
