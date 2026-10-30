# -*- coding: utf-8 -*-
"""v3 结构化简历 → Word 生成器。

输入统一 JSON 数据模型（editor-v3：resume + theme），输出 A4 .docx。
**导出保真目标**：跟随编辑时选择的版式族（single/headerBar/sidebar/twoCol）、
主色、标题样式（underline/leftbar/bar/plain）、字体字号行距间距与模块显隐
排序，使下载成品与编辑预览保持同一版式（结构一致；Word 流式排版与网页
存在像素级差异属正常）。
"""
import base64
import io
import re
from typing import List, Optional, Tuple

from docx import Document
from docx.document import Document as _DocxDocument
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

PX2PT = 0.75  # 794px ≈ 21cm @96dpi

FONT_NAMES = {
    "yahei": "微软雅黑", "pingfang": "微软雅黑", "songti": "宋体",
    "heiti": "黑体", "kai": "楷体",
}

# 模块元数据：key → (中文标题, 类型, 图标) —— 图标与编辑器渲染一致
MODULE_META = {
    "education":  {"label": "教育背景", "kind": "list",  "icon": "🎓"},
    "internship": {"label": "实习经历", "kind": "list",  "icon": "🏢"},
    "work":       {"label": "工作经验", "kind": "list",  "icon": "💼"},
    "project":    {"label": "项目经验", "kind": "list",  "icon": "🚀"},
    "campus":     {"label": "校园经历", "kind": "list",  "icon": "🏛️"},
    "skill":      {"label": "技能特长", "kind": "skill", "icon": "⚡"},
    "honor":      {"label": "荣誉证书", "kind": "honor", "icon": "🏆"},
    "self":       {"label": "自我评价", "kind": "text",  "icon": "✨"},
    "custom":     {"label": "自定义",   "kind": "text",  "icon": "📝"},
    "hobby":      {"label": "兴趣爱好", "kind": "tags",  "icon": "💡"},
}

CONTACT_ORDER = [
    ("phone", "电话"), ("email", "邮箱"), ("wechat", "微信"), ("city", "现居"),
]
EXTRA_FIELD_LABEL = {
    "gender": "性别", "birth_date": "出生年月", "age": "年龄",
    "hometown": "籍贯", "nation": "民族", "political_status": "政治面貌",
    "marriage": "婚姻状况", "education_degree": "学历",
}
# 左角色块族（bannerLeft）头部的两列信息顺序：先基本信息、再联系方式、最后院校专业
BL_GRID_ORDER = [
    ("birth_date", "出生年月"), ("nation", "民族"), ("hometown", "籍贯"),
    ("political_status", "政治面貌"), ("marriage", "婚姻状况"),
    ("education_degree", "学历"), ("gender", "性别"), ("age", "年龄"),
    ("phone", "电话"), ("email", "邮箱"), ("wechat", "微信"), ("city", "现居"),
    ("school", "毕业院校"), ("major", "专业"),
]

DEFAULT_THEME = {
    "primary": "#34549B", "font": "yahei", "fontSize": 13,
    "lineHeight": 1.72, "gap": 20, "nameSize": 30,
    "family": "single", "sectionStyle": "underline",
}

# sidebar 版式：这些模块进左侧深色栏（与编辑器渲染一致）
SIDE_KEYS = ("skill", "hobby")


def _hex_or(value, fallback: str) -> str:
    v = str(value or "").strip().lstrip("#")
    return v.upper() if re.fullmatch(r"[0-9a-fA-F]{6}", v) else fallback


def _shade(hex_color: str, factor: float) -> str:
    """factor<1 变深（乘系数），factor>1 向白混合。"""
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    if factor <= 1:
        r, g, b = int(r * factor), int(g * factor), int(b * factor)
    else:
        f = min(factor - 1, 1)
        r, g, b = (int(v + (255 - v) * f) for v in (r, g, b))
    return "%02X%02X%02X" % (max(0, min(255, r)), max(0, min(255, g)), max(0, min(255, b)))


def _rgb(hex_color: str) -> RGBColor:
    return RGBColor.from_string(hex_color)


def _set_run_font(run, family: str, size_pt: float, color: str,
                  bold: bool = False, italic: bool = False) -> None:
    run.font.name = family
    run.font.size = Pt(size_pt)
    run.font.color.rgb = _rgb(color)
    run.font.bold = bold
    run.font.italic = italic
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    rfonts.set(qn("w:eastAsia"), family)


def _para_border_bottom(p, color: str, size_eighth_pt: int = 6) -> None:
    ppr = p._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), str(size_eighth_pt))
    bottom.set(qn("w:space"), "2")
    bottom.set(qn("w:color"), color)
    pbdr.append(bottom)
    ppr.append(pbdr)


def _para_border_left(p, color: str) -> None:
    ppr = p._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), "18")
    left.set(qn("w:space"), "4")
    left.set(qn("w:color"), color)
    pbdr.append(left)
    ppr.append(pbdr)
    pf = p.paragraph_format
    pf.left_indent = Cm(0.15)


def _para_shading(p, fill: str) -> None:
    ppr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    ppr.append(shd)


def _tbl_cell_margin(tbl, bottom_pt: float = 0.0, top_pt: float = 0.0) -> None:
    """给（嵌套）表格设置单元格上下内边距（dxa）。

    Word 会忽略单元格末段的 space_before/行距，只有 cell margin 能稳定撑出底部留白。
    """
    tblpr = tbl._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for side, val in (("top", top_pt), ("left", 0.0), ("bottom", bottom_pt), ("right", 0.0)):
        m = OxmlElement(f"w:{side}")
        m.set(qn("w:w"), str(int(round(val * 20))))   # 1pt = 20 dxa
        m.set(qn("w:type"), "dxa")
        mar.append(m)
    look = tblpr.find(qn("w:tblLook"))
    if look is not None:
        look.addprevious(mar)
    else:
        tblpr.append(mar)


def _table_bleed(tbl, total_cm: float, left_margin_cm: float = 1.5,
                 zero_cell_margin: bool = False) -> None:
    """表格左侧负缩进顶到页面左缘，总宽 = 指定宽度（实现横幅/侧栏全幅贴边）。"""
    tbl.autofit = False
    tblpr = tbl._tbl.tblPr
    ind = OxmlElement("w:tblInd")
    ind.set(qn("w:w"), str(-int(round(left_margin_cm * 567))))
    ind.set(qn("w:type"), "dxa")
    tblpr.append(ind)
    tw = tblpr.find(qn("w:tblW"))
    if tw is None:
        tw = OxmlElement("w:tblW")
        tblpr.append(tw)
    tw.set(qn("w:w"), str(int(round(total_cm * 567))))
    tw.set(qn("w:type"), "dxa")
    if zero_cell_margin:
        # 清零单元格默认左右边距 0.19cm，让横幅文字缩进按版心精确对齐。
        # OOXML 顺序要求：w:tblCellMar（含 w:left/w:right 子元素）须位于 w:tblLook 之前。
        mar = OxmlElement("w:tblCellMar")
        for side in ("w:left", "w:right"):
            m = OxmlElement(side)
            m.set(qn("w:w"), "0")
            m.set(qn("w:type"), "dxa")
            mar.append(m)
        look = tblpr.find(qn("w:tblLook"))
        if look is not None:
            look.addprevious(mar)
        else:
            tblpr.append(mar)


def _cell_shading(cell, fill: str) -> None:
    tcpr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcpr.append(shd)


def _spacing(p, before: float = 0, after: float = 0, line: Optional[float] = None) -> None:
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    if line:
        pf.line_spacing = line


def _has(v) -> bool:
    if v is None:
        return False
    if isinstance(v, (list, tuple)):
        return len(v) > 0
    return str(v).strip() != ""


def _s(v) -> str:
    return str(v).strip() if v is not None else ""


def _range_text(x: dict) -> str:
    start = _s(x.get("start")) or _s(x.get("begin")) or _s(x.get("period"))
    if x.get("current"):
        end = "至今"
    else:
        end = _s(x.get("end"))
    if start and end:
        return f"{start} ~ {end}"
    return start or end


def _public_http_host(host: str) -> bool:
    """仅允许公网 host：拒绝 localhost / 内网 / 链路本地 / 保留地址（防 SSRF）。"""
    import ipaddress
    import socket

    host = (host or "").strip().lower()
    if not host or host in ("localhost", "localhost.localdomain"):
        return False
    candidates = []
    try:
        # 域名解析出的所有地址都必须合法，避免 DNS 重绑定绕过
        for info in socket.getaddrinfo(host, None):
            candidates.append(info[4][0])
    except Exception:  # noqa: BLE001
        return False
    if not candidates:
        return False
    for raw in candidates:
        try:
            ip = ipaddress.ip_address(raw.split("%")[0])
        except ValueError:
            return False
        if (ip.is_private or ip.is_loopback or ip.is_link_local or ip.is_multicast
                or ip.is_reserved or ip.is_unspecified):
            return False
    return True


MAX_PHOTO_BYTES = 3 * 1024 * 1024


def _photo_bytes(photo) -> Optional[bytes]:
    """dataURL 优先；http(s) 仅允许公网地址并限 3MB（失败返回 None）。"""
    if not photo:
        return None
    photo = str(photo).strip()
    m = re.match(r"^data:image/(?:png|jpe?g|webp);base64,(.+)$", photo, re.S)
    if m:
        try:
            data = base64.b64decode(m.group(1))
            return data if len(data) <= MAX_PHOTO_BYTES else None
        except Exception:  # noqa: BLE001
            return None
    if photo.startswith(("http://", "https://")):
        from urllib.parse import urlparse
        try:
            parsed = urlparse(photo)
            if not _public_http_host(parsed.hostname or ""):
                return None
            import requests
            with requests.get(photo, timeout=3, stream=True) as resp:
                if not resp.ok:
                    return None
                ctype = (resp.headers.get("Content-Type") or "").lower()
                if ctype and not ctype.startswith("image/"):
                    return None
                buf = bytearray()
                for chunk in resp.iter_content(64 * 1024):
                    buf.extend(chunk)
                    if len(buf) > MAX_PHOTO_BYTES:
                        return None
                return bytes(buf) or None
        except Exception:  # noqa: BLE001
            return None
    return None


def _norm_items(raw) -> List[dict]:
    if not isinstance(raw, list):
        return []
    return [x for x in raw if isinstance(x, dict)]


def _content_lines(x: dict) -> List[str]:
    c = x.get("content")
    if not c:
        return []
    return [ln.strip() for ln in str(c).split("\n") if ln.strip()]


BULLET_MARK_RE = re.compile(r"^(?:\d{1,2}\s*[.、．）)]|[①-⑳]|[●◆✦►▸•·])")


def _bullet(p_line, text: str, family: str, size_pt: float, color: str) -> None:
    p_line.text = ""
    # 内容已自带序号/符号标记（1. 1） ① ● ◆ ✦ ► 等）时不再叠加圆点
    prefix = "" if BULLET_MARK_RE.match(text) else "• "
    run = p_line.add_run(prefix + text)
    _set_run_font(run, family, size_pt, color)
    pf = p_line.paragraph_format
    pf.left_indent = Cm(0.35)
    pf.first_line_indent = Cm(-0.35)


class _Target:
    """统一 doc / table cell 的段落来源。

    单元格（含文本框）自带一个空段落，首次取段落时复用它，避免内容前多出一个空行；
    Document 不复用——那里的空段落是调用方特意留的占位（如头部下方的分隔线段），
    复用会导致后续段落把它吃掉、边框叠加。
    """

    def __init__(self, obj):
        self.obj = obj
        self._first_used = False

    def para(self):
        paras = getattr(self.obj, "paragraphs", None)
        if (paras is not None and not self._first_used and len(paras) == 1
                and not isinstance(self.obj, _DocxDocument)
                and not paras[0].runs and not paras[0].text.strip()):
            self._first_used = True
            return paras[0]
        return self.obj.add_paragraph()

class _Ctx:
    """渲染上下文：颜色 / 字体 / 尺寸 / 数据。"""

    def __init__(self, resume: dict, theme: dict):
        theme = {**DEFAULT_THEME, **(theme or {})}
        self.primary = _hex_or(theme.get("primary"), DEFAULT_THEME["primary"])
        self.deep = _shade(self.primary, 0.72)
        self.soft = _shade(self.primary, 1.88)
        self.body_color = _hex_or(theme.get("textColor"), "333333")
        self.gray = _shade(self.body_color, 1.34) if self.body_color != "333333" else "595959"
        self.on_dark_text = "FFFFFF"
        self.on_dark_muted = "DDE3EE"
        self.family_font = FONT_NAMES.get(theme.get("font") or "yahei", "微软雅黑")
        self.body_pt = float(theme.get("fontSize") or 13) * PX2PT
        self.name_pt = float(theme.get("nameSize") or 30) * PX2PT
        self.line_h = float(theme.get("lineHeight") or 1.72)
        self.gap_pt = float(theme.get("gap") or 20) * PX2PT
        self.family = theme.get("family") if theme.get("family") in ("single", "headerBar", "sidebar", "bannerLeft", "twoCol") else "single"
        self.section_style = theme.get("sectionStyle") if theme.get("sectionStyle") in ("underline", "leftbar", "bar", "solid", "rule", "plain") else "underline"
        self.show_icons = theme.get("showIcons") is not False
        self.page_bg = _hex_or(theme.get("pageBg"), "FFFFFF")
        # ── 像素级垂直节奏（与网页预览对齐）──
        # 网页 line-height = fontSize × 倍数（px）；Word 用固定行距(exact)按 pt 换算，
        # 避免 Word「多倍行距 ×字体默认行高」造成的 ~30% 放大导致页数不一致
        font_px = float(theme.get("fontSize") or 13)
        name_px = float(theme.get("nameSize") or 30)
        pad_px = float(theme.get("padding") or 44)
        lh = self.line_h
        self.line_pt = Pt(font_px * lh * PX2PT)          # 正文固定行距
        self.title_line_pt = Pt((font_px + 2) * lh * PX2PT)
        self.name_line_pt = Pt(name_px * lh * PX2PT)
        self.gap_pt = float(theme.get("gap") or 20) * PX2PT
        self.item_gap_pt = 12 * PX2PT                    # .rr-item margin-bottom
        self.title_after_pt = 10 * PX2PT                 # .rr-sec-title margin-bottom
        self.desc_before_pt = 4 * PX2PT                  # .rr-item-desc margin-top
        self.pad_cm = pad_px * 2.54 / 96.0               # 版心边距（px→cm）
        self.content_cm = 21.0 - self.pad_cm * 2

        basic = resume.get("basic") or {}
        self.basic = basic
        self.modules = resume.get("modules") or {}
        self.order = resume.get("order") or list(MODULE_META.keys())
        self.hidden = set(resume.get("hidden") or [])
        self.name = _s(basic.get("name")) or "姓名"
        self.intention = _s(basic.get("intention"))
        self.contacts = [(label, _s(basic.get(k))) for k, label in CONTACT_ORDER if _has(basic.get(k))]
        self.extras = [(label, _s(basic.get(k))) for k, label in EXTRA_FIELD_LABEL.items() if _has(basic.get(k))]
        self.photo_bytes = _photo_bytes(basic.get("photo"))
        # 头像宽度跟随编辑器的「头像大小」滑块（px → cm），限制在合理范围避免撑破版式
        _ps = float(theme.get("photoSize") or 92)
        self.photo_cm = max(1.6, min(_ps * 2.54 / 96.0, 4.6))

    def item_font(self):
        return self.family_font

    def grid_pairs(self):
        """bannerLeft 头部两列信息：(标签, 值) 列表。"""
        out = []
        for k, label in BL_GRID_ORDER:
            v = _s(self.basic.get(k))
            if v:
                out.append((label, v))
        return out


def _section_title(t: _Target, ctx: _Ctx, label: str, icon: str,
                   width_cm: float, first: bool, on_dark: bool = False) -> None:
    """模块标题：跟随标题样式（underline/leftbar/bar/solid/plain）；深色栏内固定白字。"""
    p = t.para()
    _spacing(p, before=0 if first else ctx.gap_pt, after=ctx.title_after_pt, line=ctx.title_line_pt)
    title_pt = ctx.body_pt + 2
    if not getattr(ctx, "show_icons", True):
        icon = ""
    text = f"{icon} {label}" if icon else label
    if on_dark:
        run = p.add_run(text)
        _set_run_font(run, ctx.item_font(), title_pt, ctx.on_dark_text, bold=True)
        _para_border_bottom(p, "FFFFFF", 8)
        return
    if ctx.section_style == "solid":
        # 实心主色横条 + 白字（版心宽度，对齐「横条式」模板）
        _para_shading(p, ctx.primary)
        pf = p.paragraph_format
        pf.left_indent = Cm(0.12)
        pf.first_line_indent = Cm(0)
        run = p.add_run(text)
        _set_run_font(run, ctx.item_font(), title_pt, "FFFFFF", bold=True)
    elif ctx.section_style == "rule":
        # 标题文字 + 右侧延伸细线：右对齐制表位（下划线引导符）把线拉满版心
        run = p.add_run(text)
        _set_run_font(run, ctx.item_font(), title_pt, ctx.deep, bold=True)
        pf = p.paragraph_format
        try:
            pf.tab_stops.add_tab_stop(Cm(max(width_cm - 0.05, 1.0)),
                                     WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.LINES)
        except Exception:  # noqa: BLE001
            pass
        tab_run = p.add_run("\t")
        _set_run_font(tab_run, ctx.item_font(), title_pt, ctx.primary)
    elif ctx.section_style == "bar":
        _para_shading(p, ctx.soft)
        # 顶部横幅已改成「细顶带 + 白底」，标题色块不再需要贴纸边，与其它版式统一
        p.paragraph_format.left_indent = Cm(0.12)
        run = p.add_run(text)
        _set_run_font(run, ctx.item_font(), title_pt, ctx.deep, bold=True)
    elif ctx.section_style == "leftbar":
        _para_border_left(p, ctx.primary)
        run = p.add_run(text)
        _set_run_font(run, ctx.item_font(), title_pt, ctx.deep, bold=True)
    elif ctx.section_style == "plain":
        run = p.add_run(text)
        _set_run_font(run, ctx.item_font(), title_pt, ctx.deep, bold=True)
    else:  # underline：标题字 deep + 主色下划线（与网页一致）
        run = p.add_run(text)
        _set_run_font(run, ctx.item_font(), title_pt, ctx.deep, bold=True)
        _para_border_bottom(p, ctx.primary, 12)


def _head_block(t: _Target, ctx: _Ctx, width_cm: float, with_extras: bool = True) -> None:
    """姓名 / 求职意向 / 联系方式（浅色头部，用于 single/twoCol/sidebar 主列）。"""
    p = t.para()
    _spacing(p, after=4.5, line=ctx.name_line_pt)
    run = p.add_run(ctx.name)
    _set_run_font(run, ctx.item_font(), ctx.name_pt, ctx.deep, bold=True)
    if ctx.intention:
        p = t.para()
        _spacing(p, after=6, line=ctx.line_pt)
        run = p.add_run("求职意向：" + ctx.intention)
        _set_run_font(run, ctx.item_font(), ctx.body_pt + 1.5, ctx.primary, bold=True)
    if ctx.contacts:
        p = t.para()
        _spacing(p, after=0, line=ctx.line_pt)
        run = p.add_run("  |  ".join(f"{label}：{v}" for label, v in ctx.contacts))
        _set_run_font(run, ctx.item_font(), ctx.body_pt, ctx.gray)
    if with_extras and ctx.extras:
        p = t.para()
        _spacing(p, after=0, line=ctx.line_pt)
        run = p.add_run("  |  ".join(f"{label}：{v}" for label, v in ctx.extras))
        _set_run_font(run, ctx.item_font(), ctx.body_pt, ctx.gray)
    p = t.para()
    _spacing(p, after=10.5)
    _para_border_bottom(p, ctx.primary, 12)


def _basic_grid(t: _Target, ctx: _Ctx, width_cm: Optional[float] = None,
                on_dark: bool = False, after_pt: float = 0.0) -> None:
    """基本信息两列网格（与编辑器 basicGrid 一致）。

    width_cm 为 None 时用整版心宽；传值时用于头部单元格内（有照片时网格并入头部）。
    on_dark=True 用于顶部横幅等主色底上（白字），与网页 `.rr-banner .rr-basic-row` 对齐。
    """
    if not ctx.extras:
        return
    total = width_cm if width_cm else ctx.content_cm
    dark = bool(on_dark)
    label_color = ctx.on_dark_muted if dark else ctx.gray
    value_color = ctx.on_dark_text if dark else ctx.gray
    doc = t.obj
    if not hasattr(doc, "add_table"):
        for label, v in ctx.extras:
            p = t.para()
            _spacing(p, after=0, line=ctx.line_pt)
            run = p.add_run(f"{label}：{v}")
            _set_run_font(run, ctx.item_font(), ctx.body_pt, value_color)
        return
    n = len(ctx.extras)
    tbl = doc.add_table(rows=(n + 1) // 2, cols=2)
    tbl.autofit = False
    # 左右内边距归零：网格文字与姓名/正文严格左对齐；after_pt 用来撑底部留白
    _tbl_cell_margin(tbl, bottom_pt=after_pt)
    for idx, (label, v) in enumerate(ctx.extras):
        cell = tbl.rows[idx // 2].cells[idx % 2]
        cell.width = Cm(round(total / 2, 2))
        p = cell.paragraphs[0]
        _spacing(p, after=0, line=ctx.line_pt)
        r1 = p.add_run(f"{label}：")
        _set_run_font(r1, ctx.item_font(), ctx.body_pt, label_color)
        r2 = p.add_run(v)
        _set_run_font(r2, ctx.item_font(), ctx.body_pt, value_color)
    # 表后 Word 必须有一个段落；python-docx 建表时已自动补了一个空段落，
    # 直接把它压扁复用，避免网格下方多出一整行空白
    paras = getattr(doc, "paragraphs", None) or []
    gap = paras[-1] if paras and not paras[-1].text.strip() else doc.add_paragraph()
    gap.paragraph_format.line_spacing = Pt(1)
    gap.paragraph_format.space_before = Pt(0)
    gap.paragraph_format.space_after = Pt(0)


def _photo_paragraph(t: _Target, ctx: _Ctx, width_cm: float, center: bool = False) -> bool:
    if not ctx.photo_bytes:
        return False
    p = t.para()
    _spacing(p, after=4)
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    try:
        run.add_picture(io.BytesIO(ctx.photo_bytes), width=Cm(width_cm))
        return True
    except Exception:  # noqa: BLE001
        return False


def _render_module(t: _Target, ctx: _Ctx, key: str, raw,
                   width_cm: float, first: bool, on_dark: bool = False) -> None:
    meta = MODULE_META[key]
    label, kind, icon = meta["label"], meta["kind"], meta["icon"]
    fam = ctx.item_font()
    body_pt = ctx.body_pt
    body_color = ctx.on_dark_text if on_dark else ctx.body_color
    muted = ctx.on_dark_muted if on_dark else ctx.gray
    accent = ctx.on_dark_text if on_dark else ctx.primary

    items = _norm_items(raw)
    tags = raw if key == "hobby" and isinstance(raw, list) else []
    text_val = _s(raw) if kind == "text" else ""
    if kind == "list" and not items:
        return
    if kind == "skill" and not items:
        return
    if kind == "honor" and not items:
        return
    if kind == "tags" and not tags:
        return
    if kind == "text" and not text_val:
        return

    _section_title(t, ctx, label, icon, width_cm, first, on_dark=on_dark)

    if kind == "tags":
        p = t.para()
        _spacing(p, after=0, line=ctx.line_pt)
        run = p.add_run("、".join(_s(x) for x in tags if _s(x)))
        _set_run_font(run, fam, body_pt, body_color)
        return

    if kind == "text":
        for ln in [x for x in text_val.split("\n") if x.strip()]:
            p = t.para()
            _spacing(p, after=0, line=ctx.line_pt)
            run = p.add_run(ln.strip())
            _set_run_font(run, fam, body_pt, body_color)
        return

    first_item = True
    for x in items:
        if kind == "list":
            title = _s(x.get("school")) or _s(x.get("company")) or _s(x.get("name"))
            if key == "education":
                sub_parts: Tuple[str, ...] = tuple(v for v in (_s(x.get("degree")), _s(x.get("major"))) if v)
            else:
                role = _s(x.get("position")) or _s(x.get("role"))
                sub_parts = (role,) if role else tuple()
            meta_text = _range_text(x)
            lines = _content_lines(x)
            if not (title or meta_text or sub_parts or lines):
                continue

            p = t.para()
            _spacing(p, before=0 if first_item else ctx.item_gap_pt, after=0, line=ctx.line_pt)
            if sub_text := " · ".join(sub_parts):
                p.paragraph_format.tab_stops.add_tab_stop(Cm(width_cm / 2), WD_TAB_ALIGNMENT.CENTER)
            if meta_text:
                p.paragraph_format.tab_stops.add_tab_stop(Cm(width_cm), WD_TAB_ALIGNMENT.RIGHT)
            if title:
                run = p.add_run(title)
                _set_run_font(run, fam, body_pt + 1, body_color, bold=True)
            if sub_text:
                run = p.add_run("\t" + sub_text)
                _set_run_font(run, fam, body_pt, muted)
            if meta_text:
                run = p.add_run("\t" + meta_text)
                _set_run_font(run, fam, body_pt, muted)
            for li, ln in enumerate(lines):
                p = t.para()
                _spacing(p, before=ctx.desc_before_pt if li == 0 else 0, after=0, line=ctx.line_pt)
                _bullet(p, ln, fam, body_pt, body_color)
            first_item = False
        elif kind == "skill":
            name_v = _s(x.get("name"))
            level = _s(x.get("level"))
            lines = _content_lines(x)
            head = name_v
            if not (head or level or lines):
                continue
            if head or level:
                p = t.para()
                _spacing(p, before=0 if first_item else ctx.item_gap_pt, after=0, line=ctx.line_pt)
                p.paragraph_format.tab_stops.add_tab_stop(Cm(width_cm), WD_TAB_ALIGNMENT.RIGHT)
                run = p.add_run(head)
                _set_run_font(run, fam, body_pt, body_color, bold=True)
                if level:
                    run = p.add_run("\t" + level)
                    _set_run_font(run, fam, body_pt, muted)
            for li, ln in enumerate(lines):
                p = t.para()
                _spacing(p, before=ctx.desc_before_pt if li == 0 else 0, after=0, line=ctx.line_pt)
                _bullet(p, ln, fam, body_pt, body_color)
            first_item = False
        elif kind == "honor":
            name_v = _s(x.get("name"))
            time_v = _s(x.get("time")) or _s(x.get("date"))
            issuer = _s(x.get("issuer"))
            if not (name_v or time_v):
                continue
            p = t.para()
            _spacing(p, before=0 if first_item else ctx.item_gap_pt, after=0, line=ctx.line_pt)
            p.paragraph_format.tab_stops.add_tab_stop(Cm(width_cm), WD_TAB_ALIGNMENT.RIGHT)
            run = p.add_run(name_v + (f"（{issuer}）" if issuer else ""))
            _set_run_font(run, fam, body_pt, body_color)
            if time_v:
                run = p.add_run("\t" + time_v)
                _set_run_font(run, fam, body_pt, muted)
            first_item = False


def _render_sections(t: _Target, ctx: _Ctx, width_cm: float,
                     skip_side: bool = False, on_dark: bool = False) -> None:
    first = True
    for key in ctx.order:
        if key in ctx.hidden or key not in MODULE_META:
            continue
        if skip_side and key in SIDE_KEYS:
            continue
        raw = ctx.modules.get(key)
        _render_module(t, ctx, key, raw, width_cm, first, on_dark=on_dark)
        first = False


def _set_page_bg(doc: Document, hex_color: str) -> None:
    """设置纸张底色：document.xml 写 <w:background>，settings.xml 开启显示。

    白色（默认）不写任何东西，保持与旧版导出完全一致。
    """
    color = (hex_color or "").upper()
    if color in ("", "FFFFFF"):
        return
    try:
        bg = OxmlElement("w:background")
        bg.set(qn("w:color"), color)
        doc.element.insert(0, bg)   # CT_Document 要求 background 在 body 之前
        st = doc.settings.element
        if st.find(qn("w:displayBackgroundShape")) is None:
            el = OxmlElement("w:displayBackgroundShape")
            seq = list(getattr(st, "_tag_seq", ()) or ())
            succ = seq[seq.index("w:displayBackgroundShape") + 1:] if "w:displayBackgroundShape" in seq else []
            st.insert_element_before(el, *succ)
    except Exception:  # noqa: BLE001
        # 底色是装饰能力，失败不影响正文导出
        pass


def _new_doc(ctx: _Ctx) -> Document:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(ctx.pad_cm)
    sec.top_margin = sec.bottom_margin = Cm(ctx.pad_cm)
    if ctx.family in ("headerBar", "sidebar", "bannerLeft"):
        # 网页中横幅/侧栏/角色块以负 margin 顶到纸张边缘（吃掉顶部留白），Word 同样贴顶
        sec.top_margin = Cm(0)
    normal = doc.styles["Normal"]
    normal.font.name = ctx.item_font()
    normal.font.size = Pt(ctx.body_pt)
    normal.font.color.rgb = _rgb(ctx.body_color)
    nrpr = normal.element.get_or_add_rPr()
    nrpr.get_or_add_rFonts().set(qn("w:eastAsia"), ctx.item_font())
    _set_page_bg(doc, ctx.page_bg)
    return doc


def _build_single_headerBar(doc: Document, ctx: _Ctx, banner: bool) -> None:
    """single / headerBar。

    headerBar（顶部横幅）= 页顶一条主色通栏细带 + 白底头部，姓名不再压深色块。
    头部内：姓名 / 意向 / 联系方式 + 基本信息网格（并入以撑高文字区，避免与照片之间留竖条空档），
    右上照片，头部下缘一条主色细线收口。
    """
    if banner:
        pad_pt = ctx.pad_cm / 2.54 * 72

        # ① 页顶主色细带（左右通栏到纸边）
        bar = doc.add_paragraph()
        _spacing(bar, before=0, after=0, line=Pt(6))
        _para_shading(bar, ctx.primary)
        bar.paragraph_format.left_indent = Cm(-ctx.pad_cm)
        bar.paragraph_format.right_indent = Cm(-ctx.pad_cm)

        # ② 头部：左文字 / 右照片（表格贴版心，不出血）
        photo_w = max(3.2, round(ctx.photo_cm + 0.4, 2)) if ctx.photo_bytes else 0.4
        text_w = round(ctx.content_cm - photo_w, 2)
        tbl = doc.add_table(rows=1, cols=2)
        _table_bleed(tbl, ctx.content_cm, left_margin_cm=0.0, zero_cell_margin=True)
        left, right = tbl.rows[0].cells
        left.width, right.width = Cm(text_w), Cm(photo_w)
        left.vertical_alignment = WD_ALIGN_VERTICAL.TOP
        right.vertical_alignment = WD_ALIGN_VERTICAL.TOP

        lt = _Target(left)
        top = lt.para()                  # 顶部留白用空段落（单元格首段设段前距会把行撑到跨页）
        _spacing(top, before=0, after=0, line=Pt(round(pad_pt * 0.62, 1)))
        p = lt.para()
        _spacing(p, before=0, after=2, line=ctx.name_line_pt)
        run = p.add_run(ctx.name)
        _set_run_font(run, ctx.item_font(), ctx.name_pt, ctx.deep, bold=True)
        if ctx.intention:
            p = lt.para()
            _spacing(p, after=2, line=ctx.line_pt)
            run = p.add_run("求职意向：" + ctx.intention)
            _set_run_font(run, ctx.item_font(), ctx.body_pt + 1.5, ctx.primary, bold=True)
        if ctx.contacts:
            p = lt.para()
            _spacing(p, after=0, line=ctx.line_pt)
            p.paragraph_format.right_indent = Cm(0.3)
            run = p.add_run("  |  ".join(f"{label}：{v}" for label, v in ctx.contacts))
            _set_run_font(run, ctx.item_font(), ctx.body_pt, ctx.gray)

        # 基本信息网格并入头部：文字区与照片同高，右侧不再留下竖条空白
        if ctx.extras:
            before = len(left.tables)
            _basic_grid(lt, ctx, text_w)
            if len(left.tables) > before:
                _table_bleed(left.tables[-1], text_w, left_margin_cm=0.0)

        if ctx.photo_bytes:
            rt = _Target(right)
            p = rt.para()
            _spacing(p, before=0, after=0)
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = p.add_run()
            try:
                run.add_picture(io.BytesIO(ctx.photo_bytes), width=Cm(round(photo_w - 0.2, 2)))
            except Exception:  # noqa: BLE001
                pass

        # ③ 头部下缘主色细线（对应网页 .rr-hb-top 的 border-bottom）
        line = doc.add_paragraph()
        _spacing(line, before=6, after=0, line=Pt(1))
        _para_border_bottom(line, ctx.primary, 12)
        sep = doc.add_paragraph()        # 隔开，避免与下一段的下边框被 Word 合并成一条
        _spacing(sep, before=0, after=ctx.gap_pt, line=Pt(1))

        _render_sections(_Target(doc), ctx, ctx.content_cm)
        return

    gap = doc.add_paragraph()  # 无横幅时的头部间隔（1pt 高）
    gap.paragraph_format.line_spacing = Pt(1)
    gap.paragraph_format.space_before = Pt(0)
    gap.paragraph_format.space_after = Pt(0)

    t = _Target(doc)
    _basic_grid(t, ctx)
    _render_sections(t, ctx, ctx.content_cm)


def _build_sidebar(doc: Document, ctx: _Ctx) -> None:
    """sidebar：左侧深色栏（照片/联系方式/基本信息/技能/爱好）+ 右侧主列。"""
    tbl = doc.add_table(rows=1, cols=2)
    _table_bleed(tbl, 19.5)
    left, right = tbl.rows[0].cells
    left.width, right.width = Cm(5.8), Cm(13.7)
    _cell_shading(left, ctx.deep)
    side_w, main_w = 5.0, 13.2

    side = _Target(left)
    _photo_paragraph(side, ctx, min(ctx.photo_cm, 4.0), center=True)
    for p0 in left.paragraphs:
        p0.paragraph_format.left_indent = Cm(0.35)
        p0.paragraph_format.right_indent = Cm(0.15)
    first_side = True
    if ctx.contacts:
        _section_title(side, ctx, "联系方式", "", side_w, first_side, on_dark=True)
        for label, v in ctx.contacts:
            p = side.para()
            _spacing(p, after=0, line=ctx.line_pt)
            run = p.add_run(f"{label}：{v}")
            _set_run_font(run, ctx.item_font(), ctx.body_pt, ctx.on_dark_text)
        first_side = False
    if ctx.extras:
        _section_title(side, ctx, "基本信息", "", side_w, first_side, on_dark=True)
        for label, v in ctx.extras:
            p = side.para()
            _spacing(p, after=0, line=ctx.line_pt)
            run = p.add_run(f"{label}：{v}")
            _set_run_font(run, ctx.item_font(), ctx.body_pt, ctx.on_dark_text)
        first_side = False

    main = _Target(right)
    p = main.para()
    _spacing(p, after=4.5, line=ctx.name_line_pt)
    run = p.add_run(ctx.name)
    _set_run_font(run, ctx.item_font(), ctx.name_pt, ctx.deep, bold=True)
    if ctx.intention:
        p = main.para()
        _spacing(p, after=6, line=ctx.line_pt)
        run = p.add_run("求职意向：" + ctx.intention)
        _set_run_font(run, ctx.item_font(), ctx.body_pt + 1.5, ctx.primary, bold=True)
    p = main.para()
    _spacing(p, after=10.5)
    _para_border_bottom(p, ctx.primary, 12)

    first_main = True
    for key in ctx.order:
        if key in ctx.hidden or key not in MODULE_META:
            continue
        raw = ctx.modules.get(key)
        if key in SIDE_KEYS:
            _render_module(side, ctx, key, raw, side_w, first_side, on_dark=True)
            if _module_has_content(ctx, key, raw):
                first_side = False
        else:
            _render_module(main, ctx, key, raw, main_w, first_main)
            if _module_has_content(ctx, key, raw):
                first_main = False


def _module_has_content(ctx: _Ctx, key: str, raw) -> bool:
    meta = MODULE_META[key]
    kind = meta["kind"]
    if kind == "tags":
        return isinstance(raw, list) and any(_s(x) for x in raw)
    if kind == "text":
        return bool(_s(raw))
    items = _norm_items(raw)
    if not items:
        return False
    if kind == "skill":
        return any(_s(x.get("name")) or _s(x.get("level")) or _content_lines(x) for x in items)
    if kind == "honor":
        return any(_s(x.get("name")) or _s(x.get("time")) or _s(x.get("date")) for x in items)
    return True


def _build_two_col(doc: Document, ctx: _Ctx) -> None:
    """twoCol：头部 + 基本信息 + 双栏流式模块。"""
    t = _Target(doc)
    p = t.para()
    _spacing(p, after=4.5, line=ctx.name_line_pt)
    run = p.add_run(ctx.name)
    _set_run_font(run, ctx.item_font(), ctx.name_pt, ctx.deep, bold=True)
    if ctx.intention:
        p = t.para()
        _spacing(p, after=6, line=ctx.line_pt)
        run = p.add_run("求职意向：" + ctx.intention)
        _set_run_font(run, ctx.item_font(), ctx.body_pt + 1.5, ctx.primary, bold=True)
    if ctx.contacts:
        p = t.para()
        _spacing(p, after=0, line=ctx.line_pt)
        run = p.add_run("  |  ".join(f"{label}：{v}" for label, v in ctx.contacts))
        _set_run_font(run, ctx.item_font(), ctx.body_pt, ctx.gray)
    if ctx.photo_bytes:
        _photo_paragraph(t, ctx, 2.4)
    p = t.para()
    _spacing(p, after=10.5)
    _para_border_bottom(p, ctx.primary, 12)
    _basic_grid(t, ctx)

    # 双栏分节（后续模块内容在两栏间流式排布）
    col_gap = 0.75   # 425 dxa ≈ 0.75cm
    cols = doc.sections[0]._sectPr.xpath("./w:cols")
    if cols:
        cols[0].set(qn("w:num"), "2")
        cols[0].set(qn("w:space"), "425")
    else:
        c = OxmlElement("w:cols")
        c.set(qn("w:num"), "2")
        c.set(qn("w:space"), "425")
        doc.sections[0]._sectPr.append(c)
    _render_sections(t, ctx, round((ctx.content_cm - col_gap) / 2, 2))


def _build_banner_left(doc: Document, ctx: _Ctx) -> None:
    """bannerLeft（左色条）：左侧主色竖条 + 大号姓名（不加底色）+ 右侧两列信息 + 右上照片 + 细线。

    与网页一致：竖条贴纸左缘出血，右侧收在版心（照片不贴纸边）；
    姓名/意向在上、两列信息在下，中间用空段落撑开，避免与照片之间出现整块空白。
    """
    photo_w = max(2.2, round(ctx.photo_cm + 0.1, 2))
    bar_w = 0.32
    total_w = round(21.0 - ctx.pad_cm, 2)
    main_w = round(total_w - bar_w - photo_w, 2)
    ml_ind = max(round(ctx.pad_cm - bar_w, 2), 0.2)   # 竖条已占 0.32cm，姓名左缘对齐正文

    tbl = doc.add_table(rows=1, cols=3)
    _table_bleed(tbl, total_w, left_margin_cm=ctx.pad_cm, zero_cell_margin=True)
    c_bar, c_main, c_photo = tbl.rows[0].cells
    c_bar.width, c_main.width, c_photo.width = Cm(bar_w), Cm(main_w), Cm(photo_w)
    _cell_shading(c_bar, ctx.primary)
    c_main.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    c_photo.vertical_alignment = WD_ALIGN_VERTICAL.TOP

    # 姓名 + 求职意向：深色字，不加底色（老大明确否掉了"名字底下压色块"的做法）
    pad_pt = ctx.pad_cm / 2.54 * 72
    mt = _Target(c_main)
    sp_top = mt.para()                       # 复用单元格自带空段落做顶部留白
    _spacing(sp_top, before=0, after=0, line=Pt(round(pad_pt * 0.52, 1)))
    p = mt.para()
    _spacing(p, before=0, after=2, line=ctx.name_line_pt)
    p.paragraph_format.left_indent = Cm(ml_ind)
    p.paragraph_format.right_indent = Cm(0.3)
    run = p.add_run(ctx.name)
    _set_run_font(run, ctx.item_font(), ctx.name_pt, ctx.body_color, bold=True)
    if ctx.intention:
        p = mt.para()
        _spacing(p, after=9, line=ctx.line_pt)
        p.paragraph_format.left_indent = Cm(ml_ind)
        p.paragraph_format.right_indent = Cm(0.3)
        run = p.add_run(ctx.intention)
        _set_run_font(run, ctx.item_font(), ctx.body_pt, ctx.gray)

    # 两列信息网格（嵌套表格），字号/行距对齐网页的 0.9em / 1.35
    pairs = ctx.grid_pairs()
    if pairs:
        n_rows = (len(pairs) + 1) // 2
        sub = c_main.add_table(rows=n_rows, cols=2)
        sub.autofit = False
        _tbl_cell_margin(sub, bottom_pt=round(pad_pt * 0.3, 1))
        for idx, (label, v) in enumerate(pairs):
            cell = sub.rows[idx // 2].cells[idx % 2]
            cell.width = Cm(round(main_w / 2 - 0.15, 2))
            cp = cell.paragraphs[0]
            _spacing(cp, after=0, line=Pt(ctx.body_pt * 1.35))
            cp.paragraph_format.left_indent = Cm(ml_ind)
            r1 = cp.add_run("◆ ")
            _set_run_font(r1, ctx.item_font(), ctx.body_pt - 1.6, ctx.primary)
            r2 = cp.add_run(f"{label}：")
            _set_run_font(r2, ctx.item_font(), ctx.body_pt - 1.0, ctx.body_color, bold=True)
            r3 = cp.add_run(v)
            _set_run_font(r3, ctx.item_font(), ctx.body_pt - 1.0, ctx.body_color)
        # 信息区下方留白：表后 Word 必补一个空段落，正好拿来做固定高度的空白
        tail = c_main.paragraphs[-1]
        if tail is not sp_top and not tail.text.strip():
            _spacing(tail, before=0, after=0, line=Pt(round(pad_pt * 0.3, 1)))

    # 右上照片：贴版心右缘、无边框
    if ctx.photo_bytes:
        pt = _Target(c_photo)
        p = pt.para()
        _spacing(p, before=0, after=0)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run()
        try:
            run.add_picture(io.BytesIO(ctx.photo_bytes), width=Cm(photo_w))
        except Exception:  # noqa: BLE001
            pass
    else:
        _spacing(_Target(c_photo).para(), before=0, after=0)

    # 细线：左侧通栏到纸边，右侧收在版心（与网页 .rr-bl-hair 一致）
    p = doc.add_paragraph()
    _spacing(p, before=5, after=0, line=Pt(1))
    p.paragraph_format.left_indent = Cm(-ctx.pad_cm)
    _para_border_bottom(p, ctx.primary, 12)

    t = _Target(doc)
    _render_sections(t, ctx, ctx.content_cm)


def generate_docx(resume: dict, theme: dict) -> bytes:
    """统一 JSON 模型 → docx 字节（跟随版式族与标题样式）。"""
    ctx = _Ctx(resume or {}, theme or {})
    doc = _new_doc(ctx)
    if ctx.family == "headerBar":
        _build_single_headerBar(doc, ctx, banner=True)
    elif ctx.family == "bannerLeft":
        _build_banner_left(doc, ctx)
    elif ctx.family == "sidebar":
        _build_sidebar(doc, ctx)
    elif ctx.family == "twoCol":
        _build_two_col(doc, ctx)
    else:
        t = _Target(doc)
        if ctx.photo_bytes:
            tbl = doc.add_table(rows=1, cols=2)
            tbl.autofit = False
            left, right = tbl.rows[0].cells
            _rw = max(2.8, round(ctx.photo_cm + 0.5, 2))
            left.width, right.width = Cm(round(ctx.content_cm - _rw, 2)), Cm(_rw)
            # 文字块与照片垂直居中对齐，避免照片把行高撑高后文字下方留出空档
            left.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            right.vertical_alignment = WD_ALIGN_VERTICAL.TOP
            ht = _Target(left)
            p = ht.para()
            _spacing(p, after=4.5, line=ctx.name_line_pt)
            run = p.add_run(ctx.name)
            _set_run_font(run, ctx.item_font(), ctx.name_pt, ctx.deep, bold=True)
            if ctx.intention:
                p = ht.para()
                _spacing(p, after=6, line=ctx.line_pt)
                run = p.add_run("求职意向：" + ctx.intention)
                _set_run_font(run, ctx.item_font(), ctx.body_pt + 1.5, ctx.primary, bold=True)
            if ctx.contacts:
                p = ht.para()
                _spacing(p, after=0, line=ctx.line_pt)
                run = p.add_run("  |  ".join(f"{label}：{v}" for label, v in ctx.contacts))
                _set_run_font(run, ctx.item_font(), ctx.body_pt, ctx.gray)
            # 扩展字段以两列网格并入头部单元格（与网页一致：网格在头部内，撑高文字区与照片匹配）
            _basic_grid(ht, ctx, round(ctx.content_cm - _rw, 2))
            rt = _Target(right)
            p = rt.para()
            _spacing(p, after=0)
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            run = p.add_run()
            try:
                run.add_picture(io.BytesIO(ctx.photo_bytes), width=Cm(ctx.photo_cm))
            except Exception:  # noqa: BLE001
                pass
            p = t.para()
            _spacing(p, before=0, after=0, line=Pt(1))   # 1pt 高，只用来画横线
            _para_border_bottom(p, ctx.primary, 12)
            # 紧接着插一个无边框空段落：否则 Word 会把「头部下线」与
            # 首个小标题的下划线识别为同一边框块合并，只画最后一条
            sp = t.para()
            _spacing(sp, before=0, after=4, line=Pt(1))
            _render_sections(t, ctx, ctx.content_cm)
        else:
            _head_block(t, ctx, ctx.content_cm, with_extras=False)
            _basic_grid(t, ctx)
            _render_sections(t, ctx, ctx.content_cm)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def pdf_to_long_png(pdf_path: str, dpi: int = 150) -> bytes:
    """PDF 各页渲染后纵向拼接为一张 PNG 长图。"""
    import fitz
    from PIL import Image

    doc = fitz.open(pdf_path)
    try:
        images = []
        for page in doc:
            pm = page.get_pixmap(matrix=fitz.Matrix(dpi / 72.0, dpi / 72.0))
            images.append(Image.frombytes("RGB", (pm.width, pm.height), pm.samples))
    finally:
        doc.close()
    if not images:
        raise RuntimeError("PDF 无页面")
    if len(images) == 1:
        out = images[0]
    else:
        width = max(im.width for im in images)
        height = sum(im.height for im in images)
        out = Image.new("RGB", (width, height), "white")
        y = 0
        for im in images:
            out.paste(im, (0, y))
            y += im.height
    buf = io.BytesIO()
    out.save(buf, "PNG", optimize=True)
    return buf.getvalue()
