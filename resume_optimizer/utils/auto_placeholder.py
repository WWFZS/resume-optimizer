# -*- coding: utf-8 -*-
"""模板自动补占位符模块。

上传的模板可能没有任何 {{占位符}}（常见于从 PDF/Word 精美模板转换而来）。
本模块通过启发式扫描模板结构（文本框 / 表格 / 正文段落），
自动识别基本信息与简历模块（教育背景、工作经历、项目经历、专业技能等）的位置，
把对应 {{占位符}} 写入模板，使系统能够把优化后的内容填进去。

规则：
- 基本信息：把示例姓名/电话/邮箱/求职意向替换为 {{姓名}}/{{电话}}/{{邮箱}}/{{求职意向}}；
- 模块区块：找到模块标题（如“教育背景”）后的内容区，整段内容替换为 {{模块名}}；
- 内容区识别：文本框模板取标题框后的下一个非空文本框；表格模板取标题行后的内容行；
- 流式段落模板取标题后的正文段落。
"""
import os
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

PLACEHOLDER_RE = re.compile(r"\{\{\s*([^{}]+?)\s*\}\}")

# 模块标题别名 -> 标准占位符标签
# 别名需覆盖「中英混排标题」（「教育背景EDUCATION」「技能水平SKILLS」），
# 供 _section_title_label() 判定「整段就是标题」的显式标题。
SECTION_ALIASES: List[Tuple[str, str]] = [
    (r"教育背景|教育经历|学历背景|Education", "教育背景"),
    (r"工作经历|工作经验|工作履历|实习经历|相关经历|职业经历|Experience", "工作经历"),
    (r"项目经历|项目经验|项目实践|主要项目", "项目经历"),
    (r"专业技能|专业技能证书|技能证书|技能特长|技能水平|Skills", "专业技能"),
    (r"个人优势|自我评价|个人简介|个人总结|个人评价|优势特长|关于我|Summary", "个人优势"),
]

# 基本信息标签 -> 占位符
# 注意：别名必须「长别名在前」——否则「电话号码：」会先被「电话」截断，
# 正则要求「电话」后紧跟冒号，匹配不到 → 该字段永远不会被注入（30855 的电话）。
BASIC_FIELDS: List[Tuple[str, str]] = [
    (r"(?:我的)?姓\s*名|Name", "姓名"),
    (r"求职意向|意向岗位|应聘岗位|目标岗位", "求职意向"),
    (r"电话号码|联系电话|手机号码|手机号|手机|电话|Tel|Phone", "电话"),
    (r"邮箱地址|电子邮箱|邮箱|Email|Mail", "邮箱"),
    (r"现居地址|现居地|现居|居住地址|居住地|所在地|工作地点|工作地|城市", "所在地"),
]

# 基本信息行特征：标题与值同行（如“工作经验：4 年”“电话：138…”），不是模块区块标题
_BASIC_LINE_PAT = re.compile(r"(姓名|电话|手机|邮箱|性别|生日|现居|政治面貌|工作年限|工作经验|求职意向)\s*[:：]")

# 显式标题判定：
# - 整段文本去掉首尾空白后必须以模块别名开头；
# - 别名之后只允许「英文/分隔符」（如「教育背景EDUCATION」「技能水平SKILLS」），
#   这样「专业技能：熟悉 Word…」这类内容行不会被误判成标题；
# - 长度上限 _TITLE_MAX_LEN，避免把整段正文当成标题。
_TITLE_MAX_LEN = 40
_TITLE_TAIL_RE = re.compile(r"^[\sA-Za-z/&·\-—]*$")


def _section_title_label(text: str) -> Optional[str]:
    """判定一段文本是否为「模块标题行」，是则返回标准标签，否则返回 None。

    旧实现只在表格首行里找标题，文本框模板的标题（「教育背景EDUCATION」
    这种中英混排、且与内容同框的纯文本段落）完全没被识别，导致
    _detect_sections 返回空 → 该模块永远不会被注入占位符。
    """
    t = (text or "").strip()
    if not t or len(t) > _TITLE_MAX_LEN:
        return None
    if PLACEHOLDER_RE.search(t):
        return None
    for pat, label in SECTION_ALIASES:
        m = re.search(pat, t, re.IGNORECASE)
        if not m:
            continue
        if m.start() > 0:
            continue
        if not _TITLE_TAIL_RE.match(t[m.end():]):
            continue
        return label
    return None


def _is_field_line(text: str) -> bool:
    """是否属于「基本信息字段行」（如「电话号码：135…」「求职意向：行政主管」）。

    只认**已知基础字段**，不能用泛化的「xxx：yyy」——否则像
    「语言技能：英语CET6、粤语」这种真正的模块内容会被误当成字段行而跳过，
    导致该模块注入不出占位符。
    """
    t = (text or "").strip()
    if not t:
        return False
    if _is_basic_info_line(t):
        return True
    for pat, _label in BASIC_FIELDS:
        if re.match(r"^(?:%s)\s*[:：]" % pat, t, re.IGNORECASE):
            return True
    return False


def _is_basic_info_line(text: str) -> bool:
    """判断一行是否属于基本信息（标题:值 同行），而非模块区块标题。"""
    if _BASIC_LINE_PAT.search(text):
        return True
    if re.match(r"^工作经验\s*[:：]\s*\d", text):
        return True
    return False


def _para_text(p, doc) -> str:
    return (Paragraph(p, doc).text or "")


def _set_para_text(p, doc, text: str) -> None:
    """整段替换文本，保留首个 run 的格式。"""
    para = Paragraph(p, doc)
    if para.runs:
        first = para.runs[0]
        # 清空该 run 的文本节点
        for child in list(first._r):
            tag = child.tag.split("}")[-1]
            if tag in ("t", "br", "cr", "tab", "instrText"):
                first._r.remove(child)
        t_el = first._r.makeelement(qn("w:t"), {})
        t_el.set(qn("xml:space"), "preserve")
        t_el.text = text
        rPr = first._r.find(qn("w:rPr"))
        if rPr is not None:
            rPr.addnext(t_el)
        else:
            first._r.insert(0, t_el)
        for run in para.runs[1:]:
            for child in list(run._r):
                tag = child.tag.split("}")[-1]
                if tag in ("t", "br", "cr", "tab", "instrText"):
                    run._r.remove(child)
    else:
        para.add_run(text)


def _iter_textboxes(doc):
    """遍历模板中的文本框，返回 (anchor, [段落元素])，跳过 Fallback 重复副本。"""
    seen = set()
    for tx in doc.element.findall(".//" + qn("w:txbxContent")):
        # 从文本框自身向上找：若处于 mc:Fallback 分支则跳过（部分模板 Fallback 在 anchor 之下）
        c0 = tx
        in_fallback = False
        while c0 is not None:
            if c0.tag.split("}")[-1] == "Fallback":
                in_fallback = True
                break
            c0 = c0.getparent()
        if in_fallback:
            continue
        cur = tx
        while cur is not None and cur.tag.split("}")[-1] != "anchor":
            cur = cur.getparent()
        if cur is None:
            continue
        yield cur, list(tx.iter(qn("w:p")))


def _box_position(anchor) -> Tuple[Optional[int], Optional[int]]:
    """从 wp:anchor 读取文本框定位 (H, V)（EMU）；取不到返回 (None, None)。"""
    if anchor is None:
        return (None, None)
    h = v = None
    for pe in anchor.findall(qn("wp:positionH")):
        off = pe.find(qn("wp:posOffset"))
        if off is not None and off.text:
            try:
                h = int(off.text)
            except ValueError:
                h = None
    for pe in anchor.findall(qn("wp:positionV")):
        off = pe.find(qn("wp:posOffset"))
        if off is not None and off.text:
            try:
                v = int(off.text)
            except ValueError:
                v = None
    return (h, v)


def _collect_textbox_boxes(doc) -> List[Dict[str, object]]:
    """把模板里的文本框整理成有序列表：每项含段落、文本、锚点位置。

    顺序按「垂直位置优先、水平位置次之」排序，即视觉阅读顺序；
    取不到位置时回退到文档顺序。
    """
    boxes: List[Dict[str, object]] = []
    for idx, (anchor, paras) in enumerate(_iter_textboxes(doc)):
        units = []
        for p in paras:
            t = _para_text(p, doc).strip()
            if t:
                units.append((t, p))
        if not units:
            continue
        h, v = _box_position(anchor)
        boxes.append({"doc_idx": idx, "units": units, "h": h, "v": v})
    if boxes and all(b["v"] is not None for b in boxes):
        boxes.sort(key=lambda b: (b["v"], b["h"] if b["h"] is not None else 0))
    return boxes


def _detect_textbox_sections(doc) -> Dict[str, List[Tuple[int, object]]]:
    """在文本框模板里定位「模块标题」与「内容区」。

    两种结构都要覆盖：
    ① 同框结构：标题与内容在**同一个**文本框里（如 30855 的「教育背景EDUCATION」
       与其下的院校/学历/课程段落同框）——取标题之后的非标题段落作为内容区；
    ② 异框结构：标题单独一个框，内容在**相邻**的框里（如 30841 的「教育背景」框
       紧跟「{{教育背景}}」内容框）——取标题之后第一个非标题、非字段行的框。

    返回 {模块标签: [(标题序号, 内容区元素列表), ...]}。
    """
    boxes = _collect_textbox_boxes(doc)
    if not boxes:
        return {}

    # 预计算每个框的标题标签
    for b in boxes:
        b["title"] = _section_title_label(b["units"][0][0])

    found: Dict[str, List[Tuple[int, object]]] = {}

    # ① 同框：标题在框内第 i 段，其后段落属于该模块内容
    for bi, b in enumerate(boxes):
        units = b["units"]
        for i, (text, _p) in enumerate(units):
            label = _section_title_label(text)
            if not label:
                continue
            rest = []
            for j in range(i + 1, len(units)):
                rtext, rp = units[j]
                if _section_title_label(rtext):
                    break
                if _is_field_line(rtext):
                    continue
                rest.append(rp)
            if rest:
                found.setdefault(label, []).append((bi, rest))
            break

    # ② 异框：标题框之后、下一个标题框之前的第一个「内容框」
    title_box_idx = [bi for bi, b in enumerate(boxes) if b["title"]]
    for n, bi in enumerate(title_box_idx):
        label = boxes[bi]["title"]
        if label in found:
            continue  # 同框已处理
        end = title_box_idx[n + 1] if n + 1 < len(title_box_idx) else len(boxes)
        for cj in range(bi + 1, end):
            cb = boxes[cj]
            head = cb["units"][0][0]
            if _section_title_label(head):
                break
            if _is_field_line(head):
                continue
            if PLACEHOLDER_RE.search(" ".join(u[0] for u in cb["units"])):
                # 该框已经是占位符内容框（模板自带规范占位符），不再重复注入
                found.setdefault(label, []).append((bi, [u[1] for u in cb["units"]]))
                break
            found.setdefault(label, []).append((bi, [u[1] for u in cb["units"]]))
            break
    return found


def _detect_sections(doc, textbox_based: bool) -> Dict[str, List[Tuple[int, object]]]:
    """扫描模板，找出各模块标题的位置及其后内容区。

    返回 {模块标签: [(标题索引, 内容区元素), ...]}，索引为扫描序列下标。
    """
    found: Dict[str, List[Tuple[int, object]]] = {}
    # ── 文本框内容识别（含文本框的模板）──
    # 旧实现靠「示例文案特征」（如「公司名称」「熟悉OFFICE」）猜内容区，
    # 文本框模板的模块标题与内容常常同框/异框且文案千差万别，特征法几乎全落空
    # （17 套在售模板里旧实现检测结果**全为空**）。改用结构法：识别模块标题
    # 文本框，再按同框/异框两种结构取内容区。
    textbox_sections = _detect_textbox_sections(doc)
    if textbox_sections:
        found.update(textbox_sections)
    else:
        textbox_units = []
        for anchor, paras in _iter_textboxes(doc):
            for p in paras:
                t = _para_text(p, doc).strip()
                if t:
                    textbox_units.append((t, p))
        if textbox_units:
            units = []
            units = textbox_units
            content_groups = {}
            matched_idx = set()
            for i, (text, p) in enumerate(units):
                label = None
                if re.search(r"毕业院校|毕业时间|最高学历|所学专业|大学|学院|学校", text) and re.search(r"(20\d{2}|19\d{2})\s*[.\-~/至到年]", text):
                    label = "教育背景"
                elif re.search(r"公司名称|起止年月|工作描述|担任职务|网络科技|科技有限公司|岗位[:：]|负责|参与|策划", text):
                    label = "工作经历"
                elif re.search(r"熟悉OFFICE|熟悉Office|Microsoft Word|OFFICE软件|AUTO CAD|PHOTOSHOP", text):
                    label = "专业技能"
                if label and i not in matched_idx:
                    content_groups.setdefault(label, []).append(p)
                    matched_idx.add(i)
            for i, (text, p) in enumerate(units):
                if i in matched_idx:
                    continue
                if len(text) > 60 and not _is_basic_info_line(text):
                    content_groups.setdefault("个人优势", []).append(p)
                    matched_idx.add(i)
            for label, group in content_groups.items():
                found[label] = [(i, group) for i in range(len(group))]
    # ── 表格内容识别（含表格的模板，即使也有文本框）──
    if doc.tables:
        tables = doc.tables
        for ti, tb in enumerate(tables):
            if len(tb.rows) < 2:
                continue
            first_row = " | ".join(c.text.strip() for c in tb.rows[0].cells)
            if not first_row.strip() or PLACEHOLDER_RE.search(first_row):
                continue
            for pat, label in SECTION_ALIASES:
                m = re.search(pat, first_row, re.IGNORECASE)
                if m and m.start() <= 8:
                    # 标题行之后的内容行：只要不是新的模块标题行，都算内容
                    content_rows = []
                    for row in tb.rows[1:]:
                        rt = " ".join(c.text.strip() for c in row.cells if c.text.strip())
                        # 新模块标题行特征：单行文本且匹配模块别名（如“技能证书”“自我评价”）
                        is_new_header = False
                        if len(row.cells) >= 1 and rt and len(rt) <= 12:
                            if any(re.search(p2, rt, re.IGNORECASE) for p2, _ in SECTION_ALIASES):
                                is_new_header = True
                        if rt and not is_new_header:
                            content_rows.append(row)
                    if content_rows:
                        found.setdefault(label, []).append((ti, content_rows))
                    break
    return found


def _inject_basic_info(doc, textbox_based: bool, sections: Dict[str, List] = None) -> Dict[str, bool]:
    """把示例姓名/电话/邮箱等替换为占位符。返回 {标签: 是否成功}。

    sections：已识别出的模块内容区，用于排除内容区里的短词（避免误当姓名）。
    """
    injected: Dict[str, bool] = {}
    elems: List[Tuple[str, object]] = []
    if textbox_based:
        for tx in doc.element.findall(".//" + qn("w:txbxContent")):
            # 从文本框自身向上找 Fallback（部分模板 Fallback 在 anchor 之下）
            c0 = tx
            in_fb = False
            while c0 is not None:
                if c0.tag.split("}")[-1] == "Fallback":
                    in_fb = True
                    break
                c0 = c0.getparent()
            if in_fb:
                continue
            cur = tx
            while cur is not None and cur.tag.split("}")[-1] != "anchor":
                cur = cur.getparent()
            if cur is None:
                continue
            for p in tx.iter(qn("w:p")):
                t = _para_text(p, doc).strip()
                if t:
                    elems.append((t, p))
    else:
        for p in doc.element.iter(qn("w:p")):
            t = _para_text(p, doc).strip()
            if t:
                elems.append((t, p))
    # 表格单元格段落（基本信息常在表格里）
    for tb in doc.tables:
        seen_cells = set()
        for row in tb.rows:
            for cell in row.cells:
                if id(cell._tc) in seen_cells:
                    continue
                seen_cells.add(id(cell._tc))
                for p in cell.paragraphs:
                    t = _para_text(p._p, doc).strip()
                    if t:
                        elems.append((t, p._p))
    # 姓名：优先整段示例姓名（2-4 个汉字且非标签行）
    # 姓名：优先整段示例姓名（2-4 个汉字，且**不是**板块标题/字段标签）
    #
    # 旧实现用一份写死的 _HEADER_WORDS 白名单，漏了「工作经验」「基本信息」
    # 「实习经历」「联系方式」「个人信息」等标题行 —— 这些 2-4 个汉字的标题
    # 会被当成示例姓名替换成 {{姓名}}，随后被用户真实姓名填入，
    # 模板的章节标题就变成了「张三」。
    #
    # 现在改用标准别名库判定：能归一到「板块」或「基础字段」概念的行，
    # 一律不是姓名候选；另加结构性标题（基本信息/个人信息 等）排除。
    from . import standard_keys as _std
    try:
        from . import placeholders as _ph
    except ImportError:  # 归一模块可选，缺失时退回内置别名判定
        _ph = None

    name_pat = re.compile(r"^[一-龥·]{2,4}$")
    # 结构性标题：不归一到具体板块，但同样不是姓名
    _STRUCTURE_TITLES = {"基本信息", "个人信息", "联系方式", "联系信息",
                         "基本资料", "个人信息表", "基本信息表", "求职意向",
                         "自我评价", "个人简历", "简历信息"}
    # 候选姓名段必须**不在任何模块的内容区里**。
    # 否则像 30855「教育背景」框里的子标题「证书」「奖项」这种 2 字短词
    # 会被当成示例姓名替换成 {{姓名}}，既污染了教育背景内容，
    # 又让该模块的真实占位符 {{教育背景}} 被顶掉。
    name_candidates = []
    for t, p in elems:
        if not name_pat.match(t):
            continue
        if t in _STRUCTURE_TITLES:
            continue
        # 归一到板块概念（如「工作经验」「教育背景」）→ 是标题，不是姓名
        if _std.find_section_key(t):
            continue
        # 归一到基础字段概念（如「姓名」「电话」「性别」）→ 是标签，不是姓名
        if _std.find_field_key(t):
            continue
        # placeholders 侧的板块/字段别名（含后台自定义）也一并排除
        if _ph is not None:
            _concept = _ph.concept_of(t)
            if _concept and _concept[0] in ("section", "field"):
                continue
        if any(k in t for k in ("姓名", "电话", "邮箱", "求职", "性别", "生日", "现居")):
            continue
        name_candidates.append((t, p))
    # 排除落在模块内容区里的候选
    reserved = set()
    for _label, hits in (sections or {}).items():
        for _idx, content in hits:
            group = content if isinstance(content, (list, tuple)) else [content]
            for item in group:
                reserved.add(id(item))
    for t, p in name_candidates:
        if id(p) in reserved:
            continue
        if t not in injected:
            _set_para_text(p, doc, "{{姓名}}")
            injected["姓名"] = True
            break

    # 姓名 + 求职意向同行（如“某某某    求职意向：行政主管 岗位”）：替换行首姓名
    if "姓名" not in injected:
        mixed_pat = re.compile(r"^([\u4e00-\u9fa5·]{2,4})\s{2,}(求职意向|意向岗位|应聘岗位)\s*[:：]")
        for t, p in elems:
            m = mixed_pat.match(t)
            if m and not any(k in t for k in ("姓名", "电话", "邮箱", "性别", "生日")):
                label = m.group(2)
                _set_para_text(p, doc, "{{姓名}}    " + label + "：" + t[m.end():])
                injected["姓名"] = True
                break
    # 姓名：标签形式（姓名：某某某 / 姓 名：xxx / 我的姓名：xxx）
    if "姓名" not in injected:
        name_label_pat = re.compile(r"^(?:我的)?姓\s*名\s*[:：]\s*([^\s|｜]+)")
        for t, p in elems:
            m = name_label_pat.match(t)
            if m:
                _set_para_text(p, doc, t.replace(m.group(0), "姓名：" + "{{姓名}}"))
                injected["姓名"] = True
                break
    # 混合行：姓名：示例 / 电话：示例 等
    for t, p in elems:
        for pat, label in BASIC_FIELDS:
            if label == "姓名":
                continue  # 姓名已单独处理
            if label in injected:
                continue
            cur_text = _para_text(p, doc)  # 读取当前文本（前面的替换可能已修改该段）
            m = re.search(r"(%s)\s*[:：]\s*([^\s|｜]+)" % pat, cur_text, re.IGNORECASE)
            if m:
                _set_para_text(p, doc, cur_text.replace(m.group(0), m.group(1) + "：{{" + label + "}}"))
                injected[label] = True
                continue
            # 标签单独成行且值为空（如“手机：”“邮箱：”）：直接补占位符
            lm = re.match(r"^(%s)\s*[:：]\s*$" % pat, cur_text.strip(), re.IGNORECASE)
            if lm:
                _set_para_text(p, doc, lm.group(1) + "：{{" + label + "}}")
                injected[label] = True
    return injected


def _inject_section_content(doc, sections: Dict[str, List], textbox_based: bool) -> Dict[str, bool]:
    """把各模块内容区替换为 {{标签}}。

    幂等：内容区里若已有 {{...}} 占位符（模板自带规范占位符），跳过该模块，
    绝不用新占位符覆盖已有占位符（否则会把用户可用的 {{工作经历}} 变成 {{个人优势}}）。
    """
    injected: Dict[str, bool] = {}
    for label, hits in sections.items():
        if label in injected:
            continue
        for _, content in hits:
            if content is None:
                continue
            try:
                group = content if isinstance(content, (list, tuple)) else [content]
                # 幂等保护：内容区已有占位符 → 该模块不再注入
                has_ph = False
                for item in group:
                    if hasattr(item, "cells") and hasattr(item, "_tr"):
                        for cell in item.cells:
                            if any("{{" in (c.text or "") for c in cell._tc.iter(qn("w:t"))):
                                has_ph = True
                                break
                    else:
                        if "{{" in (_para_text(item, doc) or ""):
                            has_ph = True
                    if has_ph:
                        break
                if has_ph:
                    injected[label] = True
                    break
                rows_to_remove = []
                for gi, item in enumerate(group):
                    if hasattr(item, "cells") and hasattr(item, "_tr"):
                        # 表格行：合并单元格会重复，按唯一 tc 去重
                        seen_cells = set()
                        unique_cells = []
                        for cell in item.cells:
                            if id(cell._tc) not in seen_cells:
                                seen_cells.add(id(cell._tc))
                                unique_cells.append(cell)
                        # 首个内容行第一格放占位符，其余清空
                        for ci, cell in enumerate(unique_cells):
                            paras = cell.paragraphs
                            for pi, p in enumerate(paras):
                                if gi == 0 and ci == 0 and pi == 0:
                                    _set_para_text(p._p, doc, "{{" + label + "}}")
                                else:
                                    _set_para_text(p._p, doc, "")
                        if gi > 0:
                            rows_to_remove.append(item._tr)
                    else:
                        # 文本框段落
                        _set_para_text(item, doc, "{{" + label + "}}" if gi == 0 else "")
                for tr in rows_to_remove:
                    parent = tr.getparent()
                    if parent is not None:
                        parent.remove(tr)
                injected[label] = True
                break
            except Exception:
                continue
    return injected


def auto_inject_placeholders(docx_path: Path, out_path: Path) -> dict:
    """对模板文件自动补占位符，返回注入结果。

    result = {
        "injected": {"姓名": True, "工作经历": True, ...},
        "labels_found": [...],           # 注入后模板中的全部占位符标签
        "missing": [...],                # 仍缺失的必选标签
    }
    """
    doc = Document(str(docx_path))
    textbox_based = (len(doc.element.findall(".//" + qn("w:txbxContent"))) > 0
                     or len(doc.tables) > 0)
    sections = _detect_sections(doc, textbox_based)
    basic = _inject_basic_info(doc, textbox_based, sections)
    section_injected = _inject_section_content(doc, sections, textbox_based)
    injected = {**basic, **section_injected}
    doc.save(str(out_path))

    labels = set()
    doc2 = Document(str(out_path))
    for pe in doc2.element.iter(qn("w:p")):
        for m in PLACEHOLDER_RE.finditer(_para_text(pe, doc2) or ""):
            labels.add(m.group(1).strip())
    required = {"姓名", "联系方式", "个人优势", "工作经历", "教育背景"}
    missing = sorted(required - labels)
    return {
        "injected": injected,
        "labels_found": sorted(labels),
        "missing": missing,
        "textbox_based": textbox_based,
    }


def detect_metadata(docx_path: Path, fallback_name: str = "") -> dict:
    """自动识别模板元数据：名称、模块清单、场景说明。

    名称优先取文档标题属性，其次取文件名；
    模块通过扫描模板内标题/内容特征识别；
    场景说明按岗位关键词与模块组合生成。
    """
    try:
        doc = Document(str(docx_path))
        title = ""
        try:
            cp = doc.core_properties
            title = (cp.title or "").strip()
        except Exception:
            pass
        textbox_based = len(doc.element.findall(".//" + qn("w:txbxContent"))) > 0
        # 收集全部文本（文本框 + 表格 + 段落）
        texts = []
        seen = set()
        for tx in doc.element.findall(".//" + qn("w:txbxContent")):
            c0 = tx
            in_fb = False
            while c0 is not None:
                if c0.tag.split("}")[-1] == "Fallback":
                    in_fb = True
                    break
                c0 = c0.getparent()
            if in_fb:
                continue
            for p in tx.iter(qn("w:p")):
                t = _para_text(p, doc).strip()
                if t:
                    texts.append(t)
        for tb in doc.tables:
            for row in tb.rows:
                for cell in row.cells:
                    t = cell.text.strip()
                    if t:
                        texts.append(t)
        for p in doc.paragraphs:
            t = (p.text or "").strip()
            if t:
                texts.append(t)
        joined = "\n".join(texts)
        # 模块识别
        # 注意：这里识别的是「模板里出现过的模块名」，**不代表模板里有对应占位符**。
        # 历史上调用方把本结果直接当 DB 占位符标签写入，导致模板里只有静态标题、
        # 却没有 {{占位符}} 的模块被记进 template_placeholders → 填充时静默跳过，
        # 用户内容凭空消失。现在调用方必须以磁盘实际扫描结果为准。
        modules = []
        for pat, label in SECTION_ALIASES:
            if re.search(pat, joined, re.IGNORECASE):
                modules.append(label)
        # 基本信息字段
        for pat, label in BASIC_FIELDS:
            if re.search(pat, joined, re.IGNORECASE) and label not in modules:
                modules.append(label)
        # 去重保序
        seen_m = set()
        modules = [m for m in modules if not (m in seen_m or seen_m.add(m))]
        # 名称：文档标题（需是合法中文名称，过滤水印/网址/uuid）> 文件名
        def _is_junk_title(t):
            if not t or len(t) > 30:
                return True
            if "www" in t.lower() or ".com" in t.lower() or "http" in t.lower():
                return True
            if re.fullmatch(r"[0-9a-f]{16,}", t):
                return True
            if "简历模板资源" in t or "下载" == t or "资源网" in t:
                return True
            return False

        fname = (fallback_name or Path(docx_path).name)
        fname = fname.rsplit(".", 1)[0] if "." in fname else fname
        fname_clean = (fname.replace("简历模板", "").replace("免费", "").replace("岗位求职", "")
                       .replace("_", "").strip())
        if fname_clean.endswith("简历"):
            fname_clean = fname_clean[: -len("简历")].strip()
        name = (title if not _is_junk_title(title) else "") or fname_clean or fname or "未命名模板"
        # 场景说明
        scene_parts = []
        if re.search(r"校招|应届|毕业生", joined):
            scene_parts.append("校招应届")
        if re.search(r"社招|求职|总监|经理", joined):
            scene_parts.append("社招求职")
        if re.search(r"行政|人事|文员|秘书", joined):
            scene_parts.append("行政职能")
        if re.search(r"JAVA|开发|工程师|技术|运维", joined):
            scene_parts.append("技术岗位")
        if re.search(r"造价|工程|施工|建筑", joined):
            scene_parts.append("工程岗位")
        if re.search(r"物流|仓管|运输", joined):
            scene_parts.append("物流岗位")
        scene = "、".join(scene_parts) if scene_parts else "通用求职"
        return {
            "name": name,
            "modules": modules,
            "description": f"{scene}简历模板，包含：{'、'.join(modules) if modules else '基础信息'}等模块",
        }
    except Exception:
        return {"name": Path(docx_path).stem, "modules": [], "description": "简历模板"}


# ── 只读诊断（不写库、不改文件）───────────────────────────────


def scan_placeholder_labels(docx_path: Path) -> List[str]:
    """扫描 docx 里实际存在的 {{占位符}} 标签（权威来源）。

    这是「磁盘上的模板文件实际能填哪些模块」的唯一事实来源：
    DB 里记录的标签必须是本函数结果的子集，否则填充阶段会静默跳过。
    """
    doc = Document(str(docx_path))
    labels = set()
    for pe in doc.element.iter(qn("w:p")):
        for m in PLACEHOLDER_RE.finditer(_para_text(pe, doc) or ""):
            labels.add(m.group(1).strip())
    return sorted(labels)


def diagnose_template_placeholders(template_id, name: str, file_path: str,
                                   docx_path: Path, db_labels) -> dict:
    """只读诊断单个模板：对比 DB 标签与 docx 实际占位符。

    返回 {
      "template_id", "name", "file_path",
      "db_labels",       库内 template_placeholders 记录的标签
      "actual_labels",   docx 里实际存在的 {{占位符}}
      "missing",         DB 有但 docx 没有 → 用户内容永远填不进去（P0）
      "extra",           docx 有但 DB 没有 → 填了但没登记（反向不一致）
      "ok",              是否一致
    }
    """
    db_set = set(db_labels or [])
    try:
        actual = set(scan_placeholder_labels(Path(docx_path)))
        err = ""
    except Exception as exc:  # noqa: BLE001
        actual = set()
        err = str(exc)[:120]
    missing = sorted(db_set - actual)
    extra = sorted(actual - db_set)
    return {
        "template_id": template_id,
        "name": name,
        "file_path": file_path,
        "db_labels": sorted(db_set),
        "actual_labels": sorted(actual),
        "missing": missing,
        "extra": extra,
        "ok": not missing and not err,
        "error": err,
    }


def diagnose_all_templates(base_dir=None, only_active: bool = True) -> List[dict]:
    """只读诊断全部模板：对比 DB 标签与 docx 实际占位符。

    仅 SELECT，不做任何 INSERT/UPDATE/DELETE，不改任何模板文件。
    返回 diagnose_template_placeholders 的结果列表。
    """
    from pathlib import Path as _Path
    from . import db as _db

    base = _Path(base_dir) if base_dir else _Path(__file__).resolve().parent.parent
    with _db.get_conn() as conn:
        with conn.cursor() as cur:
            if only_active:
                cur.execute(
                    "SELECT id, name, file_path FROM templates "
                    "WHERE is_active = 1 AND is_vip = 0 ORDER BY id")
            else:
                cur.execute(
                    "SELECT id, name, file_path FROM templates "
                    "WHERE review_status <> 'rejected' ORDER BY id")
            tpls = cur.fetchall()
            if not tpls:
                return []
            marks = ",".join(["%s"] * len(tpls))
            cur.execute(
                "SELECT template_id, label FROM template_placeholders "
                f"WHERE template_id IN ({marks}) ORDER BY template_id, id",
                [int(t["id"]) for t in tpls])
            phs = cur.fetchall()
    db_labels: Dict[int, List[str]] = {}
    for r in phs:
        db_labels.setdefault(int(r["template_id"]), []).append(r["label"])
    out = []
    for t in tpls:
        fp = base / str(t["file_path"]).replace("/", os.sep)
        if not fp.exists():
            out.append({
                "template_id": t["id"], "name": t["name"],
                "file_path": t["file_path"], "db_labels": db_labels.get(t["id"], []),
                "actual_labels": [], "missing": db_labels.get(t["id"], []),
                "extra": [], "ok": False,
                "error": f"模板文件不存在: {fp}",
            })
            continue
        out.append(diagnose_template_placeholders(
            t["id"], t["name"], t["file_path"], fp, db_labels.get(t["id"], [])))
    return out


if __name__ == "__main__":  # 只读诊断入口：python -m utils.auto_placeholder
    _rows = diagnose_all_templates()
    print("=" * 100)
    print("在售模板占位符一致性诊断（只读，不写库）")
    print("=" * 100)
    for _r in _rows:
        print(f"[{'OK ' if _r['ok'] else 'BAD'}] 模板 {_r['template_id']} [{_r['name']}]")
        print(f"    DB标签 ({len(_r['db_labels'])}): {_r['db_labels']}")
        print(f"    docx实际 ({len(_r['actual_labels'])}): {_r['actual_labels']}")
        print(f"    ★DB有/docx无 (永远填不进去): {_r['missing']}")
        print(f"      docx有/DB无 (反向): {_r['extra']}")
        if _r.get("error"):
            print(f"      错误: {_r['error']}")
    _bad = [r for r in _rows if not r["ok"]]
    _tot = sum(len(r["missing"]) for r in _bad)
    print("=" * 100)
    print(f"受影响模板: {len(_bad)}/{len(_rows)}, 永远填不进去的字段总数: {_tot}")
