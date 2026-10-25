# -*- coding: utf-8 -*-
"""模板入库校验：用标准长内容测试简历渲染，检查单页/溢出/模块完整/照片适配。"""
import os
import re
import tempfile
from pathlib import Path
from typing import Dict, List, Union

import fitz
from docx import Document
from docx.oxml.ns import qn

from . import converter
from .template_render import MODULE_LABELS, render_template
from utils.tmpfiles import safe_unlink

SAFE_BOTTOM = 842.0

# 标准长内容测试简历：覆盖全部模块、内容偏长，用于压测模板
TEST_RESUME: Dict[str, str] = {
    "name": "测试求职者",
    "contact": "13800000000 | test@example.com | 上海市",
    "job_title": "数据分析师",
    "summary": "具备数据分析与业务洞察能力，擅长通过数据驱动决策，熟悉数据清洗、可视化与报告编制，"
               "拥有跨部门协作与项目管理经验，能够快速学习新技术并落地应用，结果导向，注重细节与质量，"
               "善于沟通表达，愿意长期深耕数据领域，为团队持续创造价值。",
    "work_experience": "高级数据分析师 某某科技有限公司 2021.03-2024.06\n"
                       "- 负责搭建公司经营数据分析体系，覆盖销售、运营、用户三大板块，输出月度经营分析报告。\n"
                       "- 主导用户留存分析专项，通过漏斗分析与 cohort 分析定位流失关键节点，推动产品迭代后留存率提升 12%。\n"
                       "- 搭建自动化报表看板（SQL + Python + BI 工具），将常规报表产出时间从 2 天缩短至 2 小时。\n"
                       "- 与产品、运营、研发多部门协作，支持 10+ 个业务项目的数据需求与效果评估。\n"
                       "数据分析师 某网络科技有限公司 2018.07-2021.02\n"
                       "- 负责业务数据仓库的 ETL 开发与维护，保障数据质量与时效性。\n"
                       "- 基于用户行为数据开展专题分析，输出可落地的运营策略建议。\n"
                       "- 独立完成 A/B 实验设计与效果评估，支撑多个营销活动的优化决策。",
    "projects": "校园数据竞赛项目 2020.09-2020.12\n"
                "- 作为队长带领 4 人团队完成「用户购买行为预测」竞赛项目，使用 Python 构建特征工程与 XGBoost 模型，获得省级二等奖。\n"
                "- 负责数据清洗、特征筛选与模型调参，撰写项目报告并进行答辩展示。\n"
                "- 通过该项目的历练，系统掌握了数据分析全流程，锻炼了团队协作与项目推进能力。",
    "education": "某某大学 数据科学与大数据技术 本科 2016.09-2020.06\n"
                 "主修课程：数据结构、数据库原理、操作系统、计算机网络、Python 程序设计、机器学习、数据挖掘、统计学、"
                 "数据可视化、大数据技术、算法设计与分析、软件工程、离散数学、概率论、线性代数、高等数学等。",
    "skills": "编程语言：Python、SQL、Java\n"
              "数据分析：Pandas、NumPy、Matplotlib、Power BI、Tableau、Excel 高级功能\n"
              "机器学习：Scikit-learn、XGBoost、特征工程、模型评估\n"
              "其他：Linux、Git、Hive、Spark 基础\n"
              "证书：大学英语六级、计算机二级（Python）、阿里云大数据助理工程师认证",
}



# ── 标准长内容 → 模板占位符的映射 ──
# 模板里的占位符标签是中文（如「工作经历」「教育背景」「个人优势」），
# 而 TEST_RESUME 的键是英文（work_experience/education/summary...），
# 直接按字面取交集恒为空 —— 那会让「长内容单页」校验只用姓名+电话两条短内容
# 渲染，任何真实长内容下会溢页的模板都能通过审核。
# 这里按「概念」映射：模板标签先归一到标准概念，再取对应的长内容。
_LONG_CONTENT_BY_CONCEPT = {
    # 板块
    'work_history': 'work_experience',
    'education_info': 'education',
    'skill_info': 'skills',
    'self_evaluate': 'summary',
    'campus_exp': 'projects',
    'honor_cert': 'projects',
    'projects': 'projects',
    # 基础字段
    'name': 'name',
    'contact': 'contact',
    'phone': 'contact',
    'email': 'contact',
    'intention_job': 'job_title',
    'city': 'contact',
    'location': 'contact',
}

# 兜底：标签概念识别不到时按字面同义词取
_LONG_CONTENT_BY_LABEL = {
    '工作经历': 'work_experience', '工作经验': 'work_experience',
    '实习经历': 'work_experience', '工作履历': 'work_experience',
    '教育背景': 'education', '教育经历': 'education', '学历': 'education',
    '专业技能': 'skills', '技能': 'skills', '职业技能': 'skills',
    '个人优势': 'summary', '自我评价': 'summary', '个人评价': 'summary',
    '项目经历': 'projects', '项目经验': 'projects',
    '校园经历': 'projects', '证书荣誉': 'projects',
    '求职意向': 'job_title', '应聘岗位': 'job_title', '意向岗位': 'job_title',
}


def build_test_content_map(labels) -> Dict[str, str]:
    """按模板实际拥有的占位符标签，组装「标准长内容」content_map。

    标签 → 概念（utils.placeholders.concept_of）→ 标准内容键 → TEST_RESUME 值。
    姓名/联系方式单独处理（联系方式需拆出电话与邮箱）。
    """
    from . import placeholders

    content_map: Dict[str, str] = {}
    for lab in labels:
        concept = placeholders.concept_of(lab)
        key = None
        if concept:
            key = _LONG_CONTENT_BY_CONCEPT.get(concept[1])
        if key is None:
            key = _LONG_CONTENT_BY_LABEL.get(lab.strip())
        if key is None:
            continue
        content_map[lab] = TEST_RESUME.get(key, '')

    # 姓名与联系方式：保证模板拿到可核对的值
    content_map['姓名'] = TEST_RESUME['name']
    content_map['name'] = TEST_RESUME['name']
    content_map['联系方式'] = TEST_RESUME['contact']
    content_map['contact'] = TEST_RESUME['contact']
    # 拆出电话/邮箱，供「电话：{{电话}}」「邮箱：{{邮箱}}」分开展示的模板
    m_phone = re.search(r'[0-9\- ]{7,16}', TEST_RESUME['contact'])
    m_email = re.search(r'[\w.+-]+@[\w.-]+\.\w+', TEST_RESUME['contact'])
    if m_phone:
        content_map['电话'] = m_phone.group(0).strip()
        content_map['phone'] = content_map['电话']
        content_map['手机'] = content_map['电话']
        content_map['手机号'] = content_map['电话']
    if m_email:
        content_map['邮箱'] = m_email.group(0).strip()
        content_map['email'] = content_map['邮箱']
    return content_map

def _measure_pdf(pdf_path: str) -> dict:
    pdf = fitz.open(pdf_path)
    page = pdf[0]
    if pdf.page_count > 1:
        return {"pages": pdf.page_count, "bottom": SAFE_BOTTOM + 200.0, "chars": 0}
    bottom = 0.0
    for b in page.get_text("blocks"):
        bottom = max(bottom, b[3])
    for info in page.get_image_info():
        bottom = max(bottom, info["bbox"][3])
    return {"pages": pdf.page_count, "bottom": bottom,
            "chars": len(page.get_text().strip())}


def _check_modules_rendered(docx_path: str, section_order: List[str]) -> List[str]:
    doc = Document(docx_path)
    text = "\n".join(p.text for p in doc.paragraphs)
    missing = []
    for key in section_order:
        label = MODULE_LABELS.get(key, key)
        if label and label not in text:
            missing.append(label)
    return missing


def validate_flow_template(config: dict, photo_path: Union[str, Path, None] = None) -> dict:
    """校验流式配置模板：渲染标准长内容，检查单页/溢出/模块完整/照片。"""
    section_order = config.get("section_order") or list(MODULE_LABELS.keys())
    tmp_dir = Path(tempfile.mkdtemp(prefix="tpl_validate_"))
    out_docx = tmp_dir / "test.docx"
    checks = []
    try:
        stats = render_template(TEST_RESUME, config, photo_path, str(out_docx))
        checks.append({"name": "渲染完成", "ok": True, "detail": f"字号 {stats.get('body_pt')}pt"})
        # 转 PDF 体检
        pdf_path = tmp_dir / "test.pdf"
        converter.docx_to_pdf(str(out_docx), str(pdf_path))
        m = _measure_pdf(str(pdf_path))
        checks.append({"name": "强制单页", "ok": m["pages"] == 1,
                       "detail": f"{m['pages']} 页"})
        checks.append({"name": "内容不溢出", "ok": m["bottom"] <= SAFE_BOTTOM,
                       "detail": f"内容底部 {m['bottom']:.0f}pt / 安全阈值 {SAFE_BOTTOM:.0f}pt"})
        missing = _check_modules_rendered(str(out_docx), section_order)
        checks.append({"name": "模块完整", "ok": not missing,
                       "detail": "缺失：" + "、".join(missing) if missing else "全部模块已渲染"})
        if photo_path:
            pdf = fitz.open(str(pdf_path))
            imgs = pdf[0].get_images()
            checks.append({"name": "照片适配", "ok": len(imgs) > 0,
                           "detail": f"渲染图片 {len(imgs)} 张"})
            pdf.close()
        ok = all(c["ok"] for c in checks)
        return {"ok": ok, "checks": checks}
    finally:
        for f in list(tmp_dir.iterdir()):
            safe_unlink(f)
        try:
            tmp_dir.rmdir()
        except BaseException:  # noqa: BLE001
            pass


def validate_template_file(template_path: Union[str, Path]) -> dict:
    """校验旧式（文本框）模板文件：结构 + 长内容填充渲染体检。"""
    from . import filler
    path = Path(template_path)
    checks = []
    doc = Document(str(path))
    sec = doc.sections[0]
    page_h_pt = sec.page_height / 12700.0
    # 按真实页面高度判定：安全线统一为 A4 页高 842pt，
    # 模板自身单页（含设计排满页面）即通过，避免误伤
    safe_bottom = 842.0
    w = sec.page_width / 914400 * 25.4
    h = sec.page_height / 914400 * 25.4
    checks.append({"name": "A4 页面", "ok": abs(w - 210) < 3 and abs(h - 297) < 3,
                   "detail": f"{w:.0f}x{h:.0f}mm"})
    # 占位符检查
    labels = set()
    for pe in doc.element.iter(qn("w:p")):
        from docx.text.paragraph import Paragraph
        para = Paragraph(pe, doc)
        for m in filler.PLACEHOLDER_RE.finditer(para.text or ""):
            labels.add(m.group(1).strip())
    # 必选占位符：按「概念」判定，而不是死抠字面。
    # 旧实现只要命中 2 个核心标签就整段豁免（required = set()），
    # 导致无姓名、无联系方式、无教育背景的模板也能上架。
    #
    # 新规则：模板要能承载用户数据，至少具备
    #   1) 姓名位：{{姓名}} 占位符，或模板里已有的「姓名」标签行
    #   2) 联系方式位：{{电话}}/{{邮箱}}/{{联系方式}} 占位符，或「电话/邮箱/手机/联系」
    #      标签行，或模板里本来就有可被覆盖的联系方式数据（示例手机号/邮箱）
    #   3) 至少一个可填充大模块（工作经历/教育背景/专业技能/项目经历/个人优势/校园经历/证书荣誉）
    # 标签概念用 utils.placeholders.concept_of 归一，兼容同义写法。
    from . import placeholders as _ph

    _concepts = set()
    for _lab in labels:
        _c = _ph.concept_of(_lab)
        if _c:
            _concepts.add(_c[1])

    # 模板全文（用于识别硬编码的「姓名：」「电话：」等标签行与示例数据）
    _parts = []
    for _pe in doc.element.iter(qn('w:p')):
        _parts.append(''.join((c.text or '') for c in _pe.iter()
                               if c.tag.split('}')[-1] == 't'))
    _full_text = '\n'.join(_parts)

    # 模板里字段标签常写成「姓    名：」这种带空格的排版（见模板 30844），
    # 比较前先把全文的空白去掉，否则会漏判。
    _flat = re.sub(r'\s+', '', _full_text)
    _has_name = ('name' in _concepts) or ('姓名' in _flat)
    _has_contact = bool({'contact', 'phone', 'email'} & _concepts) or any(
        _k in _flat for _k in ('电话', '邮箱', '手机', '联系')
    )
    if not _has_contact:
        # 模板没有联系方式标签，但自带示例手机号/邮箱（可被用户数据覆盖）
        _has_contact = bool(re.search(r'[0-9]{7,15}', _full_text)) or bool(
            re.search(r'[\w.+-]+@[\w.-]+\.\w+', _full_text))
    _MODULE_CONCEPTS = {'work_history', 'education_info', 'skill_info',
                        'self_evaluate', 'projects', 'campus_exp', 'honor_cert'}
    _has_module = bool(_MODULE_CONCEPTS & _concepts)

    missing = []
    if not _has_name:
        missing.append('姓名')
    if not _has_contact:
        missing.append('联系方式')
    if not _has_module:
        missing.append('至少一个可填充模块（工作经历/教育背景/专业技能等）')

    checks.append({"name": "必选占位符", "ok": not missing,
                   "detail": "缺失：" + "、".join(sorted(missing)) if missing else f"共 {len(labels)} 个占位符"})
    # 外链图片检查
    ext = []
    for rel in doc.part.rels.values():
        target = str(getattr(rel, "target_ref", ""))
        if target.startswith(("http://", "https://", "file://")) or re.match(r"^[A-Za-z]:\\\\", target):
            ext.append(target)
    checks.append({"name": "图片内嵌", "ok": not ext,
                   "detail": "存在外链：" + "、".join(ext[:2]) if ext else "无外链图片"})
    # 长内容填充渲染体检
    tmp_dir = Path(tempfile.mkdtemp(prefix="tpl_validate_"))
    try:
        content_map = build_test_content_map(labels)
        out_docx = tmp_dir / "filled.docx"
        filler.fill_template(content_map, path, out_docx)
        pdf_path = tmp_dir / "filled.pdf"
        converter.docx_to_pdf(str(out_docx), str(pdf_path))
        m = _measure_pdf(str(pdf_path))
        # 只要求渲染为 1 页即通过（不再用底部安全线卡模板）：
        # 模板原生排版保留，填充内容后只要 Word 能渲染为单页即可上架
        fits_one_page = m["pages"] == 1
        if not labels:
            # 模板本身没有占位符：填充内容为空，单页测量反映的是模板静态版式。
            # 此时只要模板自身在单页内（页面高度内）就算通过，提示人工补充占位符即可。
            checks.append({
                "name": "长内容单页",
                "ok": m["pages"] == 1 and m["bottom"] <= page_h_pt + 2.0,
                "detail": f"{m['pages']} 页，底部 {m['bottom']:.0f}pt（模板自身版式，页面高 {page_h_pt:.0f}pt）",
            })
            checks.append({
                "name": "可填充性",
                "ok": False,
                "detail": "模板无 {{占位符}}，系统无法填入内容，需人工补充占位符或改用带占位符的模板",
            })
        else:
            checks.append({"name": "长内容单页", "ok": fits_one_page,
                           "detail": f"{m['pages']} 页，底部 {m['bottom']:.0f}pt"})
    except Exception as exc:  # noqa: BLE001
        # 常见失败：模板的固定文本框装不下标准长内容，Word 打开/转换时直接报 COM 错。
        # 这说明模板版式容量不足（用户填满几个板块后同样会导出失败），
        # 应当明确告诉管理员「模板装不下长内容」，而不是抛一串 COM 错误码。
        _msg = str(exc)
        if "Microsoft Word" in _msg or "-2147352567" in _msg:
            _detail = ("模板版式装不下标准长内容（固定文本框容量不足）："
                       "用户填满多个板块后导出可能失败，建议改用可自适应文本框的模板")
        else:
            _detail = _msg[:120]
        checks.append({"name": "长内容渲染", "ok": False, "detail": _detail})
    finally:
        for f in list(tmp_dir.iterdir()):
            safe_unlink(f)
        try:
            tmp_dir.rmdir()
        except BaseException:  # noqa: BLE001
            pass
    return {"ok": all(c["ok"] for c in checks), "checks": checks}
