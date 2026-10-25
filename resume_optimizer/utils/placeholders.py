# -*- coding: utf-8 -*-
"""占位符标签解析与归一（全系统唯一实现点）。

背景：模板里 {{标签}} 的写法千差万别（{{ 工作经验 }}、{{项目经验}}、{{name}}），
而填充用的 content_map 只有固定的标准键。若直接拿原始标签查 content_map，
查不到就会把整段清空 → 产出「排版正常但内容为空」的简历，且不抛异常、无日志（P0）。

本模块提供：
- PLACEHOLDER_RE：统一的 {{标签}} 正则（filler / template_editor 共用一份）；
- canonical_key() / resolve_label()：把原始标签归一成 content_map 的标准键。
  别名来源唯一：utils/standard_keys.py（内置别名 + 后台 standard_aliases 表），
  不另起一套同义词库。

resolve_label 的归一顺序：
① 清洗（去首尾/全角空白）→ ② 概念识别（英文标准键 → 联系方式 → 板块别名 →
字段别名；调用方已知概念时优先用调用方的）→ ③ 原始标签精确命中 content_map
（模板后台配置的标签必然在 content_map 里；英文标准键除外，统一归一到中文标准键）
→ ④ 概念候选键在 content_map 里命中 → ⑤ 都没命中：概念已知但 cm 缺键 →
kind='missing'；完全未知 → kind='unknown'。
调用方据此保留模板原文并记录告警，绝不清空用户内容。
"""
import re
from collections import namedtuple
from typing import Dict, List, Optional, Tuple

from . import standard_keys

# 统一占位符正则：捕获组 + 不跨花括号 + 自动吃掉标签两侧空白
# （filler.py / template_editor.py 共用，避免两条链路解析结果不一致）
PLACEHOLDER_RE = re.compile(r"\{\{\s*([^{}]+?)\s*\}\}")

# 归一结果：
#   raw        原始标签（已清洗）
#   key        content_map 中实际命中的键（未命中为 None）
#   canonical  概念的标准中文键（content_map 缺键时也有值，用于提示文案）
#   kind       exact（原始标签精确命中）/ alias（经别名归一命中）/
#              missing（概念已知但 content_map 缺键）/ unknown（完全未知）
#   reason     归一依据（可读）
#   candidates content_map 候选键（按优先级）
Resolution = namedtuple("Resolution", "raw key canonical kind reason candidates")

# ── 板块概念 → content_map 候选键（首个为标准中文键）──
# 键名取自 standard_keys.MODULES（education_info/work_history/skill_info/
# campus_exp/honor_cert/self_evaluate）；projects 是 main.py 数据层的顶层模块
# （不在 MODULES 内），模板常写成「项目经验」，单独补一组别名。
SECTION_CANDIDATES: Dict[str, List[str]] = {
    "work_history": ["工作经历", "工作履历", "work_experience", "work_history"],
    "education_info": ["教育背景", "教育经历", "education", "education_info"],
    "skill_info": ["专业技能", "技能合集", "skills", "skill_info"],
    "self_evaluate": ["个人优势", "自我评价", "summary", "self_evaluate"],
    "campus_exp": ["校园经历", "campus_exp"],
    "honor_cert": ["证书荣誉", "honor_cert"],
    "projects": ["项目经历", "项目经验", "projects"],
}

# standard_keys.SECTION_ALIASES 之外补的板块别名（项目经历族）
EXTRA_SECTION_ALIASES: Dict[str, List[str]] = {
    "projects": ["项目经历", "项目经验", "项目实践", "主要项目", "项目描述", "projects"],
}

# ── 字段概念 → content_map 候选键（首个为标准中文键）──
# 字段集合以 standard_keys.BASIC_FIELDS 为准；键名按 content_map 的中文标签对齐
# （phone 用「电话」而不是 BASIC_FIELDS 的「手机号」：content_map 里存的是「电话」；
#  city 用「所在地」而不是「现居城市」：main.py 写入 content_map 的就是「所在地」）。
# 末尾补标准 key，兼容 fill_mapped_template 那种「content_map 直接用标准 key」的形态。
FIELD_CANDIDATES: Dict[str, List[str]] = {
    "name": ["姓名", "name"],
    "phone": ["电话", "手机号", "手机", "phone"],
    "email": ["邮箱", "email"],
    "intention_job": ["求职意向", "job_title", "intention_job"],
    "political_status": ["政治面貌", "political_status"],
    "school": ["毕业院校", "学校", "school"],
    "education_degree": ["学历", "education_degree"],
    "city": ["所在地", "城市", "现居", "现居城市", "location", "city"],
    "hobby": ["兴趣爱好", "爱好", "hobby"],
    "address": ["现居地址", "地址", "详细地址", "address"],
    "age": ["年龄", "age"],
    "gender": ["性别", "gender"],
    "nation": ["民族", "nation"],
    "birth_date": ["出生年月", "出生日期", "birth_date"],
    "height": ["身高", "height"],
    "hometown": ["籍贯", "hometown"],
    "major": ["专业", "major"],
    "salary": ["期望薪资", "salary"],
    "arrive_time": ["到岗时间", "arrive_time"],
    "gpa": ["GPA", "绩点", "gpa"],
    "wechat": ["微信", "wechat"],
}

# 「联系方式」是整段联系方式（电话+邮箱），与「电话」是不同内容项，单独成概念
CONTACT_CANDIDATES: List[str] = ["联系方式", "contact"]
CONTACT_FALLBACK: List[str] = ["电话", "手机号", "phone"]
CONTACT_LITERALS = ("联系方式", "联系信息", "contact")

# 英文标准键 → 概念（main.py 已把这些英文键塞进 content_map）
EN_CONCEPTS: Dict[str, tuple] = {
    "name": ("field", "name"),
    "phone": ("field", "phone"),
    "email": ("field", "email"),
    "job_title": ("field", "intention_job"),
    "location": ("field", "city"),
    "political_status": ("field", "political_status"),
    "skills": ("section", "skill_info"),
    "summary": ("section", "self_evaluate"),
    "education": ("section", "education_info"),
    "work_experience": ("section", "work_history"),
    "projects": ("section", "projects"),
    "contact": ("contact", "contact"),
}

# 概念标准中文键 → 类别（供提示文案区分「板块」与「字段」）
_SECTION_LABELS = {cands[0] for cands in SECTION_CANDIDATES.values()}
_FIELD_LABELS = {cands[0] for cands in FIELD_CANDIDATES.values()}

# 照片类占位符（不是内容项，由插入图片的逻辑单独处理，不计入未填充）
PHOTO_LITERALS = ("照片", "头像", "photo", "avatar", "image", "picture")


def clean_label(label: str) -> str:
    """清洗原始标签：全角空格转半角、连续空白压成一个、去首尾空白。"""
    text = str(label or "").replace("\u3000", " ")
    return re.sub(r"\s+", " ", text).strip()


def squash(text: str) -> str:
    """去掉全部空白（宽松比较用，避免换行/多 run 打散导致误判）。"""
    return re.sub(r"\s+", "", str(text or ""))


def extract_labels(text: str) -> List[str]:
    """从一段文本里提取全部占位符标签（已清洗去空白）。"""
    return [clean_label(m.group(1)) for m in PLACEHOLDER_RE.finditer(text or "")]


def is_contact_label(label: str) -> bool:
    """标签是否表示「整段联系方式」（需要按上下文拆成电话/邮箱）。"""
    return squash(label) in CONTACT_LITERALS


def is_photo_label(label: str) -> bool:
    """标签是否是照片占位符（{{照片}} / {{photo}} / {{avatar}}）。"""
    t = squash(label).lower()
    return any(squash(p).lower() in t for p in PHOTO_LITERALS)


def key_kind(key: str) -> str:
    """标准键类别：section / field / other。"""
    if key in _SECTION_LABELS:
        return "section"
    if key in _FIELD_LABELS:
        return "field"
    return "other"


def _cm_lookup(candidates, content_map: Optional[Dict[str, str]]) -> Optional[str]:
    """按候选顺序在 content_map 里找第一个存在的键（中文优先，英文大小写不敏感）。"""
    if not content_map:
        return None
    for cand in candidates:
        if cand in content_map:
            return cand
    ascii_map = {k.lower(): k for k in content_map
                 if isinstance(k, str) and k.isascii()}
    for cand in candidates:
        hit = ascii_map.get(cand.lower())
        if hit:
            return hit
    return None


def _cm_lookup_loose(raw: str, content_map: Optional[Dict[str, str]]) -> Optional[str]:
    """忽略空白差异的精确匹配（cm 键带空格/全角空格时也能命中）。"""
    if not content_map:
        return None
    target = squash(raw)
    if not target:
        return None
    for key in content_map:
        if squash(key) == target:
            return key
    return None


def _concept_candidates(kind: str, concept: str) -> List[str]:
    if kind == "section":
        return list(SECTION_CANDIDATES.get(concept, []))
    if kind == "field":
        return list(FIELD_CANDIDATES.get(concept, []))
    if kind == "contact":
        return list(CONTACT_CANDIDATES)
    return []


def concept_of(label: str) -> Optional[Tuple[str, str]]:
    """识别原始标签的概念：(类别, 概念键)。类别为 section / field / contact。"""
    raw = clean_label(label)
    if not raw:
        return None
    en = EN_CONCEPTS.get(raw.lower())
    if en:
        return en
    if squash(raw) in CONTACT_LITERALS:
        return ("contact", "contact")
    sec = standard_keys.find_section_key(raw)
    if sec:
        return ("section", sec)
    for key, aliases in EXTRA_SECTION_ALIASES.items():
        if squash(raw) in {squash(a) for a in aliases}:
            return ("section", key)
    fld = standard_keys.find_field_key(raw)
    if fld:
        return ("field", fld)
    return None


def canonical_key(label: str) -> Optional[str]:
    """标签的概念标准键（如「工作经验」→「工作经历」）；完全未知返回 None。"""
    concept = concept_of(label)
    if not concept:
        return None
    kind, ckey = concept
    cands = _concept_candidates(kind, ckey)
    return cands[0] if cands else clean_label(label)


def _reason(raw: str, kind: str) -> str:
    if raw.lower() in EN_CONCEPTS:
        return "英文标准键"
    if kind == "section":
        return "板块别名"
    if kind == "field":
        return "字段别名"
    return "联系方式别名"


def resolve_label(label: str, content_map: Optional[Dict[str, str]] = None,
                  concept: Optional[Tuple[str, str]] = None) -> Resolution:
    """把模板里的原始标签归一成 content_map 的标准键。

    content_map 为 None 时只做概念归一（key 取概念标准键）；
    传入 content_map 时 key 为 content_map 里实际命中的键。
    concept：调用方已知的概念（如字段行识别出的 address），优先于自动识别，
    避免同名标签在别的概念下被误解。
    """
    raw = clean_label(label)
    if not raw:
        return Resolution(raw, None, "", "unknown", "空标签", ())

    # ① 概念识别（英文标准键 / 联系方式 / 板块别名 / 字段别名）
    #    调用方已知概念时（concept 参数）优先用调用方给的概念，
    #    避免同名标签在别的概念下被误解（如「现居地址」既是 city 又是 address 别名）。
    concept = concept or concept_of(raw)
    if concept:
        kind, ckey = concept
        cands = tuple(_concept_candidates(kind, ckey))
        canonical = cands[0] if cands else raw
        if content_map is None:
            return Resolution(raw, canonical, canonical, "alias", _reason(raw, kind), cands)
        # ② 原始标签精确命中 content_map（模板后台配置的标签必然在 content_map 里）
        #    英文标准键除外：{{name}} 统一归一到中文标准键「姓名」，
        #    免得同一个内容出现两种键名（content_map 里中文/英文键值可能不同源）。
        if raw.lower() not in EN_CONCEPTS:
            hit = _cm_lookup([raw], content_map)
            if hit is not None:
                return Resolution(raw, hit, canonical, "exact",
                                  "原始标签精确命中", (hit,))
            loose = _cm_lookup_loose(raw, content_map)
            if loose is not None:
                return Resolution(raw, loose, canonical, "exact",
                                  "原始标签命中（忽略空白）", (loose,))
        # ③ 概念候选键命中（先中文后英文）
        hit = _cm_lookup(cands, content_map)
        if hit is not None:
            return Resolution(raw, hit, canonical, "alias", _reason(raw, kind), cands)
        if kind == "contact":
            # content_map 里没有「联系方式」时退到电话（避免整段联系方式丢失）
            hit = _cm_lookup(CONTACT_FALLBACK, content_map)
            if hit is not None:
                return Resolution(raw, hit, canonical, "alias",
                                  "联系方式别名（回退到电话）", cands + (hit,))
        # 概念已知但 content_map 里没有对应内容项
        return Resolution(raw, None, canonical, "missing",
                          "content_map 缺少对应内容项", cands)

    # ④ 完全未知的标签：只有 content_map 里原样存在才算命中
    #    （模板后台自定义的标签走这条路）
    if content_map is not None:
        hit = _cm_lookup([raw], content_map)
        if hit is not None:
            return Resolution(raw, hit, raw, "exact", "原始标签精确命中", (hit,))
        loose = _cm_lookup_loose(raw, content_map)
        if loose is not None:
            return Resolution(raw, loose, raw, "exact",
                              "原始标签命中（忽略空白）", (loose,))
    return Resolution(raw, None, "", "unknown", "未知标签", ())


def normalize_labels(text: str, content_map: Optional[Dict[str, str]] = None) -> List[str]:
    """把一段文本里的占位符标签逐个归一，返回标准键列表（未命中的保留原标签）。"""
    out = []
    for raw in extract_labels(text):
        res = resolve_label(raw, content_map)
        out.append(res.key or res.canonical or res.raw)
    return out
