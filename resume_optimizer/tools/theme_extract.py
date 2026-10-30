# -*- coding: utf-8 -*-
"""从 .docx 提取配色，生成 v3 编辑器预设建议。

用法：
    D:/python/python.exe tools/theme_extract.py 文件.docx [更多.docx ...]

输出：每个文件的主题色（theme1.xml 的 accent1-6）与正文高频颜色，
以及一段可直接粘进 editor-v3.js PRESETS / ResumeRender.THEMES 的 JSON 建议。
只读，不修改原文件。
"""
import json
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path

# Word 内置主题色槽位
THEME_SLOTS = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]


def _hex(v: str) -> str:
    v = (v or "").strip().lstrip("#").upper()
    return v if re.fullmatch(r"[0-9A-F]{6}", v) else ""


def read_theme_colors(z: zipfile.ZipFile) -> dict:
    """读取 word/theme/theme1.xml 的 accent 色（srgbClr / sysClr lastClr）。"""
    out = {}
    try:
        xml = z.read("word/theme/theme1.xml").decode("utf-8", "ignore")
    except KeyError:
        return out
    for slot in THEME_SLOTS:
        m = re.search(r'<a:%s>.*?<a:(?:srgbClr val="([0-9A-Fa-f]{6})"|sysClr[^>]*lastClr="([0-9A-Fa-f]{6})")' % slot, xml, re.S)
        if m:
            out[slot] = _hex(m.group(1) or m.group(2))
    return out


def count_doc_colors(z: zipfile.ZipFile, limit_parts=8) -> Counter:
    """统计正文里出现频率最高的文字颜色与底纹色（排除黑白灰）。"""
    cnt = Counter()
    names = [n for n in z.namelist() if n.startswith("word/") and n.endswith(".xml")]
    for name in names[:limit_parts]:
        try:
            xml = z.read(name).decode("utf-8", "ignore")
        except Exception:  # noqa: BLE001
            continue
        for m in re.finditer(r'<w:color[^>]*w:val="([0-9A-Fa-f]{6})"', xml):
            cnt[_hex(m.group(1))] += 1
        for m in re.finditer(r'<w:shd[^>]*w:fill="([0-9A-Fa-f]{6})"', xml):
            cnt[_hex(m.group(1))] += 1
    for skip in ("", "000000", "FFFFFF", "AUTO", "F2F2F2", "EEEEEE", "cccccc".upper()):
        cnt.pop(skip, None)
    return cnt


def _pick_design_color(colors: Counter):
    """从高频色里挑「像设计主色」的：有饱和度、不过亮/过暗，按出现次数×饱和度打分。"""
    import colorsys
    best, best_score = "", 0.0
    for c, n in colors.items():
        try:
            r, g, b = int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)
        except ValueError:
            continue
        _, light, sat = colorsys.rgb_to_hls(r / 255, g / 255, b / 255)
        if sat < 0.18 or light < 0.12 or light > 0.9:
            continue
        score = n * (0.4 + sat)
        if score > best_score:
            best, best_score = c, score
    return best


def suggest(path: Path) -> dict:
    with zipfile.ZipFile(path) as z:
        theme = read_theme_colors(z)
        colors = count_doc_colors(z)
    top = [c for c, _ in colors.most_common(5)]
    design = _pick_design_color(colors)
    primary = design or theme.get("accent1") or (top[0] if top else "34549B")
    return {
        "file": path.name,
        "theme_accents": theme,
        "top_colors": [{"color": "#" + c, "count": n} for c, n in colors.most_common(5)],
        "preset_suggestion": {
            "name": path.stem[:12],
            "family": "single",          # 版式需人工确认（单栏/侧栏/横幅/双栏）
            "primary": "#" + primary,
            "theme": "custom",           # 现有 THEMES 无对应色时用 custom + primary
        },
    }


def main(argv):
    files = [Path(a) for a in argv[1:] if Path(a).suffix.lower() == ".docx"]
    if not files:
        print(__doc__)
        return 1
    for f in files:
        if not f.exists():
            print(json.dumps({"file": str(f), "error": "文件不存在"}, ensure_ascii=False))
            continue
        try:
            print(json.dumps(suggest(f), ensure_ascii=False, indent=2))
        except Exception as exc:  # noqa: BLE001
            print(json.dumps({"file": str(f), "error": str(exc)[:200]}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
