# 读段落"生效属性"——合并 样式继承链(basedOn) -> 段落样式 -> run 直刷格式。
# 这是"读模板得数值"的核心, 也是方案第6节"读生效属性而非样式名"的具体实现。
# 用法: get_paragraph_effective_font(paragraph) -> ("仿宋_GB2312", 16.0, False)
# 字符属性按 docDefaults -> 样式继承链 -> 段落中占主导的 run 合并；
# 主题字体会解析为实际字体名。段落对齐/行距/缩进由 rules_from_template 读取。
from collections import defaultdict
from xml.etree import ElementTree as ET

from docx.oxml.ns import qn

_OFF_VALUES = ("0", "false", "off", "none")


def _is_on(el):
    """w:b 这类开关属性: 元素存在且 val 不是 0/false/off 才算开。"""
    return el is not None and el.get(qn("w:val")) not in _OFF_VALUES


def _merge(base, direct):
    """两层叠加, direct 优先, None 不覆盖。"""
    merged = dict(base)
    for k, v in direct.items():
        if v is not None:
            merged[k] = v
    return merged


_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _theme_fonts(paragraph):
    """返回 major/minor Latin/EastAsia/CS 的实际字体；主题缺失时为空。"""
    try:
        part = next(
            value for value in paragraph.part.package.parts
            if str(value.partname).startswith("/word/theme/theme")
            and str(value.partname).endswith(".xml")
        )
        root = ET.fromstring(part.blob)
    except (AttributeError, StopIteration, ET.ParseError):
        return {}
    result = {}
    ns = {"a": _A_NS}
    for family, node_name in (("major", "majorFont"), ("minor", "minorFont")):
        node = root.find(f".//a:fontScheme/a:{node_name}", ns)
        if node is None:
            continue
        for key, tag in (("HAnsi", "latin"), ("EastAsia", "ea"), ("Bidi", "cs")):
            value = node.find(f"a:{tag}", ns)
            if value is not None and value.get("typeface"):
                result[f"{family}{key}"] = value.get("typeface")
    return result


def _read_rpr(rpr, theme_fonts=None):
    """从一个 rPr 元素读常用字符属性, 返回 dict (只含有值项)。
    注意: w:b 元素不存在时不写 bold 键——"没说"不等于"不加粗", 要留给下层决定。
    """
    props = {}
    if rpr is None:
        return props
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is not None:
        theme_fonts = theme_fonts or {}
        eastasia = rfonts.get(qn("w:eastAsia"))
        ascii_font = rfonts.get(qn("w:ascii")) or rfonts.get(qn("w:hAnsi"))
        complex_font = rfonts.get(qn("w:cs"))
        if not eastasia and rfonts.get(qn("w:eastAsiaTheme")):
            eastasia = theme_fonts.get(rfonts.get(qn("w:eastAsiaTheme")))
        if not ascii_font:
            theme_key = rfonts.get(qn("w:asciiTheme")) or rfonts.get(qn("w:hAnsiTheme"))
            ascii_font = theme_fonts.get(theme_key)
        if not complex_font and rfonts.get(qn("w:cstheme")):
            complex_font = theme_fonts.get(rfonts.get(qn("w:cstheme")))
        if eastasia:
            props["eastasia"] = eastasia
        if ascii_font:
            props["ascii"] = ascii_font
        if complex_font:
            props["cs"] = complex_font
    sz = rpr.find(qn("w:sz"))
    if sz is not None and sz.get(qn("w:val")):
        props["size_pt"] = int(sz.get(qn("w:val"))) / 2.0
    b = rpr.find(qn("w:b"))
    if b is not None:
        props["bold"] = _is_on(b)  # 显式 w:b w:val="0" 才是"取消加粗"
    for tag, key in (("w:i", "italic"), ("w:strike", "strike")):
        element = rpr.find(qn(tag))
        if element is not None:
            props[key] = _is_on(element)
    underline = rpr.find(qn("w:u"))
    if underline is not None:
        props["underline"] = underline.get(qn("w:val"), "single") not in _OFF_VALUES
    color = rpr.find(qn("w:color"))
    if color is not None:
        value = color.get(qn("w:val"))
        if value and value.lower() != "auto":
            props["color"] = value.upper()
    highlight = rpr.find(qn("w:highlight"))
    if highlight is not None and highlight.get(qn("w:val")):
        props["highlight"] = highlight.get(qn("w:val"))
    for tag, key in (("w:caps", "caps"), ("w:smallCaps", "small_caps"),
                     ("w:rtl", "rtl")):
        element = rpr.find(qn(tag))
        if element is not None:
            props[key] = _is_on(element)
    language = rpr.find(qn("w:lang"))
    if language is not None:
        value = language.get(qn("w:val")) or language.get(qn("w:bidi"))
        if value:
            props["language"] = value
    return props


def _style_chain_props(style, theme_fonts=None):
    """沿 basedOn 继承链自底向上合并样式 rPr（先父后子，子覆盖父）。"""
    chain = []
    cur = style
    while cur is not None and getattr(cur, "element", None) is not None:
        if any(cur.element is s.element for s in chain):
            break  # basedOn 成环保护
        chain.append(cur)
        cur = cur.base_style
    result = {}
    for s in reversed(chain):
        result = _merge(
            result, _read_rpr(s.element.find(qn("w:rPr")), theme_fonts))
    return result


def _doc_defaults(paragraph, theme_fonts=None):
    try:
        styles = paragraph.part.document.styles.element
    except AttributeError:
        return {}
    defaults = styles.find(qn("w:docDefaults"))
    rpr_default = defaults.find(qn("w:rPrDefault")) if defaults is not None else None
    rpr = rpr_default.find(qn("w:rPr")) if rpr_default is not None else None
    return _read_rpr(rpr, theme_fonts)


def _dominant_run_props(paragraph, base, theme_fonts):
    """按非空字符数选择各属性的主导值，避免首个短标签 Run 污染规则。

    “未设置”也必须参与投票。例如 ``Label:`` 加粗而后面长正文没有加粗时，
    不能因为只有标签声明了 ``w:b`` 就把整个角色模板判为粗体。
    """
    missing = object()
    resolved_runs = []
    keys = set(base)
    for run in paragraph.runs:
        direct = _read_rpr(run._element.find(qn("w:rPr")), theme_fonts)
        resolved = _merge(base, direct)
        resolved_runs.append((max(1, len(run.text or "")), resolved))
        keys.update(direct)

    result = dict(base)
    for key in keys:
        votes = defaultdict(int)
        for weight, resolved in resolved_runs:
            votes[resolved.get(key, missing)] += weight
        winner = max(votes.items(), key=lambda item: item[1])[0]
        if winner is missing:
            result.pop(key, None)
        else:
            result[key] = winner
    return result


def effective_props(paragraph):
    """返回文档默认值、样式链和段落主导 Run 合成的有效字符属性。"""
    theme_fonts = _theme_fonts(paragraph)
    result = _doc_defaults(paragraph, theme_fonts)
    if paragraph.style is not None:
        result = _merge(
            result, _style_chain_props(paragraph.style, theme_fonts))
    if paragraph.runs:
        result = _dominant_run_props(paragraph, result, theme_fonts)
    return result


def get_paragraph_effective_font(paragraph):
    """返回 (eastAsia字体名, 字号pt, 是否加粗)。缺省回退 宋体/五号/否。"""
    p = effective_props(paragraph)
    return p.get("eastasia") or "宋体", p.get("size_pt") or 10.5, bool(p.get("bold"))
