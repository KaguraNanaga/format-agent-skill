# 读段落"生效属性"——合并 样式继承链(basedOn) -> 段落样式 -> run 直刷格式。
# 这是"读模板得数值"的核心, 也是方案第6节"读生效属性而非样式名"的具体实现。
# 用法: get_paragraph_effective_font(paragraph) -> ("仿宋_GB2312", 16.0, False)
# 已知边界: 不读 styles.xml 的 docDefaults(文档默认字体), 不读对齐/行距/缩进。
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


def _read_rpr(rpr):
    """从一个 rPr 元素读常用字符属性, 返回 dict (只含有值项)。
    注意: w:b 元素不存在时不写 bold 键——"没说"不等于"不加粗", 要留给下层决定。
    """
    props = {}
    if rpr is None:
        return props
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is not None:
        props["eastasia"] = rfonts.get(qn("w:eastAsia"))
        props["ascii"] = rfonts.get(qn("w:ascii"))
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
    return props


def _style_chain_props(style):
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
        result = _merge(result, _read_rpr(s.element.find(qn("w:rPr"))))
    return result


def effective_props(paragraph):
    """合并 样式链 -> 首个 run 直刷格式, 返回该段落'真正生效'的格式 dict。
    模板可能样式表干净、段落直接刷格式, 也可能全走样式, 两种都要能读对。
    """
    result = {}
    if paragraph.style is not None:
        result = _style_chain_props(paragraph.style)
    if paragraph.runs:
        result = _merge(result, _read_rpr(paragraph.runs[0]._element.find(qn("w:rPr"))))
    return result


def get_paragraph_effective_font(paragraph):
    """返回 (eastAsia字体名, 字号pt, 是否加粗)。缺省回退 宋体/五号/否。"""
    p = effective_props(paragraph)
    return p.get("eastasia") or "宋体", p.get("size_pt") or 10.5, bool(p.get("bold"))
