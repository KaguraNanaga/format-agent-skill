"""把 FormatSpec 转换成目标 DOCX 内真正的 Word 命名段落样式。

样式集负责外观，RoleMap 只负责决定每个段落绑定哪个样式。这样 Word 的样式窗格、
导航窗格和目录功能都能识别文档结构，后续修改一个样式也会联动全部对应段落。
"""

import hashlib
import re
from collections import Counter

from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

from core.numbering import (
    clear_style_numbering,
    ensure_numbering_groups,
    set_style_numbering,
)


ROLE_STYLE_NAMES = {
    "title": "文档标题",
    "subtitle": "文档副标题",
    "heading_1": "标题 1",
    "heading_2": "标题 2",
    "heading_3": "标题 3",
    "body": "格式正文",
    "signature": "落款",
    "date": "日期",
    "attachment_label": "附件标题",
    "attachment": "附件正文",
    "figure_caption": "图题",
    "table_caption": "表题",
    "other": "其他正文",
}

ROLE_STYLE_IDS = {
    "title": "FormatAgentTitle",
    "subtitle": "FormatAgentSubtitle",
    "heading_1": "FormatAgentHeading1",
    "heading_2": "FormatAgentHeading2",
    "heading_3": "FormatAgentHeading3",
    "body": "FormatAgentBody",
    "signature": "FormatAgentSignature",
    "date": "FormatAgentDate",
    "attachment_label": "FormatAgentAttachmentLabel",
    "attachment": "FormatAgentAttachment",
    "figure_caption": "FormatAgentFigureCaption",
    "table_caption": "FormatAgentTableCaption",
    "other": "FormatAgentOther",
}

# 按 Word 标准标题层级：Heading 1/2/3 对应 outlineLvl 0/1/2。
# 文档主标题也保留在顶层导航，便于从导航窗格回到文首。
DEFAULT_OUTLINE_LEVELS = {
    "title": 0,
    "heading_1": 0,
    "heading_2": 1,
    "heading_3": 2,
}

_ALIGNMENT = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def style_name_for_role(role, rule=None):
    """返回角色对应的 Word 样式名；FormatSpec 可用 style_name 显式覆盖。"""
    rule = rule or {}
    explicit = rule.get("style_name")
    if isinstance(explicit, str) and explicit.strip():
        return explicit.strip()
    if role in ROLE_STYLE_NAMES:
        return ROLE_STYLE_NAMES[role]
    readable = re.sub(r"\s+", " ", str(role).replace("_", " ")).strip() or "Custom"
    return "格式代理-" + readable


def style_id_for_role(role):
    if role in ROLE_STYLE_IDS:
        return ROLE_STYLE_IDS[role]
    digest = hashlib.sha1(str(role).encode("utf-8")).hexdigest()[:10]
    return "FormatAgentCustom" + digest


def _style_by_id(document, style_id):
    for style in document.styles:
        if style.style_id == style_id:
            return style
    return None


def _role_for_index(rolemap, idx):
    """兼容内存中的整数键和 JSON 直接读入后的字符串键。"""
    if not isinstance(rolemap, dict):
        return None
    return rolemap.get(idx, rolemap.get(str(idx)))


def _default_paragraph_style(document):
    """返回目标文档声明的默认段落样式；找不到时再退到 Normal。"""
    for style in document.styles:
        if (
            style.type == WD_STYLE_TYPE.PARAGRAPH
            and style.element.get(qn("w:default")) in {"1", "true", "on"}
        ):
            return style
    normal = _style_by_id(document, "Normal")
    if normal is not None and normal.type == WD_STYLE_TYPE.PARAGRAPH:
        return normal
    try:
        normal = document.styles["Normal"]
    except KeyError:
        normal = None
    if normal is not None and normal.type == WD_STYLE_TYPE.PARAGRAPH:
        return normal
    raise ValueError("目标文档没有可用的默认正文段落样式")


def resolve_target_body_style(document, rolemap):
    """解析排版前目标文档自己的正文样式。

    统计 RoleMap 中 body 段落的有效原样式：有显式 pStyle 时使用它，
    没有显式 pStyle 的每一段都按目标文档默认段落样式计数。这样少量带
    “Normal Indent”等特例样式的正文，不会压过大量隐式 Normal 正文。
    并列时取在文档中最先出现的样式，保证结果稳定。
    """
    counts = Counter()
    first_seen = {}
    default_style = _default_paragraph_style(document)
    for idx, paragraph in enumerate(document.paragraphs):
        if _role_for_index(rolemap, idx) != "body":
            continue
        ppr = paragraph._p.pPr
        pstyle = ppr.find(qn("w:pStyle")) if ppr is not None else None
        style_id = pstyle.get(qn("w:val")) if pstyle is not None else None
        style = _style_by_id(document, style_id) if style_id else default_style
        if style is None or style.type != WD_STYLE_TYPE.PARAGRAPH:
            continue
        counts[style.style_id] += 1
        first_seen.setdefault(style.style_id, idx)

    if counts:
        style_id = min(counts, key=lambda sid: (-counts[sid], first_seen[sid]))
        return _style_by_id(document, style_id)
    return _default_paragraph_style(document)


def _get_or_create_paragraph_style(document, name, style_id):
    style = _style_by_id(document, style_id)
    if style is None:
        # 若已有同名自定义样式，沿用其 element 并修正为稳定 ID；这样不会生成
        # “标题 1 (2)”之类的重复样式，也不会再次出现 UserStyle_1。
        same_name = next((s for s in document.styles if s.name == name), None)
        if same_name is not None and same_name.type == WD_STYLE_TYPE.PARAGRAPH:
            style = same_name
        else:
            style = document.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH, builtin=False)
            # 只给我们刚创建的样式设置稳定 ID。用户已有的同名样式可能被
            # pStyle/basedOn/next/link 引用，不能原地改 ID 造成悬空引用。
            style.element.set(qn("w:styleId"), style_id)
    if style.type != WD_STYLE_TYPE.PARAGRAPH:
        raise ValueError(f"样式 {name!r} 已存在，但不是段落样式")
    name_el = style.element.find(qn("w:name"))
    if name_el is None:
        name_el = OxmlElement("w:name")
        style.element.insert(0, name_el)
    name_el.set(qn("w:val"), name)
    style.element.set(qn("w:customStyle"), "1")
    style.hidden = False
    style.quick_style = True
    return style


def _reset_style_format(style):
    """清除样式自带的格式，避免 Word 内置主题色/默认间距污染 FormatSpec。"""
    for tag in ("w:pPr", "w:rPr"):
        el = style.element.find(qn(tag))
        if el is not None:
            style.element.remove(el)


def _get_or_add(parent, tag):
    el = parent.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        parent.append(el)
    return el


def _set_style_font(style, rule):
    rpr = style.element.get_or_add_rPr()
    rfonts = _get_or_add(rpr, "w:rFonts")
    eastasia = rule.get("font_eastasia")
    ascii_font = rule.get("font_ascii") or eastasia
    if eastasia:
        rfonts.set(qn("w:eastAsia"), eastasia)
    if ascii_font:
        rfonts.set(qn("w:ascii"), ascii_font)
        rfonts.set(qn("w:hAnsi"), ascii_font)
        rfonts.set(qn("w:cs"), ascii_font)

    size_pt = rule.get("size_pt")
    if size_pt is not None:
        style.font.size = Pt(float(size_pt))
        # style.font.size 只保证 w:sz；复杂文字字号 w:szCs 也显式写入。
        sz_cs = _get_or_add(rpr, "w:szCs")
        sz_cs.set(qn("w:val"), str(int(round(float(size_pt) * 2))))
    if rule.get("bold") is not None:
        style.font.bold = bool(rule["bold"])


def _set_style_paragraph_format(style, rule, outline_level):
    pf = style.paragraph_format
    alignment = rule.get("alignment")
    if alignment in _ALIGNMENT:
        pf.alignment = _ALIGNMENT[alignment]

    line_spacing = rule.get("line_spacing")
    if isinstance(line_spacing, dict) and line_spacing.get("pt") is not None:
        value = float(line_spacing["pt"])
        pf.line_spacing = Pt(value) if line_spacing.get("type") == "exact" else value
    if rule.get("space_before_pt") is not None:
        pf.space_before = Pt(float(rule["space_before_pt"]))
    if rule.get("space_after_pt") is not None:
        pf.space_after = Pt(float(rule["space_after_pt"]))

    ppr = style.element.get_or_add_pPr()
    chars = rule.get("first_line_indent_chars")
    if chars is not None:
        ind = ppr.get_or_add_ind()
        ind.set(qn("w:firstLineChars"), str(int(round(float(chars) * 100))))
        size_pt = float(rule.get("size_pt") or 16)
        ind.set(qn("w:firstLine"), str(int(round(size_pt * float(chars) * 20))))

    outline = ppr.get_or_add_outlineLvl() if outline_level is not None else None
    if outline is not None:
        outline.set(qn("w:val"), str(int(outline_level)))


def ensure_role_styles(document, spec, target_body_style=None):
    """在目标文档中创建/更新 FormatSpec 的全部命名样式，返回 {role: style}。"""
    result = {}
    used_names = {}
    target_body_style = target_body_style or _default_paragraph_style(document)
    if target_body_style.type != WD_STYLE_TYPE.PARAGRAPH:
        raise ValueError("目标正文样式必须是段落样式")
    roles = spec.get("roles") or {}
    for role, rule in roles.items():
        name = style_name_for_role(role, rule)
        if name in used_names:
            raise ValueError(
                f"角色 {used_names[name]!r} 和 {role!r} 使用了同一个 Word 样式名 {name!r}")
        used_names[name] = role
        style = _get_or_create_paragraph_style(document, name, style_id_for_role(role))
        _reset_style_format(style)
        # Word UI 中“样式基于：无样式”对应 OOXML 完全没有 w:basedOn。
        # setter 会删除关系；再做一次低层兜底以兼容异常旧文档。
        style.base_style = None
        based_on = style.element.find(qn("w:basedOn"))
        if based_on is not None:
            style.element.remove(based_on)
        outline_level = rule.get("outline_level", DEFAULT_OUTLINE_LEVELS.get(role))
        _set_style_font(style, rule)
        _set_style_paragraph_format(style, rule, outline_level)
        clear_style_numbering(style)
        result[role] = style

    numbering_by_role = ensure_numbering_groups(
        document, roles, {role: style.style_id for role, style in result.items()})
    for role, style in result.items():
        if role in numbering_by_role:
            num_id, level = numbering_by_role[role]
            set_style_numbering(style, num_id, level)

    # 所有由 FormatAgent 管理的样式按回车后都回到目标文档自己的正文样式。
    # 这里引用目标 styleId，不复制或改写目标正文样式本身。
    for style in result.values():
        style.next_paragraph_style = target_body_style
    return result


def _remove_if_empty(parent, child):
    if child is not None and not child.attrib and len(child) == 0:
        parent.remove(child)


def _clear_rpr_controlled_fields(rpr, controlled, linked_style_ids=None):
    if rpr is None:
        return
    tags = list(controlled)
    for tag in tags:
        el = rpr.find(tag)
        if el is not None:
            rpr.remove(el)
    rstyle = rpr.find(qn("w:rStyle"))
    if (
        rstyle is not None
        and linked_style_ids
        and rstyle.get(qn("w:val")) in linked_style_ids
    ):
        rpr.remove(rstyle)


def _clear_run_overrides(paragraph, rule, clear_character_style=False):
    controlled = []
    if rule.get("font_eastasia") or rule.get("font_ascii"):
        controlled.append(qn("w:rFonts"))
    if rule.get("size_pt") is not None:
        controlled.extend((qn("w:sz"), qn("w:szCs")))
    if rule.get("bold") is not None:
        controlled.extend((qn("w:b"), qn("w:bCs")))
    linked_style_ids = set()
    if clear_character_style:
        for style in paragraph.part.document.styles:
            if (
                style.type == WD_STYLE_TYPE.CHARACTER
                and style.element.find(qn("w:link")) is not None
            ):
                linked_style_ids.add(style.style_id)

    for run in paragraph.runs:
        rpr = run._element.find(qn("w:rPr"))
        if rpr is None:
            continue
        _clear_rpr_controlled_fields(rpr, controlled, linked_style_ids)
        _remove_if_empty(run._element, rpr)

    # Word 还允许在段落标记 pPr/rPr 上保存字符样式。旧文件中的
    # NormalCharacter -> UserStyle_1 链接正是由这里泄漏到样式对话框的。
    ppr = paragraph._p.pPr
    mark_rpr = ppr.find(qn("w:rPr")) if ppr is not None else None
    if mark_rpr is not None:
        _clear_rpr_controlled_fields(mark_rpr, controlled, linked_style_ids)
        _remove_if_empty(ppr, mark_rpr)


def _clear_paragraph_overrides(paragraph, rule, remove_numbering=False):
    ppr = paragraph._p.get_or_add_pPr()
    if rule.get("alignment") in _ALIGNMENT:
        jc = ppr.find(qn("w:jc"))
        if jc is not None:
            ppr.remove(jc)

    spacing = ppr.find(qn("w:spacing"))
    if spacing is not None:
        if isinstance(rule.get("line_spacing"), dict):
            spacing.attrib.pop(qn("w:line"), None)
            spacing.attrib.pop(qn("w:lineRule"), None)
        if rule.get("space_before_pt") is not None:
            spacing.attrib.pop(qn("w:before"), None)
        if rule.get("space_after_pt") is not None:
            spacing.attrib.pop(qn("w:after"), None)
        _remove_if_empty(ppr, spacing)

    if rule.get("first_line_indent_chars") is not None:
        ind = ppr.find(qn("w:ind"))
        if ind is not None:
            ind.attrib.pop(qn("w:firstLineChars"), None)
            ind.attrib.pop(qn("w:firstLine"), None)
            ind.attrib.pop(qn("w:hanging"), None)
            ind.attrib.pop(qn("w:hangingChars"), None)
            _remove_if_empty(ppr, ind)

    # 大纲层级由命名样式提供，段落本身不再直刷 outlineLvl。
    outline = ppr.find(qn("w:outlineLvl"))
    if outline is not None:
        ppr.remove(outline)

    if remove_numbering:
        num_pr = ppr.find(qn("w:numPr"))
        if num_pr is not None:
            ppr.remove(num_pr)


def clear_invalid_numbering_override(paragraph):
    """删除明确表示“取消编号”的段落级 numPr，保留真实自动编号。

    Word/WPS 会用 numId=0 或 ilvl=-1 表示从编号列表退出。这个残留即使不
    显示编号，也会遮蔽正文样式的首行缩进。缺少字段或无法解析的 numPr 不做
    猜测；numId>0 且 ilvl>=0 的真实自动编号始终保留。
    """
    ppr = paragraph._p.pPr
    num_pr = ppr.find(qn("w:numPr")) if ppr is not None else None
    if num_pr is None:
        return False

    def _value(tag):
        element = num_pr.find(qn(tag))
        if element is None:
            return None
        try:
            return int(element.get(qn("w:val")))
        except (TypeError, ValueError):
            return None

    num_id = _value("w:numId")
    level = _value("w:ilvl")
    if (num_id is not None and num_id <= 0) or (level is not None and level < 0):
        ppr.remove(num_pr)
        return True
    return False


_MANUAL_PREFIXES = {
    "heading_1": re.compile(r"^\s*[一二三四五六七八九十百〇零两]+[、.．]\s*"),
    "heading_2": re.compile(r"^\s*[（(][一二三四五六七八九十百〇零两]+[）)]\s*"),
    "heading_3": re.compile(r"^\s*\d+[、.．]\s*"),
}


def _strip_manual_number_prefix(paragraph, role):
    """编号样式启用时去掉手工键入的前缀，避免“一、一、标题”。"""
    pattern = _MANUAL_PREFIXES.get(role)
    match = pattern.match(paragraph.text) if pattern is not None else None
    if match is None:
        return False
    prefix = match.group(0)
    for run in paragraph.runs:
        text = run.text
        payload = [child for child in run._element if child.tag != qn("w:rPr")]
        if not text:
            # 若编号前有字段、绘图或制表符，跳过清理，绝不破坏复杂 run。
            if payload:
                return False
            continue
        if not text.startswith(prefix) or any(child.tag != qn("w:t") for child in payload):
            return False
        remaining = len(prefix)
        for text_node in payload:
            value = text_node.text or ""
            if remaining >= len(value):
                text_node.text = ""
                remaining -= len(value)
            else:
                text_node.text = value[remaining:]
                remaining = 0
                break
        return remaining == 0
    return False


def apply_named_style(paragraph, style, rule, role=None):
    """绑定命名样式并清除会遮蔽该样式的直接格式，返回受控字段列表。"""
    clear_character_style = role in {
        "title", "subtitle", "heading_1", "heading_2", "heading_3"
    }
    _clear_run_overrides(
        paragraph, rule, clear_character_style=clear_character_style)
    invalid_numbering_removed = (
        clear_invalid_numbering_override(paragraph) if role == "body" else False)
    has_numbering = isinstance(rule.get("numbering"), dict)
    stripped_prefix = _strip_manual_number_prefix(paragraph, role) if has_numbering else False
    remove_numbering = has_numbering or role in {
        "title", "subtitle", "heading_1", "heading_2", "heading_3"
    }
    _clear_paragraph_overrides(paragraph, rule, remove_numbering=remove_numbering)
    paragraph.style = style
    fields = ["paragraph_style"]
    if has_numbering:
        fields.append("automatic_numbering")
    if stripped_prefix:
        fields.append("manual_number_prefix_removed")
    if invalid_numbering_removed:
        fields.append("invalid_numbering_removed")
    fields.extend(
        field for field in (
            "font_eastasia", "font_ascii", "size_pt", "bold", "alignment",
            "line_spacing", "space_before_pt", "space_after_pt",
            "first_line_indent_chars", "outline_level",
        )
        if rule.get(field) is not None
    )
    return fields
