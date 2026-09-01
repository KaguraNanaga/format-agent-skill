"""把 FormatSpec 转换成目标 DOCX 内真正的 Word 命名段落样式。

样式集负责外观，RoleMap 只负责决定每个段落绑定哪个样式。这样 Word 的样式窗格、
导航窗格和目录功能都能识别文档结构，后续修改一个样式也会联动全部对应段落。
"""

import hashlib
import re
from collections import Counter
from copy import deepcopy

from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from docx.text.run import Run

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
    "heading_4": "标题 4",
    "body": "格式正文",
    "signature": "落款",
    "date": "日期",
    "attachment_label": "附件标题",
    "attachment": "附件正文",
    "figure_caption": "图题",
    "table_caption": "表题",
    "abstract_heading": "摘要标题",
    "abstract_body": "摘要正文",
    "keywords": "关键词",
    "chapter_heading": "章标题",
    "bibliography_heading": "参考文献标题",
    "bibliography_entry": "参考文献条目",
    "equation": "公式",
    "appendix_heading": "附录标题",
    "list_of_figures_heading": "图目录标题",
    "list_of_tables_heading": "表目录标题",
    "block_quote": "块引用",
    "code_block": "代码块",
    "byline": "作者署名",
    "affiliation": "作者单位",
    "author_note": "作者注",
    "correspondence": "通信地址",
    "salutation": "称呼",
    "complimentary_close": "结尾敬语",
    "cc": "抄送",
    "enclosure": "附件说明",
    "legal_definition": "定义条款",
    "signature_block": "签署栏",
    "table_of_authorities_heading": "引证表标题",
    "recipient": "主送机关",
    "closing": "公文结束语",
    "document_number": "发文字号",
    "copy_to": "抄送机关",
    "warning_box": "警告框",
    "caution_box": "注意框",
    "note_box": "说明框",
    "tip_box": "提示框",
    "procedure_step": "操作步骤",
    "command": "命令",
    "court_caption": "法院标题",
    "case_number": "案号",
    "brief_title": "法律文书标题",
    "table_of_contents_heading": "目录标题",
    "authority_entry": "引证表条目",
    "counsel_block": "律师信息",
    "certificate_heading": "证明标题",
    "certificate_body": "证明正文",
    "other": "其他正文",
}

ROLE_STYLE_IDS = {
    "title": "FormatAgentTitle",
    "subtitle": "FormatAgentSubtitle",
    "heading_1": "FormatAgentHeading1",
    "heading_2": "FormatAgentHeading2",
    "heading_3": "FormatAgentHeading3",
    "heading_4": "FormatAgentHeading4",
    "body": "FormatAgentBody",
    "signature": "FormatAgentSignature",
    "date": "FormatAgentDate",
    "attachment_label": "FormatAgentAttachmentLabel",
    "attachment": "FormatAgentAttachment",
    "figure_caption": "FormatAgentFigureCaption",
    "table_caption": "FormatAgentTableCaption",
    "abstract_heading": "FormatAgentAbstractHeading",
    "abstract_body": "FormatAgentAbstractBody",
    "keywords": "FormatAgentKeywords",
    "chapter_heading": "FormatAgentChapterHeading",
    "bibliography_heading": "FormatAgentBibliographyHeading",
    "bibliography_entry": "FormatAgentBibliographyEntry",
    "equation": "FormatAgentEquation",
    "appendix_heading": "FormatAgentAppendixHeading",
    "list_of_figures_heading": "FormatAgentListOfFiguresHeading",
    "list_of_tables_heading": "FormatAgentListOfTablesHeading",
    "block_quote": "FormatAgentBlockQuote",
    "code_block": "FormatAgentCodeBlock",
    "byline": "FormatAgentByline",
    "affiliation": "FormatAgentAffiliation",
    "author_note": "FormatAgentAuthorNote",
    "correspondence": "FormatAgentCorrespondence",
    "salutation": "FormatAgentSalutation",
    "complimentary_close": "FormatAgentComplimentaryClose",
    "cc": "FormatAgentCc",
    "enclosure": "FormatAgentEnclosure",
    "legal_definition": "FormatAgentLegalDefinition",
    "signature_block": "FormatAgentSignatureBlock",
    "table_of_authorities_heading": "FormatAgentTableOfAuthoritiesHeading",
    "recipient": "FormatAgentRecipient",
    "closing": "FormatAgentClosing",
    "document_number": "FormatAgentDocumentNumber",
    "copy_to": "FormatAgentCopyTo",
    "warning_box": "FormatAgentWarningBox",
    "caution_box": "FormatAgentCautionBox",
    "note_box": "FormatAgentNoteBox",
    "tip_box": "FormatAgentTipBox",
    "procedure_step": "FormatAgentProcedureStep",
    "command": "FormatAgentCommand",
    "court_caption": "FormatAgentCourtCaption",
    "case_number": "FormatAgentCaseNumber",
    "brief_title": "FormatAgentBriefTitle",
    "table_of_contents_heading": "FormatAgentTableOfContentsHeading",
    "authority_entry": "FormatAgentAuthorityEntry",
    "counsel_block": "FormatAgentCounselBlock",
    "certificate_heading": "FormatAgentCertificateHeading",
    "certificate_body": "FormatAgentCertificateBody",
    "other": "FormatAgentOther",
}

# 按 Word 标准标题层级：Heading 1/2/3 对应 outlineLvl 0/1/2。
# 文档主标题也保留在顶层导航，便于从导航窗格回到文首。
DEFAULT_OUTLINE_LEVELS = {
    "title": 0,
    "heading_1": 0,
    "heading_2": 1,
    "heading_3": 2,
    "heading_4": 3,
    "chapter_heading": 0,
    "bibliography_heading": 0,
    "appendix_heading": 0,
    "list_of_figures_heading": 0,
    "list_of_tables_heading": 0,
    "table_of_authorities_heading": 0,
    "table_of_contents_heading": 0,
    "brief_title": 0,
    "certificate_heading": 0,
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


def _get_or_add_ordered(parent, tag, *successors):
    """Add an OOXML property without violating the schema child order.

    Word is tolerant of many producer quirks, but an out-of-order ``rPr`` or
    ``pPr`` can make it repair (and occasionally discard) a generated style.
    ``python-docx`` exposes ordered descriptors for most, but not all, of the
    complex-script properties we need.
    """
    el = parent.find(qn(tag))
    if el is None:
        el = OxmlElement(tag)
        parent.insert_element_before(el, *successors)
    return el


def _set_style_font(style, rule, cleanup_mode="controlled"):
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.get_or_add_rFonts()
    eastasia = rule.get("font_eastasia")
    ascii_font = rule.get("font_ascii") or eastasia
    complex_font = rule.get("font_cs") or ascii_font
    if eastasia:
        rfonts.set(qn("w:eastAsia"), eastasia)
    if ascii_font:
        rfonts.set(qn("w:ascii"), ascii_font)
        rfonts.set(qn("w:hAnsi"), ascii_font)
    if complex_font:
        rfonts.set(qn("w:cs"), complex_font)

    if rule.get("language"):
        language = _get_or_add_ordered(
            rpr, "w:lang", "w:eastAsianLayout", "w:specVanish", "w:oMath")
        language.set(qn("w:val"), rule["language"])
        if rule.get("rtl") or rule.get("bidi"):
            language.set(qn("w:bidi"), rule["language"])
    for field, getter in (
        ("caps", rpr.get_or_add_caps),
        ("small_caps", rpr.get_or_add_smallCaps),
        ("rtl", rpr.get_or_add_rtl),
    ):
        if rule.get(field) is not None:
            element = getter()
            element.set(qn("w:val"), "1" if rule[field] else "0")

    size_pt = rule.get("size_pt")
    if size_pt is not None:
        style.font.size = Pt(float(size_pt))
        # style.font.size 只保证 w:sz；复杂文字字号 w:szCs 也显式写入。
        sz_cs = _get_or_add_ordered(
            rpr, "w:szCs", "w:highlight", "w:u", "w:effect", "w:bdr",
            "w:shd", "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em",
            "w:lang", "w:eastAsianLayout", "w:specVanish", "w:oMath")
        sz_cs.set(qn("w:val"), str(int(round(float(size_pt) * 2))))
    if rule.get("bold") is not None:
        style.font.bold = bool(rule["bold"])
    elif cleanup_mode == "strict":
        style.font.bold = False
    if rule.get("italic") is not None:
        style.font.italic = bool(rule["italic"])
    elif cleanup_mode == "strict":
        style.font.italic = False
    if rule.get("underline") is not None:
        style.font.underline = bool(rule["underline"])
    elif cleanup_mode == "strict":
        style.font.underline = False
    if rule.get("strike") is not None:
        style.font.strike = bool(rule["strike"])
    elif cleanup_mode == "strict":
        style.font.strike = False
    color = rule.get("color")
    if color:
        style.font.color.rgb = RGBColor.from_string(color.upper())
    elif cleanup_mode == "strict":
        style.font.color.rgb = RGBColor(0, 0, 0)
    if rule.get("highlight") is not None:
        highlight = _get_or_add_ordered(
            rpr, "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd",
            "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang",
            "w:eastAsianLayout", "w:specVanish", "w:oMath")
        highlight.set(qn("w:val"), str(rule["highlight"]))


def _set_style_paragraph_format(style, rule, outline_level):
    pf = style.paragraph_format
    alignment = rule.get("alignment")
    if alignment in _ALIGNMENT:
        pf.alignment = _ALIGNMENT[alignment]
    ppr = style.element.get_or_add_pPr()
    border_rule = rule.get("paragraph_border")
    if isinstance(border_rule, dict):
        p_borders = _get_or_add_ordered(
            ppr, "w:pBdr", "w:shd", "w:tabs", "w:suppressAutoHyphens",
            "w:kinsoku", "w:wordWrap", "w:overflowPunct", "w:topLinePunct",
            "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd",
            "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing",
            "w:mirrorIndents", "w:suppressOverlap", "w:jc")
        for child in list(p_borders):
            p_borders.remove(child)
        sides = border_rule.get("sides") or ["top", "bottom", "left", "right"]
        for side in sides:
            element = OxmlElement(f"w:{side}")
            element.set(qn("w:val"), str(border_rule.get("style", "single")))
            element.set(qn("w:sz"), str(int(round(float(
                border_rule.get("size_pt", 0.5)) * 8))))
            element.set(qn("w:space"), str(int(round(float(
                border_rule.get("space_pt", 1))))))
            element.set(qn("w:color"), str(
                border_rule.get("color", "000000")).upper())
            p_borders.append(element)
    shading = rule.get("shading")
    if shading:
        element = _get_or_add_ordered(
            ppr, "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku",
            "w:wordWrap", "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE",
            "w:autoSpaceDN", "w:bidi", "w:adjustRightInd", "w:snapToGrid",
            "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents",
            "w:suppressOverlap", "w:jc")
        element.set(qn("w:val"), "clear")
        element.set(qn("w:color"), "auto")
        element.set(qn("w:fill"), str(shading).upper())
    if rule.get("bidi") is not None:
        bidi = _get_or_add_ordered(
            ppr, "w:bidi", "w:adjustRightInd", "w:snapToGrid", "w:spacing",
            "w:ind", "w:contextualSpacing", "w:mirrorIndents",
            "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
            "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle",
            "w:rPr", "w:sectPr", "w:pPrChange")
        bidi.set(qn("w:val"), "1" if rule["bidi"] else "0")

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
    hanging_chars = rule.get("hanging_indent_chars")
    left_chars = rule.get("left_indent_chars")
    if hanging_chars is not None:
        left_chars = hanging_chars if left_chars is None else left_chars
        ind = ppr.get_or_add_ind()
        size_pt = float(rule.get("size_pt") or 16)
        ind.set(qn("w:leftChars"), str(int(round(float(left_chars) * 100))))
        ind.set(qn("w:left"), str(int(round(size_pt * float(left_chars) * 20))))
        ind.set(qn("w:hangingChars"), str(int(round(float(hanging_chars) * 100))))
        ind.set(qn("w:hanging"), str(int(round(size_pt * float(hanging_chars) * 20))))
    elif chars is not None:
        ind = ppr.get_or_add_ind()
        ind.set(qn("w:firstLineChars"), str(int(round(float(chars) * 100))))
        size_pt = float(rule.get("size_pt") or 16)
        ind.set(qn("w:firstLine"), str(int(round(size_pt * float(chars) * 20))))
    elif left_chars is not None:
        ind = ppr.get_or_add_ind()
        size_pt = float(rule.get("size_pt") or 16)
        ind.set(qn("w:leftChars"), str(int(round(float(left_chars) * 100))))
        ind.set(qn("w:left"), str(int(round(size_pt * float(left_chars) * 20))))

    for field, attr in (
        ("keep_with_next", "keep_with_next"),
        ("keep_together", "keep_together"),
        ("page_break_before", "page_break_before"),
        ("widow_control", "widow_control"),
    ):
        if rule.get(field) is not None:
            setattr(pf, attr, bool(rule[field]))

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
    cleanup_mode = (spec.get("cleanup") or {}).get("mode", "controlled")
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
        _set_style_font(style, rule, cleanup_mode=cleanup_mode)
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


def _clear_rpr_controlled_fields(
    rpr, controlled, linked_style_ids=None, remove_character_style=False,
):
    if rpr is None:
        return
    tags = list(controlled)
    for tag in tags:
        el = rpr.find(tag)
        if el is not None:
            rpr.remove(el)
    rstyle = rpr.find(qn("w:rStyle"))
    if rstyle is not None and (
        remove_character_style
        or linked_style_ids and rstyle.get(qn("w:val")) in linked_style_ids
    ):
        rpr.remove(rstyle)


_STRICT_RPR_TAGS = {
    qn(tag) for tag in (
        "w:rFonts", "w:sz", "w:szCs", "w:b", "w:bCs", "w:i", "w:iCs",
        "w:u", "w:color", "w:highlight", "w:strike", "w:dstrike",
        "w:caps", "w:smallCaps", "w:vanish", "w:outline", "w:shadow",
        "w:emboss", "w:imprint", "w:position", "w:spacing", "w:w",
        "w:kern", "w:shd", "w:lang", "w:rtl",
    )
}


def _clear_run_overrides(
    paragraph, rule, clear_character_style=False, cleanup_mode="controlled",
):
    controlled = set()
    if rule.get("font_eastasia") or rule.get("font_ascii") or rule.get("font_cs"):
        controlled.add(qn("w:rFonts"))
    if rule.get("size_pt") is not None:
        controlled.update((qn("w:sz"), qn("w:szCs")))
    if rule.get("bold") is not None:
        controlled.update((qn("w:b"), qn("w:bCs")))
    for field, tags in (
        ("italic", ("w:i", "w:iCs")),
        ("underline", ("w:u",)),
        ("color", ("w:color",)),
        ("highlight", ("w:highlight",)),
        ("strike", ("w:strike", "w:dstrike")),
        ("caps", ("w:caps",)),
        ("small_caps", ("w:smallCaps",)),
        ("rtl", ("w:rtl",)),
        ("language", ("w:lang",)),
    ):
        if rule.get(field) is not None:
            controlled.update(qn(tag) for tag in tags)
    if cleanup_mode == "strict":
        controlled.update(_STRICT_RPR_TAGS)
    elif cleanup_mode == "preserve_emphasis":
        # 保留作者刻意使用的粗斜体、小型大写、复杂文字方向和语言标记。
        controlled.difference_update(
            {qn("w:b"), qn("w:bCs"), qn("w:i"), qn("w:iCs"),
             qn("w:caps"), qn("w:smallCaps"), qn("w:rtl"), qn("w:lang")})
        controlled.update(
            tag for tag in _STRICT_RPR_TAGS
            if tag not in {qn("w:b"), qn("w:bCs"), qn("w:i"), qn("w:iCs"),
                           qn("w:caps"), qn("w:smallCaps"), qn("w:rtl"), qn("w:lang")}
        )
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
        _clear_rpr_controlled_fields(
            rpr, controlled, linked_style_ids,
            remove_character_style=cleanup_mode == "strict")
        _remove_if_empty(run._element, rpr)

    # Word 还允许在段落标记 pPr/rPr 上保存字符样式。旧文件中的
    # NormalCharacter -> UserStyle_1 链接正是由这里泄漏到样式对话框的。
    ppr = paragraph._p.pPr
    mark_rpr = ppr.find(qn("w:rPr")) if ppr is not None else None
    if mark_rpr is not None:
        _clear_rpr_controlled_fields(
            mark_rpr, controlled, linked_style_ids,
            remove_character_style=cleanup_mode == "strict")
        _remove_if_empty(ppr, mark_rpr)


def _clear_paragraph_overrides(
    paragraph, rule, remove_numbering=False, cleanup_mode="controlled",
):
    ppr = paragraph._p.get_or_add_pPr()
    strict = cleanup_mode == "strict"
    if strict or rule.get("bidi") is not None:
        bidi = ppr.find(qn("w:bidi"))
        if bidi is not None:
            ppr.remove(bidi)
    if strict or rule.get("alignment") in _ALIGNMENT:
        jc = ppr.find(qn("w:jc"))
        if jc is not None:
            ppr.remove(jc)

    spacing = ppr.find(qn("w:spacing"))
    if spacing is not None:
        if strict or isinstance(rule.get("line_spacing"), dict):
            spacing.attrib.pop(qn("w:line"), None)
            spacing.attrib.pop(qn("w:lineRule"), None)
        if strict or rule.get("space_before_pt") is not None:
            spacing.attrib.pop(qn("w:before"), None)
        if strict or rule.get("space_after_pt") is not None:
            spacing.attrib.pop(qn("w:after"), None)
        _remove_if_empty(ppr, spacing)

    if strict or any(
        rule.get(field) is not None for field in (
            "first_line_indent_chars", "left_indent_chars", "hanging_indent_chars"
        )
    ):
        ind = ppr.find(qn("w:ind"))
        if ind is not None:
            if strict:
                ppr.remove(ind)
            else:
                for attr in (
                    "w:firstLineChars", "w:firstLine", "w:hanging",
                    "w:hangingChars", "w:left", "w:leftChars", "w:start",
                    "w:startChars",
                ):
                    ind.attrib.pop(qn(attr), None)
                _remove_if_empty(ppr, ind)

    for field, tag in (
        ("keep_with_next", "w:keepNext"),
        ("keep_together", "w:keepLines"),
        ("page_break_before", "w:pageBreakBefore"),
        ("widow_control", "w:widowControl"),
    ):
        if strict or rule.get(field) is not None:
            element = ppr.find(qn(tag))
            if element is not None:
                ppr.remove(element)

    if strict:
        for tag in ("w:contextualSpacing", "w:shd", "w:pBdr", "w:textAlignment"):
            element = ppr.find(qn(tag))
            if element is not None:
                ppr.remove(element)
    else:
        for field, tag in (("shading", "w:shd"), ("paragraph_border", "w:pBdr")):
            if rule.get(field) is not None:
                element = ppr.find(qn(tag))
                if element is not None:
                    ppr.remove(element)

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


def _split_plain_run_at(paragraph, offset):
    """在段落字符偏移处拆分纯文本 run；遇到域、换行、绘图等复杂 run 则放弃。"""
    cursor = 0
    for run in list(paragraph.runs):
        text = run.text
        end = cursor + len(text)
        if cursor < offset < end:
            payload = [child for child in run._element if child.tag != qn("w:rPr")]
            if any(child.tag != qn("w:t") for child in payload):
                return False
            clone = deepcopy(run._element)
            run._element.addnext(clone)
            suffix = Run(clone, paragraph)
            cut = offset - cursor
            run.text = text[:cut]
            suffix.text = text[cut:]
            return True
        if offset == end:
            return True
        cursor = end
    return offset == cursor


def _apply_label_prefix(paragraph, rule):
    prefix_rule = rule.get("label_prefix")
    if not isinstance(prefix_rule, dict):
        return False
    options = prefix_rule.get("text")
    options = [options] if isinstance(options, str) else list(options or [])
    prefix = next((item for item in options if paragraph.text.startswith(item)), None)
    if not prefix or not _split_plain_run_at(paragraph, len(prefix)):
        return False

    cursor = 0
    for run in paragraph.runs:
        end = cursor + len(run.text)
        in_prefix = end <= len(prefix) and end > 0
        for field in ("bold", "italic", "underline"):
            if prefix_rule.get(field) is not None:
                setattr(run.font, field, bool(prefix_rule[field]) if in_prefix else None)
        if prefix_rule.get("color") is not None:
            run.font.color.rgb = (
                RGBColor.from_string(prefix_rule["color"].upper()) if in_prefix else None)
        cursor = end
    return True


def apply_named_style(
    paragraph, style, rule, role=None, cleanup_mode="controlled",
):
    """绑定命名样式并清除会遮蔽该样式的直接格式，返回受控字段列表。"""
    clear_character_style = role in {
        "title", "subtitle", "heading_1", "heading_2", "heading_3", "heading_4",
        "abstract_heading", "chapter_heading", "bibliography_heading",
        "appendix_heading", "brief_title", "certificate_heading",
        "table_of_contents_heading", "table_of_authorities_heading",
    }
    _clear_run_overrides(
        paragraph, rule, clear_character_style=clear_character_style,
        cleanup_mode=cleanup_mode)
    invalid_numbering_removed = (
        clear_invalid_numbering_override(paragraph) if role == "body" else False)
    has_numbering = isinstance(rule.get("numbering"), dict)
    stripped_prefix = _strip_manual_number_prefix(paragraph, role) if has_numbering else False
    remove_numbering = has_numbering or role in {
        "title", "subtitle", "heading_1", "heading_2", "heading_3", "heading_4",
        "abstract_heading", "chapter_heading", "bibliography_heading",
        "appendix_heading", "brief_title", "certificate_heading",
        "table_of_contents_heading", "table_of_authorities_heading",
    }
    _clear_paragraph_overrides(
        paragraph, rule, remove_numbering=remove_numbering,
        cleanup_mode=cleanup_mode)
    paragraph.style = style
    label_prefix_applied = _apply_label_prefix(paragraph, rule)
    fields = ["paragraph_style"]
    if cleanup_mode != "controlled":
        fields.append(f"cleanup_{cleanup_mode}")
    if label_prefix_applied:
        fields.append("label_prefix")
    if has_numbering:
        fields.append("automatic_numbering")
    if stripped_prefix:
        fields.append("manual_number_prefix_removed")
    if invalid_numbering_removed:
        fields.append("invalid_numbering_removed")
    fields.extend(
        field for field in (
            "font_eastasia", "font_ascii", "font_cs", "language", "size_pt",
            "bold", "alignment", "caps", "small_caps", "rtl", "bidi",
            "italic", "underline", "color", "highlight", "strike",
            "line_spacing", "space_before_pt", "space_after_pt",
            "first_line_indent_chars", "left_indent_chars", "hanging_indent_chars",
            "keep_with_next", "keep_together", "page_break_before",
            "widow_control", "outline_level",
            "shading", "paragraph_border",
        )
        if rule.get(field) is not None
    )
    return fields
