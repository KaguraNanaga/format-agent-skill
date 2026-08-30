"""Word 自动编号的提取结果落地器。

FormatSpec 中每个角色可以带一个 ``numbering`` object。属于同一 group 的角色
共享一个 abstractNum/num 实例，因而一级、二级标题会成为真正的 Word 多级列表，
而不是把“一、”“（一）”当普通文字写进段落。
"""

import hashlib

from docx.oxml import OxmlElement
from docx.oxml.ns import qn


_GROUP_NAME_PREFIX = "FormatAgent:"


def _int_attr(element, attr, default=-1):
    value = element.get(qn(attr))
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def _next_id(elements, attr):
    return max((_int_attr(el, attr) for el in elements), default=-1) + 1


def _child(parent, tag, value=None):
    element = OxmlElement(tag)
    if value is not None:
        element.set(qn("w:val"), str(value))
    parent.append(element)
    return element


def _find_abstract_by_group(numbering, group):
    marker = _GROUP_NAME_PREFIX + group
    for abstract in numbering.findall(qn("w:abstractNum")):
        name = abstract.find(qn("w:name"))
        if name is not None and name.get(qn("w:val")) == marker:
            return abstract
    return None


def _find_num_for_abstract(numbering, abstract_id):
    for num in numbering.findall(qn("w:num")):
        ref = num.find(qn("w:abstractNumId"))
        if ref is not None and ref.get(qn("w:val")) == str(abstract_id):
            return num
    return None


def _reset_abstract(abstract, group, level_entries):
    for child in list(abstract):
        abstract.remove(child)

    digest = hashlib.sha1(group.encode("utf-8")).hexdigest()[:8].upper()
    _child(abstract, "w:nsid", digest)
    _child(abstract, "w:multiLevelType", "multilevel")
    _child(abstract, "w:tmpl", digest)
    _child(abstract, "w:name", _GROUP_NAME_PREFIX + group)

    for level, entry in sorted(level_entries.items()):
        rule = entry["rule"]
        lvl = OxmlElement("w:lvl")
        lvl.set(qn("w:ilvl"), str(level))
        abstract.append(lvl)
        _child(lvl, "w:start", int(rule.get("start", 1)))
        _child(lvl, "w:numFmt", rule.get("num_format", "decimal"))
        if rule.get("level_restart") is not None:
            _child(lvl, "w:lvlRestart", int(rule["level_restart"]))
        if entry.get("style_id"):
            _child(lvl, "w:pStyle", entry["style_id"])
        if rule.get("is_legal"):
            _child(lvl, "w:isLgl")
        _child(lvl, "w:suff", rule.get("suffix", "tab"))
        _child(lvl, "w:lvlText", rule.get("level_text", f"%{level + 1}."))
        _child(lvl, "w:lvlJc", rule.get("alignment", "left"))

        indent_values = {
            "w:left": rule.get("left_twips"),
            "w:hanging": rule.get("hanging_twips"),
            "w:firstLine": rule.get("first_line_twips"),
        }
        if any(value is not None for value in indent_values.values()):
            ppr = _child(lvl, "w:pPr")
            tab_pos = rule.get("tab_pos_twips")
            if tab_pos is not None:
                tabs = _child(ppr, "w:tabs")
                tab = _child(tabs, "w:tab")
                tab.set(qn("w:val"), "num")
                tab.set(qn("w:pos"), str(int(tab_pos)))
            ind = _child(ppr, "w:ind")
            for attr, value in indent_values.items():
                if value is not None:
                    ind.set(qn(attr), str(int(value)))

        font_values = {
            "w:eastAsia": rule.get("font_eastasia"),
            "w:ascii": rule.get("font_ascii"),
            "w:hAnsi": rule.get("font_ascii"),
            "w:cs": rule.get("font_ascii"),
        }
        if any(font_values.values()) or rule.get("size_pt") is not None or rule.get("bold") is not None:
            rpr = _child(lvl, "w:rPr")
            if any(font_values.values()):
                fonts = _child(rpr, "w:rFonts")
                for attr, value in font_values.items():
                    if value:
                        fonts.set(qn(attr), str(value))
            if rule.get("bold") is not None:
                bold = _child(rpr, "w:b")
                bold.set(qn("w:val"), "1" if rule["bold"] else "0")
            if rule.get("size_pt") is not None:
                half_points = int(round(float(rule["size_pt"]) * 2))
                _child(rpr, "w:sz", half_points)
                _child(rpr, "w:szCs", half_points)


def ensure_numbering_groups(document, roles, style_ids=None):
    """创建/更新 FormatAgent 编号定义，返回 ``{role: (num_id, level)}``。"""
    style_ids = style_ids or {}
    groups = {}
    for role, rule in roles.items():
        numbering_rule = rule.get("numbering")
        if not isinstance(numbering_rule, dict):
            continue
        group = numbering_rule.get("group") or "headings"
        level = int(numbering_rule.get("level", 0))
        groups.setdefault(group, {})[level] = {
            "rule": numbering_rule,
            "style_id": style_ids.get(role),
        }

    if not groups:
        return {}

    numbering = document.part.numbering_part.element
    group_ids = {}
    for group, level_entries in groups.items():
        abstract = _find_abstract_by_group(numbering, group)
        if abstract is None:
            abstract = OxmlElement("w:abstractNum")
            abstract_id = _next_id(
                numbering.findall(qn("w:abstractNum")), "w:abstractNumId")
            abstract.set(qn("w:abstractNumId"), str(abstract_id))
            # abstractNum 必须位于所有 num 实例之前。
            first_num = numbering.find(qn("w:num"))
            if first_num is None:
                numbering.append(abstract)
            else:
                numbering.insert(numbering.index(first_num), abstract)
        else:
            abstract_id = _int_attr(abstract, "w:abstractNumId")
        _reset_abstract(abstract, group, level_entries)

        num = _find_num_for_abstract(numbering, abstract_id)
        if num is None:
            num = OxmlElement("w:num")
            num_id = max(
                1, _next_id(numbering.findall(qn("w:num")), "w:numId"))
            num.set(qn("w:numId"), str(num_id))
            _child(num, "w:abstractNumId", abstract_id)
            numbering.append(num)
        else:
            num_id = _int_attr(num, "w:numId")
        group_ids[group] = num_id

    result = {}
    for role, rule in roles.items():
        numbering_rule = rule.get("numbering")
        if isinstance(numbering_rule, dict):
            group = numbering_rule.get("group") or "headings"
            result[role] = (group_ids[group], int(numbering_rule.get("level", 0)))
    return result


def set_style_numbering(style, num_id, level):
    """把编号绑定到命名样式本身，而不是逐段写 direct numPr。"""
    ppr = style.element.get_or_add_pPr()
    old = ppr.find(qn("w:numPr"))
    if old is not None:
        ppr.remove(old)
    num_pr = OxmlElement("w:numPr")
    ilvl = _child(num_pr, "w:ilvl", int(level))
    num_id_el = _child(num_pr, "w:numId", int(num_id))
    # 局部变量保留明确的 OOXML 顺序，也便于静态检查。
    assert ilvl is not None and num_id_el is not None
    ppr._insert_numPr(num_pr)


def clear_style_numbering(style):
    ppr = style.element.find(qn("w:pPr"))
    if ppr is None:
        return
    num_pr = ppr.find(qn("w:numPr"))
    if num_pr is not None:
        ppr.remove(num_pr)
