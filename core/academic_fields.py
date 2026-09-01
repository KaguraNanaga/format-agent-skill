"""学术文稿的脚注、题注序号、图表目录和域更新支持。"""

from copy import deepcopy
import re

from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.opc.oxml import serialize_part_xml


_CAPTION_PATTERNS = {
    "figure_caption": re.compile(
        r"^(?P<label>图|Figure|Fig\.?)\s*(?P<number>\d+)(?![.\-]\d)", re.I),
    "table_caption": re.compile(
        r"^(?P<label>表|Table)\s*(?P<number>\d+)(?![.\-]\d)", re.I),
}
_SEQ_NAMES = {"figure_caption": "Figure", "table_caption": "Table"}
_LIST_ROLES = {
    "figures": ("list_of_figures_heading", "Figure"),
    "tables": ("list_of_tables_heading", "Table"),
}
_ALIGNMENTS = {
    "left": "left", "center": "center", "right": "right",
    "justify": "both",
}


def _role_for_index(rolemap, index):
    return rolemap.get(index, rolemap.get(str(index)))


def _editable_main_paragraphs(document):
    from core.extract import iter_main_paragraphs
    return [
        (index, paragraph)
        for index, (paragraph, depth) in enumerate(iter_main_paragraphs(document))
        if depth == 0
    ]


def _new_run_like(run, *, text=None, field_char=None, instruction=None):
    element = OxmlElement("w:r")
    rpr = run._r.find(qn("w:rPr"))
    if rpr is not None:
        element.append(deepcopy(rpr))
    if field_char is not None:
        field = OxmlElement("w:fldChar")
        field.set(qn("w:fldCharType"), field_char)
        if field_char == "begin":
            field.set(qn("w:dirty"), "true")
        element.append(field)
    elif instruction is not None:
        instr = OxmlElement("w:instrText")
        instr.set(qn("xml:space"), "preserve")
        instr.text = instruction
        element.append(instr)
    elif text is not None:
        node = OxmlElement("w:t")
        if text[:1].isspace() or text[-1:].isspace():
            node.set(qn("xml:space"), "preserve")
        node.text = text
        element.append(node)
    return element


def _existing_instruction(paragraph, keyword):
    return any(
        keyword.upper() in (element.text or "").upper()
        for element in paragraph._p.iter(qn("w:instrText"))
    )


def _next_bookmark_id(document):
    values = []
    for element in document.element.body.iter(qn("w:bookmarkStart")):
        try:
            values.append(int(element.get(qn("w:id"))))
        except (TypeError, ValueError):
            pass
    return max(values, default=0) + 1


def _unique_bookmark_name(document, base):
    existing = {
        element.get(qn("w:name"))
        for element in document.element.body.iter(qn("w:bookmarkStart"))
    }
    name = base
    counter = 2
    while name in existing:
        name = f"{base}_{counter}"
        counter += 1
    return name


def _fieldify_caption(document, paragraph, role, bookmark_id):
    if _existing_instruction(paragraph, "SEQ"):
        return None, "already_field"
    match = _CAPTION_PATTERNS[role].match(paragraph.text.strip())
    if not match:
        return None, "unsupported_caption_number"
    number = match.group("number")
    global_start = paragraph.text.find(number, match.start("number"))
    global_end = global_start + len(number)
    offset = 0
    target_run = None
    local_start = local_end = None
    for run in paragraph.runs:
        run_end = offset + len(run.text)
        if offset <= global_start and global_end <= run_end:
            target_run = run
            local_start = global_start - offset
            local_end = global_end - offset
            break
        offset = run_end
    if target_run is None:
        return None, "caption_number_spans_runs"

    original = target_run.text
    prefix, suffix = original[:local_start], original[local_end:]
    target_run.text = prefix
    parent = target_run._r.getparent()
    insert_at = parent.index(target_run._r) + 1
    bookmark_name = _unique_bookmark_name(
        document, f"FormatAgent{_SEQ_NAMES[role]}{number}")
    bookmark_start = OxmlElement("w:bookmarkStart")
    bookmark_start.set(qn("w:id"), str(bookmark_id))
    bookmark_start.set(qn("w:name"), bookmark_name)
    bookmark_end = OxmlElement("w:bookmarkEnd")
    bookmark_end.set(qn("w:id"), str(bookmark_id))
    elements = [
        bookmark_start,
        _new_run_like(target_run, field_char="begin"),
        _new_run_like(
            target_run,
            instruction=f" SEQ {_SEQ_NAMES[role]} \\* ARABIC ",
        ),
        _new_run_like(target_run, field_char="separate"),
        _new_run_like(target_run, text=number),
        _new_run_like(target_run, field_char="end"),
        bookmark_end,
    ]
    if suffix:
        elements.append(_new_run_like(target_run, text=suffix))
    for element in elements:
        parent.insert(insert_at, element)
        insert_at += 1
    return {
        "role": role,
        "number": number,
        "bookmark": bookmark_name,
        "sequence": _SEQ_NAMES[role],
    }, None


def _append_complex_field(paragraph, instruction):
    anchor = paragraph.add_run()
    for element in (
        _new_run_like(anchor, field_char="begin"),
        _new_run_like(anchor, instruction=f" {instruction} "),
        _new_run_like(anchor, field_char="separate"),
        _new_run_like(anchor, text=""),
        _new_run_like(anchor, field_char="end"),
    ):
        paragraph._p.append(element)
    paragraph._p.remove(anchor._r)


def _insert_list_fields(document, rolemap, academic, role_styles=None):
    requested = academic.get("lists") or {}
    changed = []
    diagnostics = []
    heading_by_role = {}
    for index, paragraph in _editable_main_paragraphs(document):
        role = _role_for_index(rolemap, index)
        if role in {value[0] for value in _LIST_ROLES.values()}:
            heading_by_role.setdefault(role, paragraph)
    for key, (heading_role, sequence) in _LIST_ROLES.items():
        if not requested.get(key):
            continue
        heading = heading_by_role.get(heading_role)
        if heading is None:
            diagnostics.append({
                "code": f"MISSING_{heading_role.upper()}",
                "message": f"请求生成 {key} 目录，但 RoleMap 中没有 {heading_role}",
            })
            continue
        next_element = heading._p.getnext()
        if next_element is not None and any(
            "TOC" in (node.text or "").upper()
            and f'"{sequence.upper()}"' in (node.text or "").upper()
            for node in next_element.iter(qn("w:instrText"))
        ):
            continue
        field_paragraph = document.add_paragraph()
        if role_styles and role_styles.get("body") is not None:
            field_paragraph.style = role_styles["body"]
        _append_complex_field(
            field_paragraph, f'TOC \\h \\z \\c "{sequence}"')
        heading._p.addnext(field_paragraph._p)
        changed.append(f"list_of_{key}_field")
    return changed, diagnostics


def _get_or_add(parent, tag):
    element = parent.find(qn(tag))
    if element is None:
        element = OxmlElement(tag)
        parent.append(element)
    return element


def _apply_note_rule(root, note_tag, rule):
    changed = 0
    for note in root.findall(qn(f"w:{note_tag}")):
        note_type = note.get(qn("w:type"))
        try:
            note_id = int(note.get(qn("w:id"), "1"))
        except ValueError:
            note_id = 1
        if note_type in {"separator", "continuationSeparator"} or note_id <= 0:
            continue
        for paragraph in note.iter(qn("w:p")):
            ppr = paragraph.find(qn("w:pPr"))
            if ppr is None:
                ppr = OxmlElement("w:pPr")
                paragraph.insert(0, ppr)
            if rule.get("alignment") in _ALIGNMENTS:
                jc = _get_or_add(ppr, "w:jc")
                jc.set(qn("w:val"), _ALIGNMENTS[rule["alignment"]])
            if rule.get("bidi") is not None:
                bidi = _get_or_add(ppr, "w:bidi")
                bidi.set(qn("w:val"), "1" if rule["bidi"] else "0")
            spacing_rule = rule.get("line_spacing") or {}
            if spacing_rule:
                spacing = _get_or_add(ppr, "w:spacing")
                if spacing_rule.get("type") == "exact":
                    spacing.set(qn("w:line"), str(round(spacing_rule["pt"] * 20)))
                    spacing.set(qn("w:lineRule"), "exact")
                elif spacing_rule.get("type") == "multiple":
                    spacing.set(qn("w:line"), str(round(spacing_rule["pt"] * 240)))
                    spacing.set(qn("w:lineRule"), "auto")
            if rule.get("space_after_pt") is not None:
                spacing = _get_or_add(ppr, "w:spacing")
                spacing.set(qn("w:after"), str(round(rule["space_after_pt"] * 20)))
            if rule.get("first_line_indent_chars") is not None:
                indent = _get_or_add(ppr, "w:ind")
                indent.set(
                    qn("w:firstLineChars"),
                    str(round(rule["first_line_indent_chars"] * 100)),
                )
            for run in paragraph.iter(qn("w:r")):
                rpr = run.find(qn("w:rPr"))
                if rpr is None:
                    rpr = OxmlElement("w:rPr")
                    run.insert(0, rpr)
                if (rule.get("font_eastasia") or rule.get("font_ascii")
                        or rule.get("font_cs")):
                    fonts = _get_or_add(rpr, "w:rFonts")
                    if rule.get("font_eastasia"):
                        fonts.set(qn("w:eastAsia"), rule["font_eastasia"])
                    if rule.get("font_ascii"):
                        fonts.set(qn("w:ascii"), rule["font_ascii"])
                        fonts.set(qn("w:hAnsi"), rule["font_ascii"])
                    if rule.get("font_cs"):
                        fonts.set(qn("w:cs"), rule["font_cs"])
                if rule.get("language"):
                    language = _get_or_add(rpr, "w:lang")
                    language.set(qn("w:val"), rule["language"])
                    if rule.get("rtl") or rule.get("bidi"):
                        language.set(qn("w:bidi"), rule["language"])
                if rule.get("rtl") is not None:
                    rtl = _get_or_add(rpr, "w:rtl")
                    rtl.set(qn("w:val"), "1" if rule["rtl"] else "0")
                if rule.get("size_pt") is not None:
                    for tag in ("w:sz", "w:szCs"):
                        size = _get_or_add(rpr, tag)
                        size.set(qn("w:val"), str(round(rule["size_pt"] * 2)))
            changed += 1
    return changed


def _apply_note_rules(document, notes):
    changed = []
    for kind, part_name, note_tag in (
        ("footnote", "/word/footnotes.xml", "footnote"),
        ("endnote", "/word/endnotes.xml", "endnote"),
    ):
        rule = notes.get(kind)
        if not isinstance(rule, dict) or not rule:
            continue
        part = next((
            candidate for candidate in document.part.package.parts
            if str(candidate.partname) == part_name
        ), None)
        if part is None:
            continue
        root = parse_xml(part.blob)
        count = _apply_note_rule(root, note_tag, rule)
        if count:
            part._blob = serialize_part_xml(root)
            changed.append(f"{kind}_paragraphs_{count}")
    return changed


def _set_update_fields_on_open(document):
    update = document.settings.element.find(qn("w:updateFields"))
    if update is None:
        update = OxmlElement("w:updateFields")
        document.settings.element.append(update)
    update.set(qn("w:val"), "true")


def apply_academic_features(document, spec, rolemap, role_styles=None):
    """应用明确启用的学术域与注释格式；不改写引文或文献内容。"""
    academic = spec.get("academic") or {}
    notes = spec.get("notes") or {}
    changed = []
    diagnostics = []
    captions = []

    if academic.get("caption_numbering"):
        bookmark_id = _next_bookmark_id(document)
        for index, paragraph in _editable_main_paragraphs(document):
            role = _role_for_index(rolemap, index)
            if role not in _CAPTION_PATTERNS:
                continue
            result, reason = _fieldify_caption(
                document, paragraph, role, bookmark_id)
            if result:
                result["paragraph_idx"] = index
                captions.append(result)
                bookmark_id += 1
            elif reason not in {"already_field"}:
                diagnostics.append({
                    "code": reason.upper(),
                    "paragraph_idx": index,
                    "message": "题注序号未转换；仅支持同一 run 内的简单整数序号",
                })
        if captions:
            changed.append(f"caption_sequence_fields_{len(captions)}")

    list_changes, list_diagnostics = _insert_list_fields(
        document, rolemap, academic, role_styles=role_styles)
    changed.extend(list_changes)
    diagnostics.extend(list_diagnostics)
    changed.extend(_apply_note_rules(document, notes))
    if changed or academic.get("preserve_fields"):
        _set_update_fields_on_open(document)
        changed.append("update_fields_on_open")
    return {
        "changed_fields": changed,
        "captions": captions,
        "diagnostics": diagnostics,
    }
