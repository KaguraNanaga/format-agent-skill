# FormatSpec 校验器 —— 全系统的核心契约守门员。
# 规则（PLAN.md 第 4 节）：
#   - roles.body 必填；每个角色字段齐全（font_eastasia、size_pt、alignment 至少）
#   - 数值边界：size_pt ∈ [8,72]、margin ∈ [5,50]mm、first_line_indent_chars ∈ [0,8]
#   - 非法输出带校验错误回喂 LLM 重试（由调用方负责重试）

import re

# 角色 Base 闭集；规范文字可自定义角色键（执行器对未知角色按 other 处理），
# 所以这里只对 Base 角色做提示，不拒绝未知键。
BASE_ROLES = [
    "title", "subtitle", "heading_1", "heading_2", "heading_3", "heading_4",
    "body", "signature", "date", "attachment_label", "attachment",
    "figure_caption", "table_caption", "other",
    # 论文专用语义角色。保留通用 heading/body 以兼容既有 RoleMap。
    "abstract_heading", "abstract_body", "keywords", "chapter_heading",
    "bibliography_heading", "bibliography_entry", "equation",
    "appendix_heading", "list_of_figures_heading", "list_of_tables_heading",
    "block_quote", "code_block", "byline", "affiliation", "author_note",
    "correspondence", "salutation", "complimentary_close", "cc", "enclosure",
    "legal_definition", "signature_block", "table_of_authorities_heading",
    # 中文公文。
    "recipient", "closing", "document_number", "copy_to",
    # 技术手册。
    "warning_box", "caution_box", "note_box", "tip_box",
    "procedure_step", "command",
    # 美国法律 brief / TOA。
    "court_caption", "case_number", "brief_title",
    "table_of_contents_heading", "authority_entry", "counsel_block",
    "certificate_heading", "certificate_body",
]

ROLE_REQUIRED_FIELDS = ["font_eastasia", "size_pt", "alignment"]

ALIGNMENTS = {"left", "center", "right", "justify"}

SIZE_PT_RANGE = (8, 72)
MARGIN_MM_RANGE = (5, 50)
INDENT_CHARS_RANGE = (0, 8)
NUMBERING_SUFFIXES = {"tab", "space", "nothing"}
NUMBERING_ALIGNMENTS = {"left", "center", "right"}
CLEANUP_MODES = {"controlled", "strict", "preserve_emphasis"}
DOCUMENT_PROFILES = {
    "general", "thesis",
    "english_general", "english_academic", "english_legal",
    "official_cn", "english_technical", "english_legal_brief",
}
STYLE_PACKS = {
    "apa7-student", "mla9", "ieee-journal",
    "official-cn-gbt9704", "chicago18-notes-bibliography",
    "chicago18-author-date", "turabian9-student",
    "technical-manual", "us-legal-brief",
}
PAGE_SIZES = {"A3", "A4", "A5", "letter", "legal"}
PAGE_ORIENTATIONS = {"portrait", "landscape"}
TABLE_LAYOUTS = {"autofit", "fixed"}
TABLE_VERTICAL_ALIGNMENTS = {"top", "center", "bottom"}
PAGE_NUMBER_FORMATS = {"decimal", "upperRoman", "lowerRoman", "upperLetter", "lowerLetter"}
BOOLEAN_ROLE_FIELDS = {
    "bold", "italic", "underline", "strike", "keep_with_next",
    "keep_together", "page_break_before", "widow_control",
    "caps", "small_caps", "rtl", "bidi",
}
HIGHLIGHT_VALUES = {
    "black", "blue", "cyan", "green", "magenta", "red", "yellow", "white",
    "darkBlue", "darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow",
    "darkGray", "lightGray", "none",
}


class SpecValidationError(Exception):
    def __init__(self, errors):
        self.errors = errors
        super().__init__("FormatSpec 校验失败:\n" + "\n".join(f"- {e}" for e in errors))


def validate_spec(spec):
    """校验 FormatSpec dict。合法返回 None；非法抛 SpecValidationError（带全部错误）。"""
    errors = []
    if not isinstance(spec, dict):
        raise SpecValidationError(["顶层必须是 JSON object"])

    # ---- cleanup ----
    cleanup = spec.get("cleanup")
    if cleanup is not None:
        if not isinstance(cleanup, dict):
            errors.append("cleanup 必须是 object")
        elif cleanup.get("mode") not in CLEANUP_MODES:
            errors.append(
                f"cleanup.mode={cleanup.get('mode')!r} 非法："
                f"必须是 {sorted(CLEANUP_MODES)} 之一")

    profile = spec.get("profile")
    if profile is not None and profile not in DOCUMENT_PROFILES:
        errors.append(f"profile 必须是 {sorted(DOCUMENT_PROFILES)} 之一")
    locale = spec.get("locale")
    if locale is not None and (
        not isinstance(locale, str)
        or not re.fullmatch(r"[A-Za-z]{2,3}(?:-[A-Za-z0-9]{2,8})*", locale)
    ):
        errors.append("locale 必须是 BCP 47 风格语言标签，如 zh-CN、en-US、ar-SA")
    style_pack = spec.get("style_pack")
    if style_pack is not None and style_pack not in STYLE_PACKS:
        errors.append(f"style_pack 必须是 {sorted(STYLE_PACKS)} 之一")

    # ---- structure（论文前置页、多节、页眉页码）----
    structure = spec.get("structure")
    if structure is not None:
        if not isinstance(structure, dict):
            errors.append("structure 必须是 object")
        else:
            enabled = structure.get("enabled")
            if enabled is not None and not isinstance(enabled, bool):
                errors.append("structure.enabled 必须是 boolean")
            mode = structure.get("mode")
            if mode is not None and mode != "thesis":
                errors.append("structure.mode 当前只支持 'thesis'")
            for key in ("cover", "front_matter", "page_numbering", "running_header"):
                value = structure.get(key)
                if value is not None and not isinstance(value, dict):
                    errors.append(f"structure.{key} 必须是 object")
            cover = structure.get("cover")
            if isinstance(cover, dict):
                for key in ("enabled", "logo"):
                    value = cover.get(key)
                    if value is not None and not isinstance(value, bool):
                        errors.append(f"structure.cover.{key} 必须是 boolean")
                for key in ("institution", "type_label", "date_text"):
                    value = cover.get(key)
                    if value is not None and not isinstance(value, str):
                        errors.append(f"structure.cover.{key} 必须是 string")
                metadata = cover.get("metadata")
                if metadata is not None and (
                    not isinstance(metadata, dict)
                    or any(not isinstance(k, str) or not isinstance(v, str)
                           for k, v in metadata.items())
                ):
                    errors.append("structure.cover.metadata 必须是 string → string object")
            front = structure.get("front_matter")
            if isinstance(front, dict):
                for key in ("abstract", "toc", "declarations"):
                    value = front.get(key)
                    if value is not None and not isinstance(value, bool):
                        errors.append(f"structure.front_matter.{key} 必须是 boolean")
            numbering = structure.get("page_numbering")
            if isinstance(numbering, dict):
                for key in ("front_format", "body_format"):
                    value = numbering.get(key)
                    if value is not None and value not in PAGE_NUMBER_FORMATS:
                        errors.append(
                            f"structure.page_numbering.{key} 必须是 "
                            f"{sorted(PAGE_NUMBER_FORMATS)} 之一")
                for key in ("front_start", "body_start"):
                    value = numbering.get(key)
                    if value is not None and (
                        not isinstance(value, int) or isinstance(value, bool)
                        or not (0 <= value <= 10000)
                    ):
                        errors.append(f"structure.page_numbering.{key} 必须是 0~10000 的整数")
            header = structure.get("running_header")
            if isinstance(header, dict):
                for key in ("left_text", "chapter_style_name"):
                    value = header.get(key)
                    if value is not None and not isinstance(value, str):
                        errors.append(f"structure.running_header.{key} 必须是 string")

    # ---- page ----
    page = spec.get("page")
    if page is not None:
        if not isinstance(page, dict):
            errors.append("page 必须是 object")
        else:
            size = page.get("size")
            if size is not None and size not in PAGE_SIZES:
                errors.append(f"page.size 必须是 {sorted(PAGE_SIZES)} 之一")
            width = page.get("width_mm")
            height = page.get("height_mm")
            if (width is None) != (height is None):
                errors.append("page.width_mm 与 page.height_mm 必须同时提供")
            for key, value in (("width_mm", width), ("height_mm", height)):
                if value is not None and (not _is_num(value) or not 50 <= value <= 500):
                    errors.append(f"page.{key} 必须是 50~500 毫米的数值")
            if size is not None and width is not None:
                errors.append("page.size 不能与 width_mm/height_mm 同时提供")
            orientation = page.get("orientation")
            if orientation is not None and orientation not in PAGE_ORIENTATIONS:
                errors.append(
                    f"page.orientation 必须是 {sorted(PAGE_ORIENTATIONS)} 之一")
            for flag in ("different_odd_even", "different_first_page"):
                value = page.get(flag)
                if value is not None and not isinstance(value, bool):
                    errors.append(f"page.{flag} 必须是 boolean")
            for key in ("header_distance_mm", "footer_distance_mm"):
                value = page.get(key)
                if value is not None and (
                    not _is_num(value) or not 0 <= value <= 50
                ):
                    errors.append(f"page.{key} 必须是 0~50 毫米的数值")
            margin = page.get("margin")
            if margin is not None:
                if not isinstance(margin, dict):
                    errors.append("page.margin 必须是 object")
                else:
                    for k in ("top_mm", "bottom_mm", "left_mm", "right_mm"):
                        v = margin.get(k)
                        if v is None:
                            continue
                        if not _is_num(v) or not (MARGIN_MM_RANGE[0] <= v <= MARGIN_MM_RANGE[1]):
                            errors.append(
                                f"page.margin.{k}={v!r} 非法：必须是 {MARGIN_MM_RANGE[0]}~{MARGIN_MM_RANGE[1]} 毫米的数值")
            line_grid = page.get("line_grid")
            if line_grid is not None:
                if not isinstance(line_grid, dict):
                    errors.append("page.line_grid 必须是 object")
                else:
                    v = line_grid.get("line_pt")
                    if v is not None and (not _is_num(v) or not (8 <= v <= 72)):
                        errors.append(f"page.line_grid.line_pt={v!r} 非法：必须是 8~72 磅的数值")
            # 页眉/页脚（可选）：text + 字体字号对齐；footer 支持 page_number
            for hf in (
                "header", "footer", "even_header", "even_footer",
                "first_header", "first_footer",
            ):
                sec = page.get(hf)
                if sec is None:
                    continue
                if not isinstance(sec, dict):
                    errors.append(f"page.{hf} 必须是 object")
                    continue
                v = sec.get("size_pt")
                if v is not None and (not _is_num(v) or not (SIZE_PT_RANGE[0] <= v <= SIZE_PT_RANGE[1])):
                    errors.append(f"page.{hf}.size_pt={v!r} 非法：必须是 {SIZE_PT_RANGE[0]}~{SIZE_PT_RANGE[1]} 磅")
                v = sec.get("alignment")
                if v is not None and v not in ALIGNMENTS:
                    errors.append(f"page.{hf}.alignment={v!r} 非法：必须是 {sorted(ALIGNMENTS)} 之一")
                for flag in ("page_number", "preserve_text"):
                    value = sec.get(flag)
                    if value is not None and not isinstance(value, bool):
                        errors.append(f"page.{hf}.{flag} 必须是 boolean")
                for field in ("font_eastasia", "font_ascii", "font_cs", "language"):
                    value = sec.get(field)
                    if value is not None and (not isinstance(value, str) or not value.strip()):
                        errors.append(f"page.{hf}.{field} 必须是非空字符串")
                for field in ("page_number_prefix", "page_number_suffix"):
                    value = sec.get(field)
                    if value is not None and not isinstance(value, str):
                        errors.append(f"page.{hf}.{field} 必须是 string")
                if sec.get("rtl") is not None and not isinstance(sec.get("rtl"), bool):
                    errors.append(f"page.{hf}.rtl 必须是 boolean")

            overrides = page.get("section_overrides")
            if overrides is not None:
                if not isinstance(overrides, list):
                    errors.append("page.section_overrides 必须是 array")
                else:
                    for i, override in enumerate(overrides):
                        path = f"page.section_overrides[{i}]"
                        if not isinstance(override, dict):
                            errors.append(f"{path} 必须是 object")
                            continue
                        index = override.get("section_index")
                        if not isinstance(index, int) or isinstance(index, bool) or index < 0:
                            errors.append(f"{path}.section_index 必须是非负整数")
                        value = override.get("orientation")
                        if value is not None and value not in PAGE_ORIENTATIONS:
                            errors.append(
                                f"{path}.orientation 必须是 {sorted(PAGE_ORIENTATIONS)} 之一")
                        value = override.get("size")
                        if value is not None and value not in PAGE_SIZES:
                            errors.append(f"{path}.size 必须是 {sorted(PAGE_SIZES)} 之一")
                        width = override.get("width_mm")
                        height = override.get("height_mm")
                        if (width is None) != (height is None):
                            errors.append(f"{path}.width_mm 与 height_mm 必须同时提供")
                        for key, number in (("width_mm", width), ("height_mm", height)):
                            if number is not None and (
                                not _is_num(number) or not 50 <= number <= 500
                            ):
                                errors.append(f"{path}.{key} 必须是 50~500 毫米的数值")
                        if value is not None and width is not None:
                            errors.append(f"{path}.size 不能与 width_mm/height_mm 同时提供")
                        value = override.get("margin")
                        if value is not None:
                            _validate_margin(errors, f"{path}.margin", value)
                        value = override.get("columns")
                        if value is not None and (
                            not isinstance(value, int) or isinstance(value, bool)
                            or not 1 <= value <= 4
                        ):
                            errors.append(f"{path}.columns 必须是 1~4 的整数")
                        for flag in ("different_odd_even", "different_first_page"):
                            value = override.get(flag)
                            if value is not None and not isinstance(value, bool):
                                errors.append(f"{path}.{flag} 必须是 boolean")
                        for key in ("header_distance_mm", "footer_distance_mm"):
                            number = override.get(key)
                            if number is not None and (
                                not _is_num(number) or not 0 <= number <= 50
                            ):
                                errors.append(f"{path}.{key} 必须是 0~50 毫米的数值")
                        numbering = override.get("page_numbering")
                        if numbering is not None:
                            if not isinstance(numbering, dict):
                                errors.append(f"{path}.page_numbering 必须是 object")
                            else:
                                value = numbering.get("format")
                                if value is not None and value not in PAGE_NUMBER_FORMATS:
                                    errors.append(
                                        f"{path}.page_numbering.format 必须是 "
                                        f"{sorted(PAGE_NUMBER_FORMATS)} 之一")
                                value = numbering.get("start")
                                if value is not None and (
                                    not isinstance(value, int) or isinstance(value, bool)
                                    or not 0 <= value <= 10000
                                ):
                                    errors.append(
                                        f"{path}.page_numbering.start 必须是 0~10000 的整数")
                        for hf in (
                            "header", "footer", "even_header", "even_footer",
                            "first_header", "first_footer",
                        ):
                            rule = override.get(hf)
                            if rule is None:
                                continue
                            if not isinstance(rule, dict):
                                errors.append(f"{path}.{hf} 必须是 object")
                                continue
                            value = rule.get("size_pt")
                            if value is not None and (
                                not _is_num(value)
                                or not SIZE_PT_RANGE[0] <= value <= SIZE_PT_RANGE[1]
                            ):
                                errors.append(f"{path}.{hf}.size_pt 非法")
                            value = rule.get("alignment")
                            if value is not None and value not in ALIGNMENTS:
                                errors.append(f"{path}.{hf}.alignment 非法")
                            for flag in ("page_number", "preserve_text"):
                                value = rule.get(flag)
                                if value is not None and not isinstance(value, bool):
                                    errors.append(f"{path}.{hf}.{flag} 必须是 boolean")
                            for field in ("page_number_prefix", "page_number_suffix"):
                                value = rule.get(field)
                                if value is not None and not isinstance(value, str):
                                    errors.append(f"{path}.{hf}.{field} 必须是 string")

            # 多栏（可选）：page.columns = 栏数（论文双栏常见）
            columns = page.get("columns")
            if columns is not None and (not isinstance(columns, int) or not (1 <= columns <= 4)):
                errors.append(f"page.columns={columns!r} 非法：必须是 1~4 的整数")

    # ---- toc（可选）：目录。enabled + levels（收录的标题级别）----
    toc = spec.get("toc")
    if toc is not None:
        if not isinstance(toc, dict):
            errors.append("toc 必须是 object")
        else:
            lv = toc.get("levels")
            if lv is not None and (not isinstance(lv, list)
                                   or any(not isinstance(x, int) or not (1 <= x <= 3) for x in lv)):
                errors.append(f"toc.levels={lv!r} 非法：必须是 [1,2,3] 子集")

    # ---- table（可选）：表格排版规则 ----
    table = spec.get("table")
    if table is not None:
        if not isinstance(table, dict):
            errors.append("table 必须是 object")
        else:
            v = table.get("size_pt")
            if v is not None and (not _is_num(v) or not (SIZE_PT_RANGE[0] <= v <= SIZE_PT_RANGE[1])):
                errors.append(f"table.size_pt={v!r} 非法：必须是 {SIZE_PT_RANGE[0]}~{SIZE_PT_RANGE[1]} 磅")
            for k in ("header_alignment", "body_alignment"):
                v = table.get(k)
                if v is not None and v not in ALIGNMENTS:
                    errors.append(f"table.{k}={v!r} 非法：必须是 {sorted(ALIGNMENTS)} 之一")
            for flag in (
                "borders", "autofit", "repeat_header_row", "allow_row_break",
            ):
                value = table.get(flag)
                if value is not None and not isinstance(value, bool):
                    errors.append(f"table.{flag} 必须是 boolean")
            value = table.get("alignment")
            if value is not None and value not in {"left", "center", "right"}:
                errors.append("table.alignment 必须是 left/center/right 之一")
            value = table.get("vertical_alignment")
            if value is not None and value not in TABLE_VERTICAL_ALIGNMENTS:
                errors.append(
                    f"table.vertical_alignment 必须是 {sorted(TABLE_VERTICAL_ALIGNMENTS)} 之一")
            value = table.get("layout")
            if value is not None and value not in TABLE_LAYOUTS:
                errors.append(f"table.layout 必须是 {sorted(TABLE_LAYOUTS)} 之一")
            value = table.get("width_pct")
            if value is not None and (not _is_num(value) or not 10 <= value <= 100):
                errors.append("table.width_pct 必须是 10~100 的数值")
            value = table.get("preferred_width_mm")
            if value is not None and (not _is_num(value) or not 10 <= value <= 400):
                errors.append("table.preferred_width_mm 必须是 10~400 的数值")
            widths = table.get("column_widths_pct")
            if widths is not None and (
                not isinstance(widths, list) or not widths
                or any(not _is_num(v) or not 0 < v <= 100 for v in widths)
            ):
                errors.append("table.column_widths_pct 必须是非空正数数组")
            landscape_indices = table.get("landscape_table_indices")
            if landscape_indices is not None and (
                not isinstance(landscape_indices, list)
                or not landscape_indices
                or any(
                    not isinstance(index, int) or isinstance(index, bool) or index < 0
                    for index in landscape_indices
                )
                or len(set(landscape_indices)) != len(landscape_indices)
            ):
                errors.append(
                    "table.landscape_table_indices 必须是无重复的非负整数数组")
            overrides = table.get("overrides")
            if overrides is not None:
                if not isinstance(overrides, list):
                    errors.append("table.overrides 必须是 array")
                else:
                    seen_table_indices = set()
                    for i, override in enumerate(overrides):
                        path = f"table.overrides[{i}]"
                        if not isinstance(override, dict):
                            errors.append(f"{path} 必须是 object")
                            continue
                        index = override.get("table_index")
                        if (
                            not isinstance(index, int) or isinstance(index, bool)
                            or index < 0
                        ):
                            errors.append(f"{path}.table_index 必须是非负整数")
                        elif index in seen_table_indices:
                            errors.append(f"{path}.table_index 不得重复")
                        else:
                            seen_table_indices.add(index)
                        nested_rule = {
                            key: value for key, value in override.items()
                            if key != "table_index"
                        }
                        if "overrides" in nested_rule or "landscape_table_indices" in nested_rule:
                            errors.append(f"{path} 不能嵌套 overrides/landscape_table_indices")
                            continue
                        try:
                            validate_spec({
                                "roles": {"body": {
                                    "font_eastasia": "宋体", "size_pt": 10.5,
                                    "alignment": "left",
                                }},
                                "table": nested_rule,
                            })
                        except SpecValidationError as exc:
                            errors.extend(
                                f"{path}.{message.removeprefix('table.')}"
                                for message in exc.errors
                                if not message.startswith("roles")
                            )
            margins = table.get("cell_margins_mm")
            if margins is not None:
                if not isinstance(margins, dict):
                    errors.append("table.cell_margins_mm 必须是 object")
                else:
                    for key in ("top", "bottom", "left", "right"):
                        value = margins.get(key)
                        if value is not None and (
                            not _is_num(value) or not 0 <= value <= 20
                        ):
                            errors.append(
                                f"table.cell_margins_mm.{key} 必须是 0~20 毫米的数值")

    # ---- academic fields / notes ----
    academic = spec.get("academic")
    if academic is not None:
        if not isinstance(academic, dict):
            errors.append("academic 必须是 object")
        else:
            for flag in ("caption_numbering", "preserve_fields"):
                value = academic.get(flag)
                if value is not None and not isinstance(value, bool):
                    errors.append(f"academic.{flag} 必须是 boolean")
            lists = academic.get("lists")
            if lists is not None:
                if not isinstance(lists, dict):
                    errors.append("academic.lists 必须是 object")
                else:
                    for key in ("figures", "tables"):
                        value = lists.get(key)
                        if value is not None and not isinstance(value, bool):
                            errors.append(f"academic.lists.{key} 必须是 boolean")

    notes = spec.get("notes")
    if notes is not None:
        if not isinstance(notes, dict):
            errors.append("notes 必须是 object")
        else:
            for kind in ("footnote", "endnote"):
                rule = notes.get(kind)
                if rule is not None:
                    _validate_note_rule(errors, f"notes.{kind}", rule)

    # ---- technical manual / legal brief ----
    technical = spec.get("technical")
    if technical is not None:
        if not isinstance(technical, dict):
            errors.append("technical 必须是 object")
        else:
            value = technical.get("validate_figure_bindings")
            if value is not None and not isinstance(value, bool):
                errors.append("technical.validate_figure_bindings 必须是 boolean")

    legal = spec.get("legal")
    if legal is not None:
        if not isinstance(legal, dict):
            errors.append("legal 必须是 object")
        else:
            for flag in ("preserve_toa", "insert_toa", "create_heading", "mark_all"):
                value = legal.get(flag)
                if value is not None and not isinstance(value, bool):
                    errors.append(f"legal.{flag} 必须是 boolean")
            instruction = legal.get("toa_instruction")
            if instruction is not None and (
                not isinstance(instruction, str)
                or not instruction.strip().upper().startswith("TOA")
            ):
                errors.append("legal.toa_instruction 必须是以 TOA 开头的域指令")
            marks = legal.get("citation_marks")
            if marks is not None:
                if not isinstance(marks, list):
                    errors.append("legal.citation_marks 必须是 array")
                else:
                    for index, mark in enumerate(marks):
                        path = f"legal.citation_marks[{index}]"
                        if not isinstance(mark, dict):
                            errors.append(f"{path} 必须是 object")
                            continue
                        for field in ("text", "long"):
                            value = mark.get(field)
                            if not isinstance(value, str) or not value.strip():
                                errors.append(f"{path}.{field} 必须是非空字符串")
                        short = mark.get("short")
                        if short is not None and (
                            not isinstance(short, str) or not short.strip()
                        ):
                            errors.append(f"{path}.short 必须是非空字符串")
                        category = mark.get("category", 1)
                        if not isinstance(category, int) or isinstance(category, bool) \
                                or not 1 <= category <= 16:
                            errors.append(f"{path}.category 必须是 1~16 的整数")

    if isinstance(table, dict) and table.get("landscape_table_indices"):
        if isinstance(page, dict) and isinstance(page.get("columns"), int) \
                and page["columns"] > 1:
            errors.append(
                "table.landscape_table_indices 不能与 page.columns 同时使用")
        if isinstance(structure, dict) and structure.get("enabled"):
            errors.append(
                "table.landscape_table_indices 不能与 structure.enabled 同时使用")

    # ---- roles ----
    roles = spec.get("roles")
    if not isinstance(roles, dict) or not roles:
        errors.append("roles 必须是非空 object")
    else:
        if "body" not in roles:
            errors.append("roles.body 必填（正文角色是兜底）")
        for role, rule in roles.items():
            if not isinstance(rule, dict):
                errors.append(f"roles.{role} 必须是 object")
                continue
            for f in ROLE_REQUIRED_FIELDS:
                if f not in rule:
                    errors.append(f"roles.{role} 缺少必填字段 {f}")
            for field in ("font_eastasia", "font_ascii", "font_cs", "language"):
                value = rule.get(field)
                if value is not None and (not isinstance(value, str) or not value.strip()):
                    errors.append(f"roles.{role}.{field} 必须是非空字符串")
            v = rule.get("size_pt")
            if v is not None and (not _is_num(v) or not (SIZE_PT_RANGE[0] <= v <= SIZE_PT_RANGE[1])):
                errors.append(f"roles.{role}.size_pt={v!r} 非法：必须是 {SIZE_PT_RANGE[0]}~{SIZE_PT_RANGE[1]} 磅")
            v = rule.get("alignment")
            if v is not None and v not in ALIGNMENTS:
                errors.append(f"roles.{role}.alignment={v!r} 非法：必须是 {sorted(ALIGNMENTS)} 之一")
            v = rule.get("first_line_indent_chars")
            if v is not None and (not _is_num(v) or not (INDENT_CHARS_RANGE[0] <= v <= INDENT_CHARS_RANGE[1])):
                errors.append(
                    f"roles.{role}.first_line_indent_chars={v!r} 非法：必须是 {INDENT_CHARS_RANGE[0]}~{INDENT_CHARS_RANGE[1]} 字符")
            for f in ("left_indent_chars", "hanging_indent_chars"):
                v = rule.get(f)
                if v is not None and (
                    not _is_num(v) or not (INDENT_CHARS_RANGE[0] <= v <= INDENT_CHARS_RANGE[1])
                ):
                    errors.append(
                        f"roles.{role}.{f}={v!r} 非法：必须是 "
                        f"{INDENT_CHARS_RANGE[0]}~{INDENT_CHARS_RANGE[1]} 字符")
            for f in BOOLEAN_ROLE_FIELDS:
                v = rule.get(f)
                if v is not None and not isinstance(v, bool):
                    errors.append(f"roles.{role}.{f} 必须是 boolean")
            color = rule.get("color")
            if color is not None and (
                not isinstance(color, str) or not re.fullmatch(r"[0-9A-Fa-f]{6}", color)
            ):
                errors.append(f"roles.{role}.color 必须是 6 位十六进制 RGB（如 000000）")
            highlight = rule.get("highlight")
            if highlight is not None and highlight not in HIGHLIGHT_VALUES:
                errors.append(
                    f"roles.{role}.highlight={highlight!r} 非法："
                    f"必须是 Word 高亮颜色枚举之一")
            shading = rule.get("shading")
            if shading is not None and (
                not isinstance(shading, str)
                or not re.fullmatch(r"[0-9A-Fa-f]{6}", shading)
            ):
                errors.append(f"roles.{role}.shading 必须是 6 位十六进制 RGB")
            border = rule.get("paragraph_border")
            if border is not None:
                if not isinstance(border, dict):
                    errors.append(f"roles.{role}.paragraph_border 必须是 object")
                else:
                    color = border.get("color", "000000")
                    if not isinstance(color, str) or not re.fullmatch(
                            r"[0-9A-Fa-f]{6}", color):
                        errors.append(
                            f"roles.{role}.paragraph_border.color 必须是 6 位十六进制 RGB")
                    style = border.get("style", "single")
                    if style not in {"single", "double", "dashed", "dotted"}:
                        errors.append(
                            f"roles.{role}.paragraph_border.style 非法")
                    size = border.get("size_pt", 0.5)
                    if not _is_num(size) or not 0.25 <= size <= 6:
                        errors.append(
                            f"roles.{role}.paragraph_border.size_pt 必须是 0.25~6")
                    space = border.get("space_pt", 1)
                    if not _is_num(space) or not 0 <= space <= 20:
                        errors.append(
                            f"roles.{role}.paragraph_border.space_pt 必须是 0~20")
                    sides = border.get("sides", ["top", "bottom", "left", "right"])
                    if not isinstance(sides, list) or not sides or any(
                        side not in {"top", "bottom", "left", "right"}
                        for side in sides
                    ):
                        errors.append(
                            f"roles.{role}.paragraph_border.sides 非法")
            label_prefix = rule.get("label_prefix")
            if label_prefix is not None:
                if not isinstance(label_prefix, dict):
                    errors.append(f"roles.{role}.label_prefix 必须是 object")
                else:
                    text = label_prefix.get("text")
                    valid_text = (
                        isinstance(text, str) and bool(text)
                        or isinstance(text, list) and bool(text)
                        and all(isinstance(item, str) and item for item in text)
                    )
                    if not valid_text:
                        errors.append(
                            f"roles.{role}.label_prefix.text 必须是非空字符串或非空字符串数组")
                    for f in ("bold", "italic", "underline"):
                        v = label_prefix.get(f)
                        if v is not None and not isinstance(v, bool):
                            errors.append(f"roles.{role}.label_prefix.{f} 必须是 boolean")
                    prefix_color = label_prefix.get("color")
                    if prefix_color is not None and (
                        not isinstance(prefix_color, str)
                        or not re.fullmatch(r"[0-9A-Fa-f]{6}", prefix_color)
                    ):
                        errors.append(
                            f"roles.{role}.label_prefix.color 必须是 6 位十六进制 RGB")
            if (
                rule.get("first_line_indent_chars") not in (None, 0)
                and rule.get("hanging_indent_chars") not in (None, 0)
            ):
                errors.append(
                    f"roles.{role} 不能同时设置首行缩进和悬挂缩进")
            ls = rule.get("line_spacing")
            if ls is not None:
                if not isinstance(ls, dict) or ls.get("type") not in ("exact", "multiple") or not _is_num(ls.get("pt")):
                    errors.append(f'roles.{role}.line_spacing 非法：必须是 {{"type": "exact"|"multiple", "pt": 数值}}')
            v = rule.get("outline_level")
            if v is not None and (not isinstance(v, int) or isinstance(v, bool) or not (0 <= v <= 8)):
                errors.append(f"roles.{role}.outline_level={v!r} 非法：必须是 0~8 的整数")
            for f in ("space_before_pt", "space_after_pt"):
                v = rule.get(f)
                if v is not None and (not _is_num(v) or not (0 <= v <= 100)):
                    errors.append(f"roles.{role}.{f}={v!r} 非法：必须是 0~100 磅的数值")
            numbering = rule.get("numbering")
            if numbering is not None:
                if not isinstance(numbering, dict):
                    errors.append(f"roles.{role}.numbering 必须是 object")
                else:
                    group = numbering.get("group")
                    if not isinstance(group, str) or not group.strip():
                        errors.append(f"roles.{role}.numbering.group 必须是非空字符串")
                    level = numbering.get("level")
                    if not isinstance(level, int) or isinstance(level, bool) or not (0 <= level <= 8):
                        errors.append(f"roles.{role}.numbering.level 必须是 0~8 的整数")
                    for field in ("num_format", "level_text"):
                        value = numbering.get(field)
                        if not isinstance(value, str) or not value:
                            errors.append(f"roles.{role}.numbering.{field} 必须是非空字符串")
                    start = numbering.get("start", 1)
                    if not isinstance(start, int) or isinstance(start, bool) or not (0 <= start <= 10000):
                        errors.append(f"roles.{role}.numbering.start 必须是 0~10000 的整数")
                    level_restart = numbering.get("level_restart")
                    if level_restart is not None and (
                        not isinstance(level_restart, int) or isinstance(level_restart, bool)
                        or not (0 <= level_restart <= 9)
                    ):
                        errors.append(
                            f"roles.{role}.numbering.level_restart 必须是 0~9 的整数")
                    suffix = numbering.get("suffix", "tab")
                    if suffix not in NUMBERING_SUFFIXES:
                        errors.append(
                            f"roles.{role}.numbering.suffix={suffix!r} 非法："
                            f"必须是 {sorted(NUMBERING_SUFFIXES)} 之一")
                    alignment = numbering.get("alignment", "left")
                    if alignment not in NUMBERING_ALIGNMENTS:
                        errors.append(
                            f"roles.{role}.numbering.alignment={alignment!r} 非法："
                            f"必须是 {sorted(NUMBERING_ALIGNMENTS)} 之一")
                    for field in (
                        "left_twips", "hanging_twips", "first_line_twips", "tab_pos_twips",
                    ):
                        value = numbering.get(field)
                        if value is not None and (
                            not isinstance(value, int) or isinstance(value, bool)
                            or not (-20000 <= value <= 20000)
                        ):
                            errors.append(
                                f"roles.{role}.numbering.{field} 必须是 -20000~20000 的整数")
                    number_size = numbering.get("size_pt")
                    if number_size is not None and (
                        not _is_num(number_size)
                        or not (SIZE_PT_RANGE[0] <= number_size <= SIZE_PT_RANGE[1])
                    ):
                        errors.append(
                            f"roles.{role}.numbering.size_pt 必须是 8~72 磅的数值")

    if errors:
        raise SpecValidationError(errors)


def _validate_margin(errors, path, margin):
    if not isinstance(margin, dict):
        errors.append(f"{path} 必须是 object")
        return
    for key in ("top_mm", "bottom_mm", "left_mm", "right_mm"):
        value = margin.get(key)
        if value is not None and (
            not _is_num(value)
            or not MARGIN_MM_RANGE[0] <= value <= MARGIN_MM_RANGE[1]
        ):
            errors.append(
                f"{path}.{key} 必须是 {MARGIN_MM_RANGE[0]}~"
                f"{MARGIN_MM_RANGE[1]} 毫米的数值")


def _validate_note_rule(errors, path, rule):
    if not isinstance(rule, dict):
        errors.append(f"{path} 必须是 object")
        return
    size = rule.get("size_pt")
    if size is not None and (not _is_num(size) or not 6 <= size <= 72):
        errors.append(f"{path}.size_pt 必须是 6~72 磅的数值")
    alignment = rule.get("alignment")
    if alignment is not None and alignment not in ALIGNMENTS:
        errors.append(f"{path}.alignment 必须是 {sorted(ALIGNMENTS)} 之一")
    for field in ("font_eastasia", "font_ascii", "font_cs", "language"):
        value = rule.get(field)
        if value is not None and (not isinstance(value, str) or not value.strip()):
            errors.append(f"{path}.{field} 必须是非空字符串")
    for field in ("rtl", "bidi"):
        value = rule.get(field)
        if value is not None and not isinstance(value, bool):
            errors.append(f"{path}.{field} 必须是 boolean")
    for field in ("first_line_indent_chars",):
        value = rule.get(field)
        if value is not None and (
            not _is_num(value) or not INDENT_CHARS_RANGE[0] <= value <= INDENT_CHARS_RANGE[1]
        ):
            errors.append(f"{path}.{field} 必须是 0~8 的数值")
    value = rule.get("space_after_pt")
    if value is not None and (not _is_num(value) or not 0 <= value <= 100):
        errors.append(f"{path}.space_after_pt 必须是 0~100 磅的数值")
    spacing = rule.get("line_spacing")
    if spacing is not None and (
        not isinstance(spacing, dict)
        or spacing.get("type") not in {"exact", "multiple"}
        or not _is_num(spacing.get("pt"))
    ):
        errors.append(
            f"{path}.line_spacing 必须是 "
            '{"type":"exact"|"multiple","pt":数值}')


def _is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)
