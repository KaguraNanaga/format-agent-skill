# FormatSpec 校验器 —— 全系统的核心契约守门员。
# 规则（PLAN.md 第 4 节）：
#   - roles.body 必填；每个角色字段齐全（font_eastasia、size_pt、alignment 至少）
#   - 数值边界：size_pt ∈ [8,72]、margin ∈ [5,50]mm、first_line_indent_chars ∈ [0,8]
#   - 非法输出带校验错误回喂 LLM 重试（由调用方负责重试）

# 角色 Base 闭集；规范文字可自定义角色键（执行器对未知角色按 other 处理），
# 所以这里只对 Base 角色做提示，不拒绝未知键。
BASE_ROLES = [
    "title", "subtitle", "heading_1", "heading_2", "heading_3",
    "body", "signature", "date", "attachment_label", "attachment",
    "figure_caption", "table_caption", "other",
]

ROLE_REQUIRED_FIELDS = ["font_eastasia", "size_pt", "alignment"]

ALIGNMENTS = {"left", "center", "right", "justify"}

SIZE_PT_RANGE = (8, 72)
MARGIN_MM_RANGE = (5, 50)
INDENT_CHARS_RANGE = (0, 8)
NUMBERING_SUFFIXES = {"tab", "space", "nothing"}
NUMBERING_ALIGNMENTS = {"left", "center", "right"}


class SpecValidationError(Exception):
    def __init__(self, errors):
        self.errors = errors
        super().__init__("FormatSpec 校验失败:\n" + "\n".join(f"- {e}" for e in errors))


def validate_spec(spec):
    """校验 FormatSpec dict。合法返回 None；非法抛 SpecValidationError（带全部错误）。"""
    errors = []
    if not isinstance(spec, dict):
        raise SpecValidationError(["顶层必须是 JSON object"])

    # ---- page ----
    page = spec.get("page")
    if page is not None:
        if not isinstance(page, dict):
            errors.append("page 必须是 object")
        else:
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
            for hf in ("header", "footer"):
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


def _is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)
