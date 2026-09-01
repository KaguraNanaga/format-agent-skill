"""可审计的英文文体基线包；投稿机构模板始终优先。"""

from copy import deepcopy


def _rule(size, alignment="left", **overrides):
    value = {
        "font_eastasia": "Times New Roman",
        "font_ascii": "Times New Roman",
        "size_pt": size,
        "alignment": alignment,
    }
    value.update(overrides)
    return value


def _cn_rule(size, alignment="left", font="仿宋_GB2312", **overrides):
    value = {
        "font_eastasia": font,
        "font_ascii": "Times New Roman",
        "size_pt": size,
        "alignment": alignment,
    }
    value.update(overrides)
    return value


def _apa7_student(options):
    double = {"type": "multiple", "pt": 2.0}
    body = _rule(
        12, "left", line_spacing=double, first_line_indent_chars=3,
        space_before_pt=0, space_after_pt=0, widow_control=True)
    heading_1 = _rule(
        12, "center", bold=True, line_spacing=double,
        keep_with_next=True, keep_together=True)
    heading_2 = _rule(
        12, "left", bold=True, line_spacing=double,
        keep_with_next=True, keep_together=True)
    heading_3 = _rule(
        12, "left", bold=True, italic=True, line_spacing=double,
        keep_with_next=True, keep_together=True)
    bibliography = _rule(
        12, "left", line_spacing=double, left_indent_chars=3,
        hanging_indent_chars=3, space_before_pt=0, space_after_pt=0)
    return {
        "style_pack": "apa7-student",
        "style_pack_notes": [
            "APA 7 permits several accessible font choices; this baseline selects Times New Roman 12 pt.",
            "Course or institution instructions override this baseline.",
            "Citation wording and reference metadata are preserved, not rewritten.",
        ],
        "profile": "english_academic",
        "locale": "en-US",
        "cleanup": {"mode": "preserve_emphasis"},
        "page": {
            "size": "letter", "orientation": "portrait",
            "margin": {"top_mm": 25.4, "bottom_mm": 25.4,
                       "left_mm": 25.4, "right_mm": 25.4},
            "header": {"text": "", "page_number": True,
                       "font_ascii": "Times New Roman", "size_pt": 12,
                       "alignment": "right"},
        },
        "academic": {"preserve_fields": True, "caption_numbering": False,
                     "lists": {"figures": False, "tables": False}},
        "notes": {
            "footnote": _rule(10, "left", line_spacing={"type": "multiple", "pt": 1.0}),
            "endnote": _rule(10, "left", line_spacing={"type": "multiple", "pt": 1.0}),
        },
        "table": {
            "font_ascii": "Times New Roman", "font_eastasia": "Times New Roman",
            "size_pt": 10, "alignment": "center", "layout": "autofit",
            "width_pct": 100, "repeat_header_row": True,
            "allow_row_break": False, "header_alignment": "center",
            "body_alignment": "left", "vertical_alignment": "center",
        },
        "roles": {
            "title": _rule(12, "center", bold=True, line_spacing=double),
            "heading_1": heading_1,
            "heading_2": heading_2,
            "heading_3": heading_3,
            "body": body,
            "abstract_heading": heading_1,
            "abstract_body": _rule(12, "left", line_spacing=double,
                                   first_line_indent_chars=0),
            "keywords": _rule(12, "left", line_spacing=double,
                              first_line_indent_chars=0),
            "chapter_heading": heading_1,
            "figure_caption": _rule(12, "left", line_spacing=double),
            "table_caption": _rule(12, "left", line_spacing=double),
            "bibliography_heading": _rule(
                12, "center", bold=True, line_spacing=double,
                page_break_before=True, keep_with_next=True),
            "bibliography_entry": bibliography,
            "list_of_figures_heading": heading_1,
            "list_of_tables_heading": heading_1,
            "appendix_heading": heading_1,
            "equation": _rule(12, "center", line_spacing=double),
            "block_quote": _rule(
                12, "left", line_spacing=double, left_indent_chars=3,
                first_line_indent_chars=0),
            "code_block": _rule(
                10, "left", font_ascii="Courier New", line_spacing=double,
                first_line_indent_chars=0),
            "byline": _rule(12, "center", line_spacing=double),
            "affiliation": _rule(12, "center", line_spacing=double),
            "author_note": _rule(12, "left", line_spacing=double),
            "correspondence": _rule(12, "left", line_spacing=double),
        },
    }


def _mla9(options):
    double = {"type": "multiple", "pt": 2.0}
    body = _rule(
        12, "left", line_spacing=double, first_line_indent_chars=3,
        space_before_pt=0, space_after_pt=0)
    running_head = str(options.get("running_head") or "").strip()
    header = {
        "text": running_head, "page_number": True,
        "font_ascii": "Times New Roman", "size_pt": 12,
        "alignment": "right",
    }
    notes = []
    if not running_head:
        notes.append(
            "MLA running heads normally include the author's surname; only the page number was added because running_head was not provided.")
    return {
        "style_pack": "mla9",
        "style_pack_notes": notes + [
            "MLA allows another readable font; this baseline selects Times New Roman 12 pt.",
            "Instructor requirements override this baseline.",
            "Works Cited entries and in-text citations are preserved, not rewritten.",
        ],
        "profile": "english_academic",
        "locale": "en-US",
        "cleanup": {"mode": "preserve_emphasis"},
        "page": {
            "size": "letter", "orientation": "portrait",
            "margin": {"top_mm": 25.4, "bottom_mm": 25.4,
                       "left_mm": 25.4, "right_mm": 25.4},
            "header": header,
        },
        "academic": {"preserve_fields": True, "caption_numbering": False,
                     "lists": {"figures": False, "tables": False}},
        "notes": {
            "footnote": _rule(12, "left", line_spacing=double),
            "endnote": _rule(12, "left", line_spacing=double),
        },
        "roles": {
            "title": _rule(12, "center", line_spacing=double),
            "heading_1": _rule(12, "left", bold=True, line_spacing=double,
                               keep_with_next=True),
            "heading_2": _rule(12, "left", italic=True, line_spacing=double,
                               keep_with_next=True),
            "heading_3": _rule(12, "left", line_spacing=double,
                               keep_with_next=True),
            "body": body,
            "abstract_heading": _rule(12, "center", line_spacing=double),
            "abstract_body": _rule(12, "left", line_spacing=double),
            "keywords": _rule(12, "left", line_spacing=double),
            "chapter_heading": _rule(12, "left", bold=True,
                                     line_spacing=double, keep_with_next=True),
            "figure_caption": _rule(12, "left", line_spacing=double),
            "table_caption": _rule(12, "left", line_spacing=double),
            "bibliography_heading": _rule(
                12, "center", line_spacing=double,
                page_break_before=True, keep_with_next=True),
            "bibliography_entry": _rule(
                12, "left", line_spacing=double, left_indent_chars=3,
                hanging_indent_chars=3, space_before_pt=0, space_after_pt=0),
            "list_of_figures_heading": _rule(12, "center", line_spacing=double),
            "list_of_tables_heading": _rule(12, "center", line_spacing=double),
            "appendix_heading": _rule(12, "center", line_spacing=double),
            "equation": _rule(12, "center", line_spacing=double),
            "block_quote": _rule(
                12, "left", line_spacing=double, left_indent_chars=3,
                first_line_indent_chars=0),
            "code_block": _rule(
                10, "left", font_ascii="Courier New", line_spacing=double,
                first_line_indent_chars=0),
            "byline": _rule(12, "center", line_spacing=double),
            "affiliation": _rule(12, "center", line_spacing=double),
        },
    }


def _ieee_journal(options):
    single = {"type": "multiple", "pt": 1.0}
    return {
        "style_pack": "ieee-journal",
        "style_pack_notes": [
            "This is a conservative IEEE journal baseline, not a substitute for the selected publication's official Word template.",
            "Reference numbers, citation order, author names, and bibliographic metadata are preserved and must be validated separately.",
        ],
        "profile": "english_academic",
        "locale": "en-US",
        "cleanup": {"mode": "preserve_emphasis"},
        "page": {
            "size": "letter", "orientation": "portrait", "columns": 2,
            "margin": {"top_mm": 19.05, "bottom_mm": 25.4,
                       "left_mm": 15.875, "right_mm": 15.875},
        },
        "academic": {"preserve_fields": True, "caption_numbering": False,
                     "lists": {"figures": False, "tables": False}},
        "notes": {
            "footnote": _rule(8, "left", line_spacing=single),
            "endnote": _rule(8, "left", line_spacing=single),
        },
        "table": {
            "font_ascii": "Times New Roman", "font_eastasia": "Times New Roman",
            "size_pt": 8, "alignment": "center", "layout": "autofit",
            "width_pct": 100, "repeat_header_row": True,
            "allow_row_break": False, "header_alignment": "center",
            "body_alignment": "center", "vertical_alignment": "center",
        },
        "roles": {
            "title": _rule(24, "center", line_spacing=single),
            "subtitle": _rule(11, "center", line_spacing=single),
            "heading_1": _rule(10, "center", line_spacing=single,
                               keep_with_next=True),
            "heading_2": _rule(10, "left", italic=True, line_spacing=single,
                               keep_with_next=True),
            "heading_3": _rule(10, "left", italic=True, line_spacing=single,
                               keep_with_next=True),
            "body": _rule(10, "justify", line_spacing=single,
                          first_line_indent_chars=1),
            "abstract_heading": _rule(9, "left", bold=True, line_spacing=single),
            "abstract_body": _rule(9, "justify", line_spacing=single),
            "keywords": _rule(9, "justify", line_spacing=single),
            "chapter_heading": _rule(10, "center", line_spacing=single,
                                     keep_with_next=True),
            "figure_caption": _rule(8, "center", line_spacing=single),
            "table_caption": _rule(8, "center", line_spacing=single),
            "bibliography_heading": _rule(10, "center", line_spacing=single,
                                          keep_with_next=True),
            "bibliography_entry": _rule(
                8, "left", line_spacing=single, left_indent_chars=2,
                hanging_indent_chars=2, space_after_pt=0),
            "list_of_figures_heading": _rule(10, "center", line_spacing=single),
            "list_of_tables_heading": _rule(10, "center", line_spacing=single),
            "appendix_heading": _rule(10, "center", line_spacing=single),
            "equation": _rule(10, "center", line_spacing=single),
            "block_quote": _rule(
                10, "justify", line_spacing=single, left_indent_chars=1,
                first_line_indent_chars=0),
            "code_block": _rule(
                8, "left", font_ascii="Courier New", line_spacing=single,
                first_line_indent_chars=0),
            "byline": _rule(11, "center", line_spacing=single),
            "affiliation": _rule(9, "center", line_spacing=single),
        },
    }


def _official_cn_gbt9704(options):
    exact_28 = {"type": "exact", "pt": 28}
    body = _cn_rule(
        16, "justify", first_line_indent_chars=2, line_spacing=exact_28,
        space_before_pt=0, space_after_pt=0, widow_control=True)
    return {
        "style_pack": "official-cn-gbt9704",
        "style_pack_notes": [
            "按 GB/T 9704-2012 的常见版式参数提供保守基线；发文机关标志、红色分隔线、印章和版记需要专门模板。",
            "各机关实施细则和用户提供的正式模板优先于本基线。",
            "只调整版式，不生成或改写公文内容。",
        ],
        "profile": "official_cn",
        "locale": "zh-CN",
        "cleanup": {"mode": "controlled"},
        "page": {
            "size": "A4", "orientation": "portrait",
            "margin": {"top_mm": 37, "bottom_mm": 35,
                       "left_mm": 28, "right_mm": 26},
            "footer_distance_mm": 7,
            "different_odd_even": True,
            "footer": {
                "text": "", "page_number": True,
                "page_number_prefix": "— ", "page_number_suffix": " —",
                "font_eastasia": "宋体", "font_ascii": "Times New Roman",
                "size_pt": 14, "alignment": "right",
            },
            "even_footer": {
                "text": "", "page_number": True,
                "page_number_prefix": "— ", "page_number_suffix": " —",
                "font_eastasia": "宋体", "font_ascii": "Times New Roman",
                "size_pt": 14, "alignment": "left",
            },
        },
        "roles": {
            "title": _cn_rule(22, "center", font="方正小标宋简体",
                              line_spacing=exact_28),
            "subtitle": _cn_rule(16, "center", line_spacing=exact_28),
            "document_number": _cn_rule(16, "center", line_spacing=exact_28),
            "recipient": _cn_rule(16, "left", line_spacing=exact_28),
            "heading_1": _cn_rule(
                16, "justify", font="黑体", first_line_indent_chars=2,
                line_spacing=exact_28, keep_with_next=True),
            "heading_2": _cn_rule(
                16, "justify", font="楷体_GB2312", first_line_indent_chars=2,
                line_spacing=exact_28, keep_with_next=True),
            "heading_3": _cn_rule(
                16, "justify", bold=True, first_line_indent_chars=2,
                line_spacing=exact_28, keep_with_next=True),
            "heading_4": body,
            "body": body,
            "closing": body,
            "signature": _cn_rule(16, "right", line_spacing=exact_28),
            "date": _cn_rule(16, "right", line_spacing=exact_28),
            "attachment_label": _cn_rule(16, "left", line_spacing=exact_28),
            "attachment": body,
            "copy_to": _cn_rule(14, "left", font="仿宋_GB2312",
                                line_spacing={"type": "exact", "pt": 26}),
        },
    }


def _chicago_base(name, author_date=False):
    double = {"type": "multiple", "pt": 2.0}
    single = {"type": "multiple", "pt": 1.0}
    body = _rule(12, "left", line_spacing=double, first_line_indent_chars=3,
                 space_before_pt=0, space_after_pt=0, widow_control=True)
    entry = _rule(
        12, "left", line_spacing=single, left_indent_chars=3,
        hanging_indent_chars=3, space_after_pt=6)
    notes = [
        "Chicago 18 / Turabian 的引证内容不由排版器改写；请用正式引证工具核对。",
        "学校、期刊或出版社模板优先于本通用基线。",
    ]
    if author_date:
        notes.append("此包采用 author-date 的参考文献版式基线。")
    else:
        notes.append("此包采用 notes-bibliography 的脚注/书目版式基线。")
    return {
        "style_pack": name,
        "style_pack_notes": notes,
        "profile": "english_academic", "locale": "en-US",
        "cleanup": {"mode": "preserve_emphasis"},
        "page": {
            "size": "letter", "orientation": "portrait",
            "margin": {"top_mm": 25.4, "bottom_mm": 25.4,
                       "left_mm": 25.4, "right_mm": 25.4},
            "header": {"text": "", "page_number": True,
                       "font_ascii": "Times New Roman", "size_pt": 12,
                       "alignment": "right"},
        },
        "academic": {"preserve_fields": True, "caption_numbering": False,
                     "lists": {"figures": False, "tables": False}},
        "notes": {
            "footnote": _rule(10, "left", line_spacing=single),
            "endnote": _rule(10, "left", line_spacing=single),
        },
        "roles": {
            "title": _rule(14, "center", bold=True, line_spacing=double),
            "heading_1": _rule(12, "center", bold=True, line_spacing=double,
                               keep_with_next=True),
            "heading_2": _rule(12, "left", bold=True, line_spacing=double,
                               keep_with_next=True),
            "heading_3": _rule(12, "left", italic=True, line_spacing=double,
                               keep_with_next=True),
            "body": body,
            "abstract_heading": _rule(12, "center", bold=True,
                                      line_spacing=double),
            "abstract_body": _rule(12, "left", line_spacing=double),
            "keywords": _rule(12, "left", line_spacing=double),
            "chapter_heading": _rule(12, "center", bold=True,
                                     line_spacing=double, keep_with_next=True),
            "bibliography_heading": _rule(
                12, "center", bold=True, line_spacing=double,
                page_break_before=True, keep_with_next=True),
            "bibliography_entry": entry,
            "block_quote": _rule(
                12, "left", line_spacing=single, left_indent_chars=3,
                first_line_indent_chars=0, space_before_pt=6, space_after_pt=6),
            "figure_caption": _rule(10, "center", line_spacing=single),
            "table_caption": _rule(10, "center", line_spacing=single),
            "equation": _rule(12, "center", line_spacing=double),
            "appendix_heading": _rule(12, "center", bold=True,
                                      line_spacing=double),
            "code_block": _rule(10, "left", font_ascii="Courier New",
                                line_spacing=single),
        },
    }


def _chicago_notes(options):
    return _chicago_base("chicago18-notes-bibliography")


def _chicago_author_date(options):
    return _chicago_base("chicago18-author-date", author_date=True)


def _turabian_student(options):
    spec = _chicago_base("turabian9-student")
    spec["style_pack_notes"].append(
        "Turabian 学生论文的院系要求可能规定标题页、页码位置和前置页，需优先遵循。")
    return spec


def _technical_manual(options):
    single = {"type": "multiple", "pt": 1.0}
    body = _rule(11, "left", line_spacing={"type": "multiple", "pt": 1.15},
                 space_after_pt=6, widow_control=True)
    callout = dict(
        line_spacing=single, space_before_pt=6, space_after_pt=6,
        left_indent_chars=1, keep_together=True)
    return {
        "style_pack": "technical-manual",
        "style_pack_notes": [
            "提供代码块、命令、警告/注意/说明/提示框和图题邻接绑定的基础能力。",
            "不移动浮动对象，也不承诺修复文本框、跨栏图或复杂 DTP 布局。",
        ],
        "profile": "english_technical", "locale": "en-US",
        "cleanup": {"mode": "preserve_emphasis"},
        "page": {"size": "A4", "orientation": "portrait",
                 "margin": {"top_mm": 20, "bottom_mm": 20,
                            "left_mm": 22, "right_mm": 22}},
        "technical": {"validate_figure_bindings": True},
        "table": {"font_ascii": "Arial", "font_eastasia": "微软雅黑",
                  "size_pt": 9, "alignment": "center", "layout": "autofit",
                  "width_pct": 100, "repeat_header_row": True,
                  "allow_row_break": False, "header_alignment": "left",
                  "body_alignment": "left", "vertical_alignment": "top"},
        "roles": {
            "title": _rule(22, "left", font_ascii="Arial", bold=True),
            "heading_1": _rule(16, "left", font_ascii="Arial", bold=True,
                               keep_with_next=True, space_before_pt=18,
                               space_after_pt=6),
            "heading_2": _rule(13, "left", font_ascii="Arial", bold=True,
                               keep_with_next=True, space_before_pt=12,
                               space_after_pt=4),
            "heading_3": _rule(11, "left", font_ascii="Arial", bold=True,
                               keep_with_next=True, space_before_pt=8,
                               space_after_pt=3),
            "body": body,
            "procedure_step": _rule(11, "left", font_ascii="Arial", bold=True,
                                    keep_with_next=True, space_before_pt=6),
            "code_block": _rule(9, "left", font_ascii="Consolas",
                                shading="F3F4F6", paragraph_border={
                                    "color": "D1D5DB", "size_pt": 0.5,
                                    "space_pt": 3, "sides": ["top", "bottom", "left", "right"]},
                                line_spacing=single, space_before_pt=4,
                                space_after_pt=4, keep_together=True),
            "command": _rule(9, "left", font_ascii="Consolas", bold=True,
                             shading="EEF2FF", line_spacing=single,
                             keep_together=True),
            "warning_box": _rule(10, "left", font_ascii="Arial", bold=True,
                                 shading="FEE2E2", paragraph_border={
                                     "color": "DC2626", "size_pt": 1,
                                     "space_pt": 4, "sides": ["left"]}, **callout),
            "caution_box": _rule(10, "left", font_ascii="Arial", bold=True,
                                 shading="FEF3C7", paragraph_border={
                                     "color": "D97706", "size_pt": 1,
                                     "space_pt": 4, "sides": ["left"]}, **callout),
            "note_box": _rule(10, "left", font_ascii="Arial",
                              shading="E0F2FE", paragraph_border={
                                  "color": "0284C7", "size_pt": 1,
                                  "space_pt": 4, "sides": ["left"]}, **callout),
            "tip_box": _rule(10, "left", font_ascii="Arial",
                             shading="DCFCE7", paragraph_border={
                                 "color": "16A34A", "size_pt": 1,
                                 "space_pt": 4, "sides": ["left"]}, **callout),
            "figure_caption": _rule(9, "center", font_ascii="Arial",
                                    keep_with_next=True, space_after_pt=6),
            "table_caption": _rule(9, "left", font_ascii="Arial", bold=True,
                                   keep_with_next=True),
        },
    }


def _us_legal_brief(options):
    double = {"type": "multiple", "pt": 2.0}
    single = {"type": "multiple", "pt": 1.0}
    legal = {
        "preserve_toa": True,
        "insert_toa": bool(options.get("insert_toa", False)),
        "create_heading": bool(options.get("create_heading", False)),
        "toa_instruction": str(options.get("toa_instruction") or "TOA \\h \\c \"1\"") ,
    }
    if isinstance(options.get("citation_marks"), list):
        legal["citation_marks"] = deepcopy(options["citation_marks"])
    return {
        "style_pack": "us-legal-brief",
        "style_pack_notes": [
            "这是通用美国法院 brief 基线；具体法院规则（字体、字数、行号、封面颜色和电子提交）必须另行核对。",
            "现有 TA/TOA 域被保留；只有显式提供 citation_marks / insert_toa 时才新增域。",
            "引证准确性和 Bluebook 合规性不由排版器判断。",
        ],
        "profile": "english_legal_brief", "locale": "en-US",
        "cleanup": {"mode": "preserve_emphasis"},
        "page": {"size": "letter", "orientation": "portrait",
                 "margin": {"top_mm": 25.4, "bottom_mm": 25.4,
                            "left_mm": 25.4, "right_mm": 25.4},
                 "footer": {"text": "", "page_number": True,
                            "font_ascii": "Times New Roman", "size_pt": 12,
                            "alignment": "center"}},
        "legal": legal,
        "roles": {
            "title": _rule(14, "center", bold=True, line_spacing=single),
            "court_caption": _rule(12, "center", bold=True, line_spacing=single,
                                   keep_together=True),
            "case_number": _rule(12, "center", bold=True, line_spacing=single),
            "brief_title": _rule(14, "center", bold=True, caps=True,
                                 line_spacing=single, keep_together=True),
            "heading_1": _rule(12, "center", bold=True, caps=True,
                               line_spacing=double, keep_with_next=True),
            "heading_2": _rule(12, "left", bold=True, line_spacing=double,
                               keep_with_next=True),
            "heading_3": _rule(12, "left", italic=True, line_spacing=double,
                               keep_with_next=True),
            "body": _rule(12, "justify", line_spacing=double,
                          first_line_indent_chars=3, widow_control=True),
            "block_quote": _rule(12, "justify", line_spacing=single,
                                 left_indent_chars=3, first_line_indent_chars=0),
            "table_of_contents_heading": _rule(
                12, "center", bold=True, caps=True, page_break_before=True,
                keep_with_next=True),
            "table_of_authorities_heading": _rule(
                12, "center", bold=True, caps=True, page_break_before=True,
                keep_with_next=True),
            "authority_entry": _rule(12, "left", line_spacing=single,
                                     left_indent_chars=3,
                                     hanging_indent_chars=3),
            "counsel_block": _rule(12, "left", line_spacing=single,
                                   keep_together=True),
            "certificate_heading": _rule(12, "center", bold=True,
                                         caps=True, page_break_before=True,
                                         keep_with_next=True),
            "certificate_body": _rule(12, "left", line_spacing=double),
            "signature_block": _rule(12, "left", line_spacing=single,
                                     keep_together=True),
            "legal_definition": _rule(12, "justify", line_spacing=double),
        },
    }


_BUILDERS = {
    "apa7-student": _apa7_student,
    "mla9": _mla9,
    "ieee-journal": _ieee_journal,
    "official-cn-gbt9704": _official_cn_gbt9704,
    "chicago18-notes-bibliography": _chicago_notes,
    "chicago18-author-date": _chicago_author_date,
    "turabian9-student": _turabian_student,
    "technical-manual": _technical_manual,
    "us-legal-brief": _us_legal_brief,
}


def get_style_pack(name, **options):
    try:
        spec = _BUILDERS[name](options)
    except KeyError as exc:
        raise ValueError(
            f"未知 Style Pack {name!r}；可选：{', '.join(sorted(_BUILDERS))}") from exc
    from core.schema import validate_spec
    validate_spec(spec)
    return deepcopy(spec)


def list_style_packs():
    return sorted(_BUILDERS)
