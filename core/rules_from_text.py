# 规范文字 → FormatSpec（PLAN.md 6.1 prompt）。
# extract_rules(spec_text, llm) -> FormatSpec；schema 校验失败把错误拼进 prompt 回喂，<=2 次。

import json

from core.schema import SpecValidationError, validate_spec

PROMPT_TEMPLATE = """你是公文/文章排版规范解析器。把用户给的格式规范文字，转换成下面这个 JSON schema，
只输出 JSON，不要任何解释。
schema 角色枚举: title/subtitle/heading_1/heading_2/heading_3/heading_4/body/signature/date/
attachment_label/attachment/figure_caption/table_caption/other，以及论文专用的
abstract_heading/abstract_body/keywords/chapter_heading/bibliography_heading/
bibliography_entry/equation/appendix_heading。规范里没提到的角色不要输出；没提到的字段不要编。
图目录与表目录标题使用 list_of_figures_heading/list_of_tables_heading。
顶层结构: {{"page": {{"size": "A4", "margin": {{"top_mm":..,"bottom_mm":..,"left_mm":..,"right_mm":..}},
"line_grid": {{"line_pt":..}}}}, "roles": {{"body": {{...}}, ...}}}}。roles.body 必填。
字段: font_eastasia(中文字体名)/font_ascii/size_pt(磅)/bold/italic/underline/color(六位RGB)/
alignment(left|center|right|justify)/first_line_indent_chars/left_indent_chars/
hanging_indent_chars(字符数)/line_spacing({{"type":"exact"|"multiple","pt":..}})/
keep_with_next/keep_together/page_break_before/widow_control。
摘要、关键词的行内标签可用 label_prefix：
{{"text":["摘要：","摘要:"],"bold":true}}，只作用于前缀，不把整段加粗。
若规范明确要求自动编号，可在相应角色增加 numbering：
{{"group":"headings","level":0~8,"num_format":"chineseCounting|decimal|...",
"level_text":"%1、","start":1,"suffix":"tab|space|nothing","alignment":"left|center|right"}}。
同一套多级标题必须使用相同 group；没有明确编号要求时不要添加 numbering。
论文规范可在顶层输出 "profile":"thesis" 和 "cleanup":{{"mode":"strict"}}；
中文标准公文可使用 official_cn；英文文稿使用 english_general / english_academic /
english_legal / english_technical / english_legal_brief，并在顶层输出 locale（如 en-US、ar-SA）；英文论文和法律文稿
默认 cleanup.mode=preserve_emphasis，以保留斜体、小型大写、字符样式、语言和 RTL 等语义格式。
复杂文字规则可使用 font_cs、language、rtl、bidi；英文高级角色包括 block_quote、code_block、
byline、affiliation、author_note、correspondence、salutation、complimentary_close、cc、enclosure、
legal_definition、signature_block、table_of_authorities_heading。
中文公文角色包括 recipient、closing、document_number、copy_to；技术手册包括
warning_box/caution_box/note_box/tip_box、procedure_step、command；法律 brief 还包括
court_caption、case_number、brief_title、table_of_contents_heading、authority_entry、
counsel_block、certificate_heading、certificate_body。
页面可使用 size=A3|A4|A5|letter|legal 或成对的 width_mm/height_mm，并使用
orientation=portrait|landscape；已有多节文档只在
page.section_overrides 中按 section_index 点名修改横向节。默认/偶数页/首页页眉页脚分别用
header/footer、even_header/even_footer、first_header/first_footer，并用
different_odd_even、different_first_page 开启对应 Story。
页眉/页脚距离使用 page.header_distance_mm/footer_distance_mm；页码外侧文字使用
page_number_prefix/page_number_suffix。
表格几何可用 table.layout、alignment、width_pct、column_widths_pct、cell_margins_mm、
repeat_header_row、allow_row_break、vertical_alignment；不要改变行列数或合并关系。
脚注/尾注文字规则放在 notes.footnote/notes.endnote。只有规范明确要求自动题注或图表目录时，
才输出 academic.caption_numbering=true 以及 academic.lists.figures/tables=true；
只有规范明确点名需要横向排版的表格序号时，才输出 table.landscape_table_indices（从 0 开始）；
系统保留 CITATION/BIBLIOGRAPHY/REF 等现有域，不改写引用内容。
技术手册只有明确要求图文邻接检查时才输出
technical.validate_figure_bindings=true。法律引证表仅在用户提供精确标记时输出
legal.citation_marks（text/long/short/category）；只有明确要求插入 TOA 域时才输出
legal.insert_toa=true，不能自行生成或改写引证。
一般中文文档默认 profile=general、cleanup.mode=controlled。页面字段: page.margin(毫米)/page.line_grid.line_pt。
数值必须合理: size_pt 8~72, margin 5~50, first_line_indent_chars 0~8。
每个角色至少要有 font_eastasia、size_pt、alignment 三个字段。
规范文字如下：
{spec_text}"""

RETRY_SUFFIX = """
你上一次的输出校验未通过，错误如下：
{errors}
请修正后重新输出完整 JSON，仍然只输出 JSON。"""


def extract_rules(spec_text, llm, max_retries=2, on_event=None):
    """规范文字 → 校验通过的 FormatSpec。重试耗尽仍非法则抛 SpecValidationError。"""
    on_event = on_event or (lambda msg: None)
    prompt = PROMPT_TEMPLATE.format(spec_text=spec_text)
    last_err = None
    for attempt in range(max_retries + 1):
        spec = llm.chat_json(prompt)
        try:
            validate_spec(spec)
            if attempt:
                on_event(f"自我修正成功（第 {attempt + 1} 次输出通过校验）")
            return spec
        except SpecValidationError as e:
            last_err = e
            if attempt < max_retries:
                on_event("FormatSpec 未通过校验，正在把错误回喂给模型自我修正：\n"
                         + "\n".join(f"  - {x}" for x in e.errors))
                errors = "\n".join(f"- {x}" for x in e.errors)
                prompt = PROMPT_TEMPLATE.format(spec_text=spec_text) + RETRY_SUFFIX.format(errors=errors)
    raise last_err


if __name__ == "__main__":
    import sys
    from core.llm import LLMClient
    with open(sys.argv[1], encoding="utf-8") as f:
        text = f.read()
    spec = extract_rules(text, LLMClient())
    print(json.dumps(spec, ensure_ascii=False, indent=2))
