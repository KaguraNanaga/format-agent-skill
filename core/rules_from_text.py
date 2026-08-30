# 规范文字 → FormatSpec（PLAN.md 6.1 prompt）。
# extract_rules(spec_text, llm) -> FormatSpec；schema 校验失败把错误拼进 prompt 回喂，<=2 次。

import json

from core.schema import SpecValidationError, validate_spec

PROMPT_TEMPLATE = """你是公文/文章排版规范解析器。把用户给的格式规范文字，转换成下面这个 JSON schema，
只输出 JSON，不要任何解释。
schema 角色枚举: title/subtitle/heading_1/heading_2/heading_3/body/signature/date/
attachment_label/attachment/other。规范里没提到的角色不要输出；没提到的字段不要编。
顶层结构: {{"page": {{"size": "A4", "margin": {{"top_mm":..,"bottom_mm":..,"left_mm":..,"right_mm":..}},
"line_grid": {{"line_pt":..}}}}, "roles": {{"body": {{...}}, ...}}}}。roles.body 必填。
字段: font_eastasia(中文字体名)/font_ascii/size_pt(磅)/bold/alignment(left|center|
right|justify)/first_line_indent_chars(字符数)/line_spacing({{"type":"exact"|"multiple","pt":..}})。
若规范明确要求自动编号，可在相应角色增加 numbering：
{{"group":"headings","level":0~8,"num_format":"chineseCounting|decimal|...",
"level_text":"%1、","start":1,"suffix":"tab|space|nothing","alignment":"left|center|right"}}。
同一套多级标题必须使用相同 group；没有明确编号要求时不要添加 numbering。
页面字段: page.margin(毫米)/page.line_grid.line_pt。
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
