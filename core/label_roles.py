# 段落清单 → RoleMap（PLAN.md 6.2 prompt）。
# label_roles(paragraphs, llm) -> dict[int, str]；每 40 段一批；
# 校验 role ∈ 枚举、idx 全覆盖；失败重试 <=2 次。

import json
import re

from core.extract import manual_number_prefix
from core.profiles import detect_profile_from_texts
from core.schema import BASE_ROLES

BATCH_SIZE = 40

# 中文公文标题编号惯例（确定性识别，不走 LLM，解决二级标题识别不稳）：
#   一、    → heading_1    （一）  → heading_2
# 数字前缀 1./2、/1.2 可能是标题，也可能是正文列表，不再单凭前缀强制标题。
# 图表题注（论文/招股书类文档）：
#   图1 xxx / 图 2-1 xxx → figure_caption    表1 xxx / 表 2-1 xxx → table_caption
# 仅当行较短且不以句读结尾时才认定为标题（正文句也可能以"一、"开头）。
_HEADING_PATTERNS = [
    (re.compile(r"^图\s*\d+([-.]\d+)?"), "figure_caption"),
    (re.compile(r"^表\s*\d+([-.]\d+)?"), "table_caption"),
    (re.compile(r"^[一二三四五六七八九十百]+、"), "heading_1"),
    (re.compile(r"^[（(][一二三四五六七八九十]+[）)]"), "heading_2"),
    # 法律文书/合同条款惯例：第一章/第X编 → heading_1；第X节/第X条 → heading_2
    (re.compile(r"^第[一二三四五六七八九十百零〇\d]+[章编部][\s：:、]"), "heading_1"),
    (re.compile(r"^第[一二三四五六七八九十百零〇\d]+[节条][\s：:、]"), "heading_2"),
]
_HEADING_MAX_LEN = 40

_THESIS_EXACT_ROLES = {
    "摘要": "abstract_heading",
    "摘 要": "abstract_heading",
    "参考文献": "bibliography_heading",
    "参考 文献": "bibliography_heading",
}

_ENGLISH_ACADEMIC_EXACT_ROLES = {
    "abstract": "abstract_heading",
    "references": "bibliography_heading",
    "bibliography": "bibliography_heading",
    "works cited": "bibliography_heading",
    "list of figures": "list_of_figures_heading",
    "list of tables": "list_of_tables_heading",
    "keywords": "keywords",
    "keyword": "keywords",
    "author note": "author_note",
    "table of authorities": "table_of_authorities_heading",
}

_ENGLISH_ACADEMIC_SECTIONS = {
    "introduction", "literature review", "related work", "method", "methods",
    "methodology", "results", "discussion", "conclusion", "conclusions",
    "acknowledgment", "acknowledgments", "acknowledgement", "acknowledgements",
}


def detect_document_profile(paragraphs):
    """用高置信语义信号识别受支持的中文与英文文稿。"""
    texts = [str(p.get("text") or "").strip() for p in paragraphs if not p.get("in_table")]
    return detect_profile_from_texts(texts)


def _heading_role_from_metadata(paragraph):
    if not paragraph:
        return None
    outline = paragraph.get("outline_level")
    if isinstance(outline, int) and not isinstance(outline, bool):
        if outline == 0:
            return "heading_1"
        if outline == 1:
            return "heading_2"
        if 2 <= outline <= 8:
            return "heading_3"
    style_name = str(paragraph.get("style_name") or "").strip().lower()
    match = re.search(r"(?:heading|title| 标题|标题)\s*([123])\b", " " + style_name)
    if match:
        return f"heading_{match.group(1)}"
    return None


def _looks_like_body_style(paragraph):
    style_name = str((paragraph or {}).get("style_name") or "").strip().lower()
    return bool(
        style_name in {"normal", "normal indent", "正文", "格式正文", "body", "body text"}
        or "正文" in style_name or style_name.startswith("body")
    )


def _auto_number_heading_role(paragraph):
    """真自动编号（文字里没有前缀）的标题惯例，与手工前缀的判定保持一致：
    chineseCounting "%1、"（显示为"一、"）→ heading_1；
    level_text 为 "（%1）" 形态（显示为"（一）"）→ heading_2。
    只在调用方确认"短行、非完整句"之后调用；decimal 等列表编号返回 None。
    """
    fmt = str((paragraph or {}).get("num_format") or "").strip()
    level_text = str((paragraph or {}).get("level_text") or "")
    if fmt in ("chineseCounting", "chineseCountingThousand", "chineseLegalSimplified") \
            or re.search(r"%1\s*、", level_text):
        return "heading_1"
    if re.search(r"[（(]\s*%1\s*[）)]", level_text):
        return "heading_2"
    return None


def _numbered_body_evidence(paragraph):
    """判断数字段是否有足够证据应作为正文列表项。"""
    if not paragraph:
        return False
    if paragraph.get("list_kind") not in {"manual", "automatic"}:
        return False
    if paragraph.get("list_sequence"):
        return True
    if paragraph.get("ends_with_sentence_punct"):
        return True
    if int(paragraph.get("char_count") or 0) > _HEADING_MAX_LEN:
        return True
    if _looks_like_body_style(paragraph):
        return True
    if (
        paragraph.get("numbering_status") in {"automatic", "cancelled"}
        and _heading_role_from_metadata(paragraph) is None
    ):
        return True
    return False


def regex_role(text, paragraph=None, profile=None):
    """确定性预识别；无把握返回 None（交给 LLM）。

    ``paragraph`` 是 extract.py 的元数据记录。对数字前缀，必须使用
    样式/大纲/真自动编号/序列等结构证据，前缀本身不足以判标题。
    """
    t = text.strip()
    if not t:
        return None

    compact = t.replace(" ", "")
    lower = re.sub(r"\s+", " ", t).strip().lower()
    if profile == "official_cn":
        if re.search(r"〔\d{4}〕\d+号$", compact):
            return "document_number"
        if re.match(r"^抄送[：:]", compact):
            return "copy_to"
        if re.match(r"^附件[：:]", compact):
            return "attachment_label"
        if re.match(r"^(?:特此通知|特此报告|特此请示|特此函告)[。！]?$", compact):
            return "closing"
        if (
            len(t) <= 40 and re.search(r"[：:]$", t)
            and not re.match(r"^(?:摘要|关键词|关键字|附件|抄送|签发人)[：:]", compact)
        ):
            return "recipient"
        if re.match(r"^[（(]\d+[）)]", t):
            return "heading_4"

    if profile == "english_technical":
        callouts = (
            (r"^WARNING\s*[:—-]", "warning_box"),
            (r"^(?:CAUTION|IMPORTANT)\s*[:—-]", "caution_box"),
            (r"^NOTE\s*[:—-]", "note_box"),
            (r"^TIP\s*[:—-]", "tip_box"),
        )
        for pattern, role in callouts:
            if re.match(pattern, t, re.I):
                return role
        if re.match(r"^(?:Step|Procedure)\s+\d+\b", t, re.I):
            return "procedure_step"
        if re.match(r"^(?:\$|>|PS>)\s*\S+", t) or re.match(
            r"^(?:GET|POST|PUT|PATCH|DELETE)\s+/\S+", t, re.I
        ):
            return "command"

    if profile == "english_legal_brief":
        if lower == "table of contents":
            return "table_of_contents_heading"
        if lower == "table of authorities":
            return "table_of_authorities_heading"
        if re.match(r"^IN THE .+ COURT\b", t, re.I):
            return "court_caption"
        if re.match(r"^(?:CASE\s+)?NO\.\s*[A-Z0-9:-]+", t, re.I):
            return "case_number"
        if re.match(r"^(?:BRIEF OF|APPELLANT'?S BRIEF|APPELLEE'?S BRIEF)", t, re.I):
            return "brief_title"
        if re.match(r"^(?:CERTIFICATE OF|CERTIFICATION OF)\b", t, re.I):
            return "certificate_heading"
        if re.match(r"^(?:Respectfully submitted|Counsel for|Attorney for)\b", t, re.I):
            return "counsel_block"
        if re.match(r"^.+(?:\.{3,}|\t+)\s*\d+(?:\s*,\s*\d+)*$", t):
            return "authority_entry"
    if t in _THESIS_EXACT_ROLES or compact in _THESIS_EXACT_ROLES:
        return _THESIS_EXACT_ROLES.get(t, _THESIS_EXACT_ROLES.get(compact))
    if compact in {"图目录", "插图目录"}:
        return "list_of_figures_heading"
    if compact in {"表目录", "表格目录"}:
        return "list_of_tables_heading"
    if re.match(r"^摘要[：:]", compact):
        return "abstract_body"
    if re.match(r"^(关键词|关键字)[：:]", compact):
        return "keywords"
    if re.match(r"^\[\s*\d+\s*\]", t):
        return "bibliography_entry"
    if lower in _ENGLISH_ACADEMIC_EXACT_ROLES:
        return _ENGLISH_ACADEMIC_EXACT_ROLES[lower]
    if re.match(r"^Abstract\s*[:—-]", t, re.I):
        return "abstract_body"
    if re.match(r"^Keywords?\s*[:—-]", t, re.I):
        return "keywords"
    if re.match(r"^(Fig(?:ure)?\.?|Table)\s+\d+(?:[.-]\d+)*\b", t, re.I):
        return "figure_caption" if t.lower().startswith(("fig", "figure")) else "table_caption"
    if re.match(r"^(附录|Appendix)\s*[A-Z一二三四五六七八九十\d]*", t, re.I):
        return "appendix_heading"
    if re.match(r"^Equation\s+(?:\d+|[IVXLC]+)(?:[.-]\d+)*\b", t, re.I):
        return "equation"
    if len(t) <= 80 and "=" in t and not t.endswith(("。", "；", ";")):
        return "equation"

    style_name = str((paragraph or {}).get("style_name") or "").lower()
    ascii_font = str((paragraph or {}).get("font_ascii") or "").lower()
    if "code" in style_name or "代码" in style_name or ascii_font in {
        "consolas", "courier new", "menlo", "monaco", "source code pro",
    }:
        return "code_block"
    if "quote" in style_name or "引用" in style_name or "blockquote" in style_name:
        return "block_quote"
    if (
        profile == "english_academic"
        and int((paragraph or {}).get("indent_left_twips") or 0) >= 720
        and len(t) >= 80
    ):
        return "block_quote"
    if re.match(r"^By\s+[A-Z][\w.'’-]+(?:\s+[A-Z][\w.'’-]+){0,5}$", t):
        return "byline"
    if re.search(
        r"\b(Department|University|Institute|College|School|Faculty|Laboratory)\b",
        t, re.I,
    ) and len(t) <= 140:
        return "affiliation"
    if re.match(
        r"^(Correspondence concerning|Corresponding author|Address correspondence|E-?mail\s*:)",
        t, re.I,
    ):
        return "correspondence"
    if re.match(r"^(Dear\b|To Whom It May Concern\b)", t, re.I):
        return "salutation"
    if re.match(
        r"^(Sincerely|Yours (?:sincerely|faithfully)|Respectfully(?: submitted)?)[,.]?$",
        t, re.I,
    ):
        return "complimentary_close"
    if re.match(r"^(?:CC|Cc|Copy)\s*:", t):
        return "cc"
    if re.match(r"^Enclosures?\s*:", t, re.I):
        return "enclosure"
    if profile in {"english_legal", "english_legal_brief"} and re.match(
        r"^[“\"]?[A-Z][^”\"]{0,50}[”\"]?\s+(?:means|shall mean|has the meaning)",
        t,
    ):
        return "legal_definition"
    if profile in {"english_legal", "english_legal_brief"} and re.match(
        r"^(?:IN WITNESS WHEREOF\b|SIGNATURES?\b|(?:By|Name|Title)\s*:\s*)",
        t, re.I,
    ):
        return "signature_block"

    manual = (paragraph or {}).get("manual_number") or (
        (manual_number_prefix(t) or {}).get("label"))
    metadata_heading = _heading_role_from_metadata(paragraph)
    is_long = int((paragraph or {}).get("char_count") or len(t)) > _HEADING_MAX_LEN
    ends_sentence = bool((paragraph or {}).get("ends_with_sentence_punct")) or t.endswith(
        ("。", "；", ";", "！", "!", "？", "?", "，", ",", "：", ":"))

    if manual:
        if is_long or ends_sentence:
            return "body"
        if metadata_heading:
            return "chapter_heading" if profile == "thesis" and metadata_heading == "heading_1" else metadata_heading
        if _numbered_body_evidence(paragraph):
            return "body"
        if profile in {"english_legal", "english_legal_brief"}:
            depth = str(manual).count(".") + 1
            return "heading_2" if depth <= 2 else "heading_3"
        if profile == "english_academic":
            return "chapter_heading" if "." not in str(manual) else "heading_2"
        return None

    # 真自动编号的标记不在 paragraph.text 里，需靠 numPr + 大纲/样式判断。
    if (paragraph or {}).get("numbering_status") == "automatic":
        if is_long or ends_sentence:
            return "body"
        if metadata_heading:
            return "chapter_heading" if profile == "thesis" and metadata_heading == "heading_1" else metadata_heading
        # 公文标题编号惯例（一、/（一））优先于"正文列表"推断
        auto_heading = _auto_number_heading_role(paragraph)
        if auto_heading:
            return "chapter_heading" if profile == "thesis" and auto_heading == "heading_1" else auto_heading
        if _numbered_body_evidence(paragraph):
            return "body"

    if is_long or ends_sentence:
        return None
    if profile == "english_academic":
        if lower in _ENGLISH_ACADEMIC_SECTIONS:
            return "chapter_heading"
        if re.match(r"^Chapter\s+(?:\d+|[IVXLC]+)\b", t, re.I):
            return "chapter_heading"
        section = re.match(r"^(?:Part|Section)\s+(\d+(?:\.\d+)*)\b", t, re.I)
        if section:
            depth = section.group(1).count(".") + 1
            return "heading_1" if depth == 1 else (
                "heading_2" if depth == 2 else "heading_3")
        numbered = re.match(r"^(\d+(?:\.\d+)*)\s+\S", t)
        if numbered:
            depth = numbered.group(1).count(".") + 1
            return "chapter_heading" if depth == 1 else (
                "heading_2" if depth == 2 else "heading_3")
    if profile in {"english_legal", "english_legal_brief"}:
        if re.match(r"^ARTICLE\s+(?:\d+|[IVXLC]+)\b", t, re.I):
            return "heading_1"
        section = re.match(r"^(?:Section\s+)?(\d+(?:\.\d+)*)\b", t, re.I)
        if section:
            depth = section.group(1).count(".") + 1
            return "heading_2" if depth <= 2 else "heading_3"
        if re.match(r"^(Exhibit|Schedule|Annex)\s+[A-Z\d]+\b", t, re.I):
            return "attachment_label"
    if profile in {"english_academic", "english_legal", "english_legal_brief",
                   "english_technical"}:
        if re.match(r"^[A-Z][.)]\s+\S", t):
            return "heading_2"
        if re.match(r"^\([a-z]\)\s+\S", t):
            return "heading_3"
        if re.match(r"^[IVXLC]+[.)]\s+\S", t, re.I):
            return "heading_2"
    for pat, role in _HEADING_PATTERNS:
        if pat.match(t):
            return "chapter_heading" if profile == "thesis" and role == "heading_1" else role
    return None

PROMPT_TEMPLATE = """你是文档结构标注器。当前文档 profile={profile}。给每一段标注角色，角色只能从枚举里选:
{roles}。
判断依据: 文字内容、位置顺序、当前格式提示。落款单位通常在末尾、署名感强;
日期含"年/月/日"; 标题通常在最前且独立成行。
中文标题层级惯例: "一、"开头多为一级标题(heading_1)，"（一）"开头多为二级标题
(heading_2)。"1."/"2、"/"1.2" 可能是三级标题，也可能是正文编号项：
长句、句末句号、连续 1/2/3 序列、Normal/正文样式或无大纲级别时应标 body；
只有短且独立成行，并有 Heading/标题样式或 outline_level 等强证据时才标 heading_3。
numbering_status=automatic 是真 Word 自动编号；cancelled 表示 numId=0 或 ilvl<0，
它不是自动编号。list_kind=manual 表示数字是文本内容，不能擅自删除。
"图1 xxx"/"图2-1 xxx"这类独立成行的是图片题注(figure_caption)，"表1 xxx"是表格题注
(table_caption)。
论文中摘要标题/摘要正文/关键词分别使用 abstract_heading、abstract_body、keywords；
章标题使用 chapter_heading；“参考文献”及其 [1] 条目分别使用
bibliography_heading、bibliography_entry；公式与附录标题使用 equation、appendix_heading。
“图目录/插图目录/List of Figures”和“表目录/List of Tables”分别使用
list_of_figures_heading、list_of_tables_heading。
英文 academic profile 中 Abstract/Keywords/References/Works Cited、Chapter、
Introduction/Methods/Results/Discussion/Conclusion、Section、Equation、Figure/Table caption 使用对应论文角色；
Byline、Affiliation、Author Note、Correspondence、Block Quote 和 Code Block 使用各自角色。
英文 legal profile 中 ARTICLE 为 heading_1，Section 1.1 为 heading_2，更深条款为 heading_3；
字母条款 A. 与 (a)、定义条款、签署栏、Exhibit/Schedule/Annex 和 Table of Authorities
使用对应角色。英文普通信函还应区分 salutation、complimentary_close、cc、enclosure、signature/date。
official_cn profile 还应区分 document_number、recipient、closing、copy_to 和 heading_4。
english_technical profile 区分 procedure_step、command/code_block 以及
warning_box/caution_box/note_box/tip_box。english_legal_brief profile 区分 court_caption、
case_number、brief_title、table_of_contents_heading、table_of_authorities_heading、
authority_entry、counsel_block、certificate_heading/certificate_body。
输入是 JSON 数组，包含文本/字数/样式/大纲级别、list_kind、numbering_status、
num_id/num_level/num_format/level_text、手工前缀、是否连续序列以及段落/编号缩进。
输出严格为 {{"roles": [{{"idx": 0, "role": "title"}}, ...]}} 的 JSON 对象，
roles 数组必须覆盖所有输入 idx，不多不少。
段落清单：
{paragraphs}"""

RETRY_SUFFIX = """
你上一次的输出校验未通过，错误：{error}
请重新输出完整 JSON 数组，必须恰好覆盖这些 idx: {idx_list}。"""

_ROLE_SET = set(BASE_ROLES)


def _validate_rolemap(items, expected_idxs):
    """校验 LLM 输出：结构、role 合法、idx 恰好全覆盖。返回 dict[int,str] 或抛 ValueError。
    兼容两种形态：裸数组 [...]（老 prompt），或包一层对象 {"roles": [...]}（JSON 模式友好）。
    """
    if isinstance(items, dict):
        # 优先认 "roles" 键，其次取第一个 list 类型的值
        if isinstance(items.get("roles"), list):
            items = items["roles"]
        else:
            for v in items.values():
                if isinstance(v, list):
                    items = v
                    break
    if not isinstance(items, list):
        raise ValueError("输出必须是 JSON 数组或含数组的 JSON 对象")
    rolemap = {}
    for it in items:
        if not isinstance(it, dict) or "idx" not in it or "role" not in it:
            raise ValueError(f"数组元素必须是 {{idx, role}} 对象，收到: {it!r}")
        idx, role = it["idx"], it["role"]
        if role not in _ROLE_SET:
            raise ValueError(f"非法角色 {role!r}（idx={idx}）")
        rolemap[idx] = role
    got, want = set(rolemap), set(expected_idxs)
    if got != want:
        missing = sorted(want - got)
        extra = sorted(got - want)
        raise ValueError(f"idx 覆盖不符：缺少 {missing}，多出 {extra}")
    return rolemap


def _label_batch(batch, llm, max_retries=2, on_event=None, profile="general"):
    on_event = on_event or (lambda msg: None)
    expected = [p["idx"] for p in batch]
    metadata_fields = (
        "idx", "text", "char_count", "ends_with_sentence_punct",
        "size_pt", "bold", "italic", "underline", "color", "alignment",
        "font_ascii", "font_cs", "language", "caps", "small_caps", "rtl",
        "style_name", "outline_level",
        "space_before_pt", "space_after_pt", "list_kind", "manual_number",
        "list_sequence", "numbering_status", "numbering_source", "num_id",
        "num_level", "num_format", "level_text", "indent_left_twips",
        "indent_hanging_twips", "indent_first_line_twips",
        "numbering_left_twips", "numbering_hanging_twips",
        "numbering_first_line_twips",
    )
    payload = json.dumps(
        [{key: p.get(key) for key in metadata_fields} for p in batch],
        ensure_ascii=False)
    prompt = PROMPT_TEMPLATE.format(
        roles="/".join(BASE_ROLES), paragraphs=payload, profile=profile)
    last_err = None
    for attempt in range(max_retries + 1):
        try:
            items = llm.chat_json(prompt)
            return _validate_rolemap(items, expected)
        except (ValueError, KeyError, TypeError) as e:
            last_err = e
            if attempt < max_retries:
                on_event(f"角色标注未通过校验（{e}），正在要求模型重标")
                prompt = (PROMPT_TEMPLATE.format(
                    roles="/".join(BASE_ROLES), paragraphs=payload, profile=profile)
                          + RETRY_SUFFIX.format(error=e, idx_list=expected))
    raise ValueError(f"角色标注失败（重试 {max_retries} 次后放弃）: {last_err}")


def label_roles(paragraphs, llm, on_event=None, profile=None):
    """整篇段落清单 → {idx: role}。表格内段落（in_table=True）不送标注，直接标 other。"""
    on_event = on_event or (lambda msg: None)
    profile = profile or detect_document_profile(paragraphs)
    rolemap = {}
    todo = []
    for p in paragraphs:
        if not p.get("editable", True) or p.get("in_table"):
            rolemap[p["idx"]] = "other"
            continue
        hit = regex_role(p.get("text", ""), p, profile=profile)
        if hit:
            rolemap[p["idx"]] = hit  # 编号惯例命中的标题：确定性识别，不送 LLM
        else:
            todo.append(p)
    if rolemap:
        on_event(f"编号惯例直接识别 {len(rolemap)} 段（含表格跳过），其余 {len(todo)} 段送模型标注")
    n_batches = (len(todo) + BATCH_SIZE - 1) // BATCH_SIZE
    for i in range(0, len(todo), BATCH_SIZE):
        batch_no = i // BATCH_SIZE + 1
        if n_batches > 1:
            on_event(f"标注第 {batch_no}/{n_batches} 批段落（{len(todo[i:i + BATCH_SIZE])} 段）")
        rolemap.update(_label_batch(
            todo[i:i + BATCH_SIZE], llm, on_event=on_event, profile=profile))
    return rolemap


if __name__ == "__main__":
    import sys
    from core.extract import extract_paragraphs
    from core.llm import LLMClient
    paras = extract_paragraphs(sys.argv[1])
    rm = label_roles(paras, LLMClient())
    print(json.dumps(rm, ensure_ascii=False, indent=2))
