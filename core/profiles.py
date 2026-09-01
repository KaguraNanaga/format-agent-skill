"""文档语言/文体 Profile 的确定性检测。"""

import re


PROFILES = {
    "general", "thesis",
    "english_general", "english_academic", "english_legal",
    "official_cn", "english_technical", "english_legal_brief",
}
ENGLISH_PROFILES = {
    "english_general", "english_academic", "english_legal",
    "english_technical", "english_legal_brief",
}

_CJK_RE = re.compile(r"[\u3400-\u9fff]")
_LATIN_RE = re.compile(r"[A-Za-z]")


def is_english_profile(profile):
    return profile in ENGLISH_PROFILES


def _english_dominant(texts):
    joined = " ".join(texts)
    latin = len(_LATIN_RE.findall(joined))
    cjk = len(_CJK_RE.findall(joined))
    return latin >= 20 and latin >= max(1, cjk * 2)


def detect_profile_from_texts(values):
    """从纯文本序列识别中文论文/公文及常见英文 Profile。"""
    texts = [str(value or "").strip() for value in values if str(value or "").strip()]
    compact = [text.replace(" ", "") for text in texts]

    chinese_signals = sum((
        any(text == "摘要" or re.match(r"^摘要[：:]", text) for text in compact),
        any(re.match(r"^(关键词|关键字)[：:]", text) for text in compact),
        any(text == "参考文献" for text in compact),
        any("学位论文" in text or "毕业论文" in text for text in compact),
    ))
    if chinese_signals >= 2 or (
        chinese_signals >= 1
        and any("学位论文" in text or "毕业论文" in text for text in compact)
    ):
        return "thesis"

    official_patterns = (
        r"〔\d{4}〕\d+号$",
        r"^(?:主送机关|抄送|签发人)[：:]",
        r"^(?:特此通知|特此报告|特此请示|特此函告)[。！]?$",
        r"^(?:通知|通报|请示|报告|批复|函|纪要)$",
        r"^附件[：:]",
    )
    official_signals = sum(
        1 for pattern in official_patterns
        if any(re.search(pattern, text) for text in compact)
    )
    if official_signals >= 2:
        return "official_cn"

    if not _english_dominant(texts):
        return "general"

    brief_patterns = (
        r"^IN THE .+ COURT\b",
        r"^(?:CASE\s+)?NO\.\s*[A-Z0-9:-]+",
        r"^TABLE OF AUTHORITIES$",
        r"^BRIEF OF\b",
        r"\bV\.?\s+\b",
        r"\b(?:APPELLANT|APPELLEE|PETITIONER|RESPONDENT)\b",
    )
    brief_signals = sum(
        1 for pattern in brief_patterns
        if any(re.search(pattern, text, re.I) for text in texts)
    )
    if brief_signals >= 2:
        return "english_legal_brief"

    technical_patterns = (
        r"^(?:WARNING|CAUTION|IMPORTANT|NOTE|TIP)\s*[:—-]",
        r"^(?:Step\s+\d+|Procedure\s+\d+)\b",
        r"^(?:GET|POST|PUT|PATCH|DELETE)\s+/\S+",
        r"^(?:\$|>|PS>)\s*\S+",
        r"^(?:Installation|Configuration|Troubleshooting|Prerequisites)$",
        r"\b(?:API|SDK|CLI)\b",
    )
    technical_signals = sum(
        1 for pattern in technical_patterns
        if any(re.search(pattern, text, re.I) for text in texts)
    )
    if technical_signals >= 2:
        return "english_technical"

    legal_patterns = (
        r"^ARTICLE\s+[IVXLC\d]+\b",
        r"^SECTION\s+\d+(?:\.\d+)*\b",
        r"^WHEREAS\b",
        r"\bIN WITNESS WHEREOF\b",
        r"\b(?:PLAINTIFF|DEFENDANT|APPELLANT|APPELLEE)\b",
        r"\b(?:AGREEMENT|INDENTURE|AFFIDAVIT)\b",
    )
    legal_signals = sum(
        1 for pattern in legal_patterns
        if any(re.search(pattern, text, re.I) for text in texts)
    )
    if legal_signals >= 2:
        return "english_legal"

    academic_patterns = (
        r"^Abstract\b", r"^Keywords?(?:\s*[:—-].*)?$", r"^References$",
        r"^(Bibliography|Works Cited)$",
        r"^(Introduction|Literature Review|Methods?|Methodology|Results?|Discussion|Conclusions?)$",
        r"\bdoi\s*:\s*10\.", r"^Author Note$", r"^Corresponding Author\b",
    )
    academic_signals = sum(
        1 for pattern in academic_patterns
        if any(re.search(pattern, text, re.I) for text in texts)
    )
    if academic_signals >= 2:
        return "english_academic"
    return "english_general"
