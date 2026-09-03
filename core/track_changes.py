# 修订模式（Word 审阅-修订）支持：
# 在应用格式修改前快照段落的 pPr 和每个 run 的 rPr，修改后以
# w:pPrChange / w:rPrChange 写回——Word 打开后"审阅→修订"可见格式化改动。
# 正文内容我们没有改动，所以只有属性级修订，没有 w:ins/w:del。

import copy
from datetime import datetime, timezone

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

DEFAULT_AUTHOR = "FormatAgent"


def snapshot_paragraph(paragraph):
    """记录修改前的段落属性与每个 run 的字符属性（深拷贝）。"""
    ppr = paragraph._p.pPr
    old_ppr = copy.deepcopy(ppr) if ppr is not None else None
    old_rprs = []
    for run in paragraph.runs:
        rpr = run._element.find(qn("w:rPr"))
        old_rprs.append(copy.deepcopy(rpr) if rpr is not None else None)
    return old_ppr, old_rprs


def _set_change_attrs(el, rev_id, author, date):
    el.set(qn("w:id"), str(rev_id))
    el.set(qn("w:author"), author)
    el.set(qn("w:date"), date)


def mark_paragraph_revision(paragraph, snapshot, author=DEFAULT_AUTHOR,
                            rev_id_start=1, date=None):
    """把快照写成修订标记。返回下一个可用的修订 id。
    w:pPrChange 作为 pPr 的最后一个子元素；w:rPrChange 同理放在 rPr 末尾。
    """
    old_ppr, old_rprs = snapshot
    date = date or datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    rev_id = rev_id_start

    ppr = paragraph._p.pPr
    if old_ppr is not None and ppr is not None:
        # 旧 pPr 里如果本身已有 pPrChange（原文档自带修订），先剥掉，避免嵌套
        for nested in old_ppr.findall(qn("w:pPrChange")):
            old_ppr.remove(nested)
        change = OxmlElement("w:pPrChange")
        _set_change_attrs(change, rev_id, author, date)
        change.append(old_ppr)
        ppr.append(change)  # pPrChange 在 pPr 序列里排最后
        rev_id += 1

    for run, old_rpr in zip(paragraph.runs, old_rprs):
        if old_rpr is None:
            continue
        for nested in old_rpr.findall(qn("w:rPrChange")):
            old_rpr.remove(nested)
        rpr = run._element.find(qn("w:rPr"))
        if rpr is None:
            # 修改后 rPr 被清空了：补一个只带修订记录的 rPr（rPr 必须是 run 的首子元素）
            rpr = OxmlElement("w:rPr")
            run._element.insert(0, rpr)
        change = OxmlElement("w:rPrChange")
        _set_change_attrs(change, rev_id, author, date)
        change.append(old_rpr)
        rpr.append(change)  # rPrChange 在 rPr 序列里排最后
        rev_id += 1
    return rev_id


# WPS 兼容：修订稿如果让“当前格式”完全依赖命名样式继承（run 里没有直接
# 格式），WPS 在接受修订后无法解析字体，整篇回退到默认字体。Word 自己录制
# 格式修订时本来就会把新属性显式写成直接格式，因此这里在保留 pStyle 的前提
# 下，把样式的关键属性摊平到段落与 run 的直接格式里。只影响修订稿产物；
# 干净稿仍然只写命名样式。

# CT_RPr / CT_PPr 的 Schema 子元素顺序（节选），修订记录恒在最后。
_RPR_ORDER = (
    "w:rStyle", "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:caps",
    "w:smallCaps", "w:strike", "w:color", "w:spacing", "w:w",
    "w:kern", "w:position", "w:sz", "w:szCs", "w:highlight", "w:u",
    "w:vertAlign", "w:lang",
)
_PPR_ORDER = (
    "w:pStyle", "w:keepNext", "w:keepLines", "w:pageBreakBefore",
    "w:widowControl", "w:numPr", "w:pBdr", "w:shd", "w:tabs",
    "w:kinsoku", "w:wordWrap", "w:overflowPunct", "w:topLinePunct",
    "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd",
    "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing",
    "w:jc", "w:textDirection", "w:textAlignment", "w:outlineLvl",
    "w:rPr",
)
# 编号（numPr）与大纲级别（outlineLvl）由样式与编号保留逻辑负责，不摊平。
_PPR_FLATTEN = ("w:spacing", "w:ind", "w:jc", "w:contextualSpacing")
_RPR_FLATTEN = (
    "w:rFonts", "w:b", "w:bCs", "w:i", "w:iCs", "w:color",
    "w:sz", "w:szCs", "w:highlight", "w:lang",
)


def _ordered_insert(parent, new_el, order):
    names = [qn(tag) for tag in order]
    try:
        new_pos = names.index(new_el.tag)
    except ValueError:
        new_pos = len(names)
    for child in parent:
        if child.tag in (qn("w:rPrChange"), qn("w:pPrChange")):
            child.addprevious(new_el)
            return
        try:
            child_pos = names.index(child.tag)
        except ValueError:
            child_pos = len(names)
        if child_pos > new_pos:
            child.addprevious(new_el)
            return
    parent.append(new_el)


def _merge_style_rpr(parent, style_rpr, is_run):
    """把样式字符属性补进 parent 的 w:rPr；parent 没有 rPr 且确有属性可补时新建。
    is_run=True 时 parent 是 w:r（rPr 必须是首子元素），否则 parent 是 w:pPr。
    """
    if style_rpr is None:
        return
    wanted = [style_rpr.find(qn(tag)) for tag in _RPR_FLATTEN]
    wanted = [el for el in wanted if el is not None]
    if not wanted:
        return
    target_rpr = parent.find(qn("w:rPr"))
    if target_rpr is None:
        target_rpr = OxmlElement("w:rPr")
        if is_run:
            parent.insert(0, target_rpr)
        else:
            _ordered_insert(parent, target_rpr, _PPR_ORDER)
    existing = {child.tag for child in target_rpr}
    for src in wanted:
        if src.tag in existing:
            continue
        _ordered_insert(target_rpr, copy.deepcopy(src), _RPR_ORDER)
        existing.add(src.tag)


def flatten_style_formatting(paragraph, style_element):
    """把样式（w:style 元素）的关键格式摊平成段落的当前直接格式。

    已存在的直接格式优先（如保留的编号、有意加粗），只补缺项；
    w:pPrChange / w:rPrChange 修订记录不受影响。
    """
    if style_element is None:
        return
    style_ppr = style_element.find(qn("w:pPr"))
    style_rpr = style_element.find(qn("w:rPr"))
    ppr = paragraph._p.pPr
    if ppr is None:
        return
    if style_ppr is not None:
        existing = {child.tag for child in ppr}
        for tag in _PPR_FLATTEN:
            qtag = qn(tag)
            if qtag in existing:
                continue
            src = style_ppr.find(qtag)
            if src is None:
                continue
            _ordered_insert(ppr, copy.deepcopy(src), _PPR_ORDER)
            existing.add(qtag)
    _merge_style_rpr(ppr, style_rpr, is_run=False)  # 段落标记
    for run in paragraph.runs:
        _merge_style_rpr(run._element, style_rpr, is_run=True)
