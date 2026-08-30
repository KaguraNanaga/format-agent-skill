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
