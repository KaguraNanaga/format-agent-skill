# 通用格式排版 Agent — 执行器关键函数（赛前实测用）
# 直接跑通这三个函数，确认能改出正确 docx，再接入主流程。
# 只依赖 python-docx + lxml。环境：.venv-win7（python-docx 1.2.0）。

"""三处关键 XML 写法:
   - w:firstLineChars=200  (仓库已有先例: scripts/formatter.py)
   - w:spacing w:line / w:lineRule 单位是 1/20 磅  (Word 规范)
   - w:docGrid 行网格 (仓库无实现, 本文件提供)
   注意: 中文字体必须同时设置 w:ascii 和 w:eastAsia, python-docx 默认只设西文。
"""
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def set_run_fonts(run, eastasia=None, ascii_font=None, size_pt=None, bold=None):
    """设置 run 的中西文字体、字号、加粗。eastasia 是必须项，否则中文不生效。"""
    rpr = run._element.get_or_add_rPr()
    # 1) 字体（rFonts 各属性之间无顺序要求，OOXML 属性本就无序）
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    if ascii_font:
        rfonts.set(qn("w:ascii"), ascii_font)
        rfonts.set(qn("w:hAnsi"), ascii_font)
    if eastasia:
        rfonts.set(qn("w:eastAsia"), eastasia)
    # 2) 字号 (pt -> half-points)，w:sz 与 w:szCs 一起设
    if size_pt is not None:
        for tag in ("w:sz", "w:szCs"):
            sz = rpr.find(qn(tag))
            if sz is None:
                sz = OxmlElement(tag)
                rpr.append(sz)
            sz.set(qn("w:val"), str(int(round(size_pt * 2))))  # 1 pt = 2 half-point
    # 3) 加粗。bold=False 必须显式写 w:val="0"，不能只删元素——
    #    若样式层本身加粗（如标题样式），删掉直接格式也压不住。
    if bold is not None:
        for tag in ("w:b", "w:bCs"):
            b = rpr.find(qn(tag))
            if b is None:
                b = OxmlElement(tag)
                rpr.append(b)
            if bold:
                if b.get(qn("w:val")) is not None:
                    del b.attrib[qn("w:val")]  # 元素存在且无 val = 开
            else:
                b.set(qn("w:val"), "0")


def set_paragraph_fixed_spacing(paragraph, line_pt=None, before_pt=None, after_pt=None):
    """固定行距。line_pt 单位是磅, XML 中 w:line 单位是 1/20 磅。
    例: 28磅 -> line="560" lineRule="exact"。before/after 同理 1/20磅。
    """
    ppr = paragraph._p.get_or_add_pPr()
    spacing = ppr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        ppr.append(spacing)
    if line_pt is not None:
        spacing.set(qn("w:line"), str(int(round(line_pt * 20))))
        # "exact" 固定磅值; "atLeast" 最小; "auto" 多倍
        spacing.set(qn("w:lineRule"), "exact")
    if before_pt is not None:
        spacing.set(qn("w:before"), str(int(round(before_pt * 20))))
    if after_pt is not None:
        spacing.set(qn("w:after"), str(int(round(after_pt * 20))))


def set_first_line_indent_chars(paragraph, chars):
    """首行缩进按'字符' (2字符 = 200)。注意: 只有当字号已知时 Word 才显示为 X字符,
    所以要先设字号再设缩进, 或同时设置 w:firstLine (磅值) 作为回退。
    """
    ppr = paragraph._p.get_or_add_pPr()
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        ppr.append(ind)
    # firstLineChars 单位 = 1/100 字符 -> 2字符 = 200
    ind.set(qn("w:firstLineChars"), str(int(round(chars * 100))))
    # 同时给一个磅值回退 (用 2 字符 * 当前字号近似), 防止 Word 老版本只认磅值
    font_size = _effective_font_size(paragraph)
    if font_size:
        ind.set(qn("w:firstLine"), str(int(round(font_size * chars * 20))))


def set_outline_level(paragraph, level):
    """设置段落大纲级别（0-8，对应 Word 大纲 1-9 级），导航窗格据此显示结构。
    level=None 表示清除大纲级别（回归正文文本，不进导航窗格）。
    返回 True 表示有实际改动。

    注意 OOXML 中 pPr 的子元素有顺序要求：w:outlineLvl 必须位于 w:rPr 之前，
    乱序追加可能导致 Word 打开时提示"无法读取的内容"。
    """
    ppr = paragraph._p.get_or_add_pPr()
    ol = ppr.find(qn("w:outlineLvl"))
    if level is None:
        if ol is not None:
            ppr.remove(ol)
            return True
        return False
    if ol is not None and ol.get(qn("w:val")) == str(int(level)):
        return False
    if ol is None:
        ol = OxmlElement("w:outlineLvl")
        rpr = ppr.find(qn("w:rPr"))
        if rpr is not None:
            rpr.addprevious(ol)
        else:
            ppr.append(ol)
    ol.set(qn("w:val"), str(int(level)))
    return True


def set_doc_grid(document, line_pt=28.0):
    """开启页面行网格 (w:docGrid type="linesAndChars")，linePitch = 行距磅值 × 20。

    说明: 公文"每行 28 字"不靠 charSpace 硬设——A4 版心 (210-28-26=156mm≈442pt)
    配三号字 (16pt) 自然得约 28 字/行；每页行数由 linePitch 与版心高度决定。
    所以这里只设 type 和 linePitch，不伪造 charSpace，也不动 w:cols。
    前提: 页边距与正文字号已按 FormatSpec 设置。
    """
    section = document.sections[0]
    sect_pr = section._sectPr
    doc_grid = sect_pr.find(qn("w:docGrid"))
    if doc_grid is None:
        doc_grid = OxmlElement("w:docGrid")
        sect_pr.append(doc_grid)
    doc_grid.set(qn("w:type"), "linesAndChars")
    doc_grid.set(qn("w:linePitch"), str(int(round(line_pt * 20))))


def _effective_font_size(paragraph):
    """读取段落第一个 run 的生效字号(pt), 供 firstLine 磅值回退用。"""
    for run in paragraph.runs:
        rpr = run._element.find(qn("w:rPr"))
        if rpr is not None:
            for sz_tag in (qn("w:sz"), qn("w:szCs")):
                sz = rpr.find(sz_tag)
                if sz is not None and sz.get(qn("w:val")):
                    return int(sz.get(qn("w:val"))) / 2.0
    # 回退: 默认 16pt (三号)
    return 16.0
