"""技术手册的保守增强：只绑定现有图形与相邻题注，不移动对象。"""

from docx.oxml.ns import qn

from core.extract import iter_main_paragraphs


def _role(rolemap, index):
    return rolemap.get(index, rolemap.get(str(index)))


def _has_drawing(paragraph):
    return any(
        element.tag in {qn("w:drawing"), qn("w:pict")}
        for element in paragraph._p.iter()
    )


def apply_technical_features(document, spec, rolemap):
    """对图形段与相邻 figure_caption 设置分页绑定，返回可审计诊断。"""
    config = spec.get("technical") or {}
    enabled = spec.get("profile") == "english_technical" or bool(config)
    if not enabled:
        return {"changed_fields": [], "diagnostics": []}

    top_level = [
        (index, paragraph)
        for index, (paragraph, table_depth) in enumerate(iter_main_paragraphs(document))
        if table_depth == 0
    ]
    position = {index: offset for offset, (index, _) in enumerate(top_level)}
    bindings = 0
    diagnostics = []
    for index, paragraph in top_level:
        if not _has_drawing(paragraph):
            continue
        offset = position[index]
        before = top_level[offset - 1] if offset > 0 else None
        after = top_level[offset + 1] if offset + 1 < len(top_level) else None
        caption = None
        placement = None
        if after and _role(rolemap, after[0]) == "figure_caption":
            caption, placement = after[1], "after"
            paragraph.paragraph_format.keep_with_next = True
        elif before and _role(rolemap, before[0]) == "figure_caption":
            caption, placement = before[1], "before"
            caption.paragraph_format.keep_with_next = True
        if caption is None:
            if config.get("validate_figure_bindings", True):
                diagnostics.append({
                    "code": "UNBOUND_FIGURE",
                    "paragraph_index": index,
                    "message": "图形段前后没有已识别的 figure_caption；未移动图形。",
                })
            continue
        paragraph.paragraph_format.keep_together = True
        caption.paragraph_format.keep_together = True
        bindings += 1
        diagnostics.append({
            "code": "FIGURE_CAPTION_BOUND",
            "paragraph_index": index,
            "placement": placement,
        })

    changed = [f"figure_caption_bindings_{bindings}"] if bindings else []
    return {"changed_fields": changed, "diagnostics": diagnostics}
