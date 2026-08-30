# 执行器接线（PLAN.md 第 7 节）：
# apply_format(docx_path, spec, rolemap, out_path) -> changelog
# 对每个段落按 RoleMap 取角色、从 FormatSpec 取规则，调用 core/executor.py 的
# 确定性函数改 XML；页边距/行网格走 section 级别。LLM 不碰 docx。

from docx import Document
from docx.shared import Mm

from core.executor import (
    set_doc_grid,
)
from core.style_set import (
    apply_named_style,
    clear_invalid_numbering_override,
    ensure_role_styles,
    resolve_target_body_style,
)
from core.track_changes import mark_paragraph_revision, snapshot_paragraph


def apply_format(docx_path, spec, rolemap, out_path, track=False):
    """应用 FormatSpec × RoleMap，输出 docx，返回 changelog list[dict]。
    rolemap: {idx: role}（idx 对应 extract.py 的段落序号）。
    模板未明确指定的角色统一与正文保持一致（套用 body 规则与样式）。
    表格内段落（idx >= len(doc.paragraphs)）v1 跳过。
    track=True 时输出修订模式文档：段落/字符格式改动以 w:pPrChange /
    w:rPrChange 记录，Word 审阅视图可见。
    """
    doc = Document(docx_path)
    roles = spec.get("roles", {})
    # 必须在创建/更新 FormatAgent 样式之前解析，避免把新样式误认成目标原样式。
    target_body_style = resolve_target_body_style(doc, rolemap)
    role_styles = ensure_role_styles(
        doc, spec, target_body_style=target_body_style)

    # ---- 页面级 ----
    page = spec.get("page") or {}
    margin = page.get("margin") or {}
    section = doc.sections[0]
    if margin.get("top_mm") is not None:
        section.top_margin = Mm(margin["top_mm"])
    if margin.get("bottom_mm") is not None:
        section.bottom_margin = Mm(margin["bottom_mm"])
    if margin.get("left_mm") is not None:
        section.left_margin = Mm(margin["left_mm"])
    if margin.get("right_mm") is not None:
        section.right_margin = Mm(margin["right_mm"])
    line_grid = page.get("line_grid") or {}
    if line_grid.get("line_pt") is not None:
        set_doc_grid(doc, line_pt=line_grid["line_pt"])

    # ---- 段落级 ----
    changelog = []
    rev_id = 1
    for idx, p in enumerate(doc.paragraphs):
        role = rolemap.get(idx, rolemap.get(str(idx)))
        if role is None:
            continue  # 未被标注的段落不动
        snapshot = snapshot_paragraph(p) if track else None
        if role in roles:
            rule = roles[role]
            style = role_styles[role]
            changed = apply_named_style(p, style, rule, role=role)
            fallback_to_target_body = False
        else:
            # 模板未规定的角色与正文保持一致：套用 body 的规则和命名样式，
            # 同时清掉会遮蔽样式的直接格式（如原文自带的加粗/异体字）。
            # 真实自动编号保留（body 不在清编号集合内），仅清“取消编号”残留。
            body_rule = roles.get("body", {})
            body_style = role_styles.get("body", target_body_style)
            changed = apply_named_style(p, body_style, body_rule, role="body")
            style = body_style
            fallback_to_target_body = True
        if track:
            rev_id = mark_paragraph_revision(p, snapshot, rev_id_start=rev_id)
        changelog.append({
            "idx": idx,
            "role": role,
            "style_name": style.name,
            "text": p.text.strip()[:30],
            "changed_fields": changed,
            "fallback_to_target_body": fallback_to_target_body,
        })

    doc.save(out_path)
    return changelog


def write_report(changelog, spec, report_path):
    """把 changelog 写成 markdown 修改对照报告。"""
    lines = ["# 排版修改对照报告", ""]
    page = spec.get("page") or {}
    if page:
        lines.append("## 页面设置")
        margin = page.get("margin") or {}
        if margin:
            lines.append(
                f"- 页边距（mm）：上 {margin.get('top_mm', '-')} / 下 {margin.get('bottom_mm', '-')}"
                f" / 左 {margin.get('left_mm', '-')} / 右 {margin.get('right_mm', '-')}")
        lg = page.get("line_grid") or {}
        if lg.get("line_pt"):
            lines.append(f"- 行网格：{lg['line_pt']} 磅/行")
        lines.append("")
    lines.append("## 段落修改明细")
    lines.append("")
    lines.append("| 段落 | 角色 | Word 样式 | 改动字段 | 内容摘要 |")
    lines.append("|---|---|---|---|---|")
    for c in changelog:
        fields = ", ".join(c["changed_fields"]) if c["changed_fields"] else "（无字段改动）"
        lines.append(
            f"| {c['idx']} | {c['role']} | {c.get('style_name', '-')} | {fields} | {c['text']} |")
    lines.append("")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return report_path
