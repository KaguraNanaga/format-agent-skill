"""DOCX 全 Story 扫描与排版能力预检。

扫描不修改文件。它覆盖正文、页眉页脚、脚注尾注、批注、文本框、内容控件、
修订、嵌入对象等 OOXML Story/结构，明确区分“能排版”和“只会原样保留”。
"""

from copy import deepcopy
from collections import Counter
import os
import re
import zipfile
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
PR_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"w": W_NS, "r": R_NS, "m": M_NS, "wp": WP_NS}

_FIELD_KINDS = (
    "BIBLIOGRAPHY", "CITATION", "NOTEREF", "PAGEREF", "STYLEREF",
    "TOA", "TA", "INDEX", "XE", "TOC", "SEQ", "REF", "PAGE",
)
_FIELD_PATTERN = re.compile(
    r"\b(" + "|".join(_FIELD_KINDS) + r")\b(?:\s+([^\s\\]+))?",
    re.I,
)
_FIELD_KIND_PATTERN = re.compile(r"\s*([A-Za-z][A-Za-z0-9]*)")
_UNSAFE_FIELD_KINDS = {
    "DDE", "DDEAUTO", "LINK", "INCLUDETEXT", "INCLUDEPICTURE",
    "DATABASE", "RD",
}


class PreflightError(RuntimeError):
    pass


class PreflightBlockedError(PreflightError):
    def __init__(self, report):
        self.report = report
        codes = ", ".join(item["code"] for item in report.get("blockers", []))
        super().__init__(f"能力预检未通过：{codes or '存在阻断项'}")


def _part_kind(name):
    if name == "word/document.xml":
        return "main"
    if name.startswith("word/header") and name.endswith(".xml"):
        return "header"
    if name.startswith("word/footer") and name.endswith(".xml"):
        return "footer"
    mapping = {
        "word/footnotes.xml": "footnotes",
        "word/endnotes.xml": "endnotes",
        "word/comments.xml": "comments",
    }
    return mapping.get(name)


def _count(root, path):
    return len(root.findall(path, NS))


def _onoff_true(value):
    return str(value or "0").strip().lower() in {"1", "true", "on"}


def _nested_table_count(root):
    nested = set()
    for table in root.findall(".//w:tbl", NS):
        nested.update(table.findall(".//w:tbl", NS))
    return len(nested)


def _section_inventory(main):
    sections = []
    for index, sect_pr in enumerate(main.findall(".//w:sectPr", NS)):
        page_size = sect_pr.find("w:pgSz", NS)
        margin = sect_pr.find("w:pgMar", NS)
        columns = sect_pr.find("w:cols", NS)
        numbering = sect_pr.find("w:pgNumType", NS)
        section_type = sect_pr.find("w:type", NS)
        sections.append({
            "section_index": index,
            "type": section_type.get(f"{{{W_NS}}}val") if section_type is not None else "nextPage",
            "page_width_twips": int(page_size.get(f"{{{W_NS}}}w"))
            if page_size is not None and (page_size.get(f"{{{W_NS}}}w") or "").isdigit() else None,
            "page_height_twips": int(page_size.get(f"{{{W_NS}}}h"))
            if page_size is not None and (page_size.get(f"{{{W_NS}}}h") or "").isdigit() else None,
            "orientation": page_size.get(f"{{{W_NS}}}orient", "portrait")
            if page_size is not None else None,
            "margin_twips": {
                edge: int(margin.get(f"{{{W_NS}}}{edge}"))
                for edge in ("top", "bottom", "left", "right")
                if margin is not None and (margin.get(f"{{{W_NS}}}{edge}") or "").isdigit()
            },
            "columns": int(columns.get(f"{{{W_NS}}}num", "1"))
            if columns is not None and (columns.get(f"{{{W_NS}}}num", "1") or "").isdigit() else 1,
            "page_number_format": numbering.get(f"{{{W_NS}}}fmt")
            if numbering is not None else None,
            "page_number_start": int(numbering.get(f"{{{W_NS}}}start"))
            if numbering is not None and (numbering.get(f"{{{W_NS}}}start") or "").isdigit() else None,
            "header_references": len(sect_pr.findall("w:headerReference", NS)),
            "footer_references": len(sect_pr.findall("w:footerReference", NS)),
            "different_first_page": sect_pr.find("w:titlePg", NS) is not None,
        })
    return sections


def _field_data(root):
    # Word 会把一条域指令任意拆到多个 instrText 节点中；按 fldChar
    # begin/end 边界重组，避免 REF 与书签名被拆开后漏报。
    instructions = []
    field_stack = []
    for element in root.iter():
        if element.tag == f"{{{W_NS}}}fldChar":
            kind = element.get(f"{{{W_NS}}}fldCharType")
            if kind == "begin":
                field_stack.append([])
            elif kind == "end" and field_stack:
                chunks = field_stack.pop()
                instruction = "".join(chunks).strip()
                if instruction:
                    instructions.append(instruction)
        elif element.tag == f"{{{W_NS}}}instrText":
            text = element.text or ""
            if field_stack:
                field_stack[-1].append(text)
            elif text.strip():
                # 容忍缺少 fldChar 包裹的非标准生成器输出。
                instructions.append(text.strip())
    for chunks in field_stack:
        instruction = "".join(chunks).strip()
        if instruction:
            instructions.append(instruction)
    instructions.extend(
        (element.get(f"{{{W_NS}}}instr") or "").strip()
        for element in root.findall(".//w:fldSimple", NS)
        if (element.get(f"{{{W_NS}}}instr") or "").strip()
    )
    field_types = Counter()
    all_field_types = Counter()
    cross_reference_targets = []
    for instruction in instructions:
        first = _FIELD_KIND_PATTERN.match(instruction)
        if first:
            all_field_types[first.group(1).upper()] += 1
        for match in _FIELD_PATTERN.finditer(instruction):
            kind = match.group(1).upper()
            field_types[kind] += 1
            if kind in {"REF", "PAGEREF", "NOTEREF"} and match.group(2):
                cross_reference_targets.append(match.group(2).strip('"'))
    bookmarks = {
        element.get(f"{{{W_NS}}}name")
        for element in root.findall(".//w:bookmarkStart", NS)
        if element.get(f"{{{W_NS}}}name")
    }
    return {
        "instructions": instructions,
        "field_types": dict(sorted(field_types.items())),
        "all_field_types": dict(sorted(all_field_types.items())),
        "cross_reference_targets": cross_reference_targets,
        "bookmarks": sorted(bookmarks),
    }


def _story_stats(name, root):
    kind = _part_kind(name)
    entries = 0
    if kind in {"footnotes", "endnotes"}:
        item_tag = "footnote" if kind == "footnotes" else "endnote"
        for item in root.findall(f".//w:{item_tag}", NS):
            item_type = item.get(f"{{{W_NS}}}type")
            item_id = item.get(f"{{{W_NS}}}id")
            # Word 自带 separator / continuationSeparator 不是用户注释。
            if item_type in {"separator", "continuationSeparator"}:
                continue
            try:
                if item_id is not None and int(item_id) <= 0:
                    continue
            except ValueError:
                pass
            entries += 1
    elif kind == "comments":
        entries = _count(root, ".//w:comment")
    field_data = _field_data(root)
    return {
        "part": name,
        "kind": kind,
        "entries": entries,
        "paragraphs": _count(root, ".//w:p"),
        "tables": _count(root, ".//w:tbl"),
        "nested_tables": _nested_table_count(root),
        "merged_cells": _count(root, ".//w:gridSpan") + _count(root, ".//w:vMerge"),
        "textboxes": _count(root, ".//w:txbxContent"),
        "content_controls": _count(root, ".//w:sdt"),
        "tracked_changes": sum(
            _count(root, f".//w:{tag}")
            for tag in ("ins", "del", "moveFrom", "moveTo")
        ),
        "drawings": _count(root, ".//w:drawing") + _count(root, ".//w:pict"),
        "floating_drawings": _count(root, ".//wp:anchor"),
        "inline_drawings": _count(root, ".//wp:inline"),
        # oMathPara contains oMath children, so counting both double-counts
        # displayed equations.  oMath also covers inline equations.
        "math_objects": _count(root, ".//m:oMath"),
        "fields": len(field_data["instructions"]),
        "field_types": field_data["field_types"],
        "all_field_types": field_data["all_field_types"],
        "field_instructions": field_data["instructions"],
        "cross_reference_targets": field_data["cross_reference_targets"],
        "bookmarks": field_data["bookmarks"],
        "hyperlinks": _count(root, ".//w:hyperlink"),
        "alt_chunks": _count(root, ".//w:altChunk"),
    }


def _risk(code, severity, message, count=1, action="preserve_only"):
    return {
        "code": code,
        "severity": severity,
        "message": message,
        "count": int(count),
        "action": action,
    }


def scan_docx(docx_path):
    """只读扫描 DOCX 包，返回完整 Story/复杂结构清单。"""
    path = os.path.abspath(docx_path)
    if os.path.splitext(path)[1].lower() != ".docx":
        raise PreflightError("当前只支持标准 .docx；不接受 .doc/.docm/.rtf/.odt/.wps")
    if not os.path.isfile(path):
        raise PreflightError(f"目标文档不存在：{path}")

    try:
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
            if "word/document.xml" not in names:
                raise PreflightError("文件不是有效的 Word DOCX：缺少 word/document.xml")
            stories = []
            roots = {}
            for name in sorted(names):
                if _part_kind(name) is None:
                    continue
                try:
                    root = ET.fromstring(archive.read(name))
                except ET.ParseError as exc:
                    raise PreflightError(f"OOXML 部件无法解析：{name}: {exc}") from exc
                roots[name] = root
                stories.append(_story_stats(name, root))

            main = roots["word/document.xml"]
            settings = None
            if "word/settings.xml" in names:
                settings = ET.fromstring(archive.read("word/settings.xml"))
            external_relationship_types = Counter()
            for name in sorted(
                item for item in names if item.endswith(".rels")
            ):
                try:
                    relationships = ET.fromstring(archive.read(name))
                except ET.ParseError as exc:
                    raise PreflightError(
                        f"OOXML 关系部件无法解析：{name}: {exc}") from exc
                for relationship in relationships.findall(
                    f"{{{PR_NS}}}Relationship"
                ):
                    if (relationship.get("TargetMode") or "").lower() != "external":
                        continue
                    rel_type = (relationship.get("Type") or "unknown").rsplit("/", 1)[-1]
                    external_relationship_types[rel_type] += 1
            scan = {
                "path": path,
                "file_size": os.path.getsize(path),
                "story_parts": stories,
                "story_counts": {},
                "section_count": _count(main, ".//w:sectPr"),
                "sections": _section_inventory(main),
                "top_level_table_count": _count(main, "./w:body/w:tbl"),
                "embedded_objects": sum(
                    1 for name in names if name.startswith("word/embeddings/")),
                "media_files": sum(
                    1 for name in names if name.startswith("word/media/")),
                "external_relationship_count": sum(
                    external_relationship_types.values()),
                "external_relationship_types": dict(
                    sorted(external_relationship_types.items())),
                "has_macros": "word/vbaProject.bin" in names,
                "has_digital_signatures": any(
                    name.startswith("_xmlsignatures/") for name in names),
                # WPS 会在未保护文档中写入 enforcement="0" 的占位元素；
                # 仅元素存在并不代表编辑保护已启用。
                "has_document_protection": bool(
                    settings is not None
                    and (protection := settings.find(
                        "w:documentProtection", NS)) is not None
                    and _onoff_true(protection.get(
                        f"{{{W_NS}}}enforcement"))),
            }
    except zipfile.BadZipFile as exc:
        raise PreflightError("文件不是有效的 DOCX/ZIP 包") from exc

    kinds = {story["kind"] for story in scan["story_parts"]}
    for kind in sorted(kinds):
        items = [story for story in scan["story_parts"] if story["kind"] == kind]
        scan["story_counts"][kind] = {
            key: sum(item[key] for item in items)
            for key in (
                "entries", "paragraphs", "tables", "textboxes", "content_controls",
                "nested_tables", "merged_cells", "tracked_changes", "drawings",
                "floating_drawings", "inline_drawings", "math_objects",
                "fields", "hyperlinks", "alt_chunks",
            )
        }
        scan["story_counts"][kind]["parts"] = len(items)
    field_inventory = Counter()
    all_field_inventory = Counter()
    bookmarks = set()
    cross_reference_targets = []
    for story in scan["story_parts"]:
        field_inventory.update(story.get("field_types") or {})
        all_field_inventory.update(story.get("all_field_types") or {})
        bookmarks.update(story.get("bookmarks") or [])
        cross_reference_targets.extend(
            story.get("cross_reference_targets") or [])
    scan["field_inventory"] = dict(sorted(field_inventory.items()))
    scan["all_field_inventory"] = dict(sorted(all_field_inventory.items()))
    scan["unsafe_field_inventory"] = {
        kind: count for kind, count in sorted(all_field_inventory.items())
        if kind in _UNSAFE_FIELD_KINDS
    }
    scan["bookmark_names"] = sorted(bookmarks)
    scan["cross_reference_targets"] = cross_reference_targets
    scan["broken_cross_references"] = sorted({
        target for target in cross_reference_targets
        if target not in bookmarks
    })
    return scan


def assess_preflight(
    scan, spec=None, allow_risky_structure=False, template_source=False,
):
    """结合 FormatSpec 评估风险；阻断项默认禁止继续排版。"""
    report = deepcopy(scan)
    spec = spec or {}
    risks = []
    totals = {
        key: sum(story[key] for story in scan.get("story_parts", []))
        for key in (
            "textboxes", "content_controls", "tracked_changes", "drawings",
            "nested_tables", "merged_cells", "floating_drawings", "math_objects",
            "fields", "alt_chunks",
        )
    }
    story_counts = scan.get("story_counts", {})
    for kind, label in (
        ("footnotes", "脚注"), ("endnotes", "尾注"), ("comments", "批注"),
    ):
        count = (story_counts.get(kind) or {}).get("entries", 0)
        if count:
            note_key = {"footnotes": "footnote", "endnotes": "endnote"}.get(kind)
            note_rule = (spec.get("notes") or {}).get(note_key) if note_key else None
            if note_rule:
                risks.append(_risk(
                    f"FORMAT_{kind.upper()}", "warning",
                    f"检测到{label}；将应用注释文字规则并保留编号、引用关系及复杂结构",
                    count, "format_story"))
            else:
                risks.append(_risk(
                    f"PRESERVE_ONLY_{kind.upper()}", "warning",
                    f"检测到{label}；未提供注释规则，保持原样",
                    count))
    for key, label in (
        ("textboxes", "文本框"), ("content_controls", "内容控件"),
        ("drawings", "绘图/浮动对象"),
    ):
        if totals[key]:
            risks.append(_risk(
                f"PRESERVE_ONLY_{key.upper()}", "warning",
                f"检测到{label}；当前不会按角色重排，只保证尽量保留",
                totals[key]))
    if totals["nested_tables"]:
        risks.append(_risk(
            "PRESERVE_ONLY_NESTED_TABLES", "warning",
            "检测到嵌套表格；会纳入阅读顺序和文字完整性校验，但不推断其语义角色",
            totals["nested_tables"]))
    if totals["merged_cells"]:
        risks.append(_risk(
            "COMPLEX_TABLE_GEOMETRY", "warning",
            "检测到跨行/跨列单元格；只调整显式表格规则并保留合并关系",
            totals["merged_cells"]))
    if totals["math_objects"]:
        risks.append(_risk(
            "PRESERVE_ONLY_MATH_OBJECTS", "warning",
            "检测到 Word 公式对象；当前保留公式 XML，只调整承载段落样式",
            totals["math_objects"]))
    if totals["fields"]:
        risks.append(_risk(
            "FIELD_REFRESH_REQUIRED", "warning",
            "检测到 Word 域；保留域代码并校验交叉引用目标，域结果需在 Word 中刷新",
            totals["fields"], "refresh_fields"))
    if scan.get("unsafe_field_inventory"):
        kinds = ", ".join(sorted(scan["unsafe_field_inventory"]))
        risks.append(_risk(
            "UNSAFE_EXTERNAL_FIELDS", "warning",
            f"检测到可能访问外部资源的域（{kinds}）；自动域刷新将跳过这些域",
            sum(scan["unsafe_field_inventory"].values()), "skip_refresh"))
    if scan.get("external_relationship_count"):
        risks.append(_risk(
            "EXTERNAL_RELATIONSHIPS", "warning",
            "检测到外部超链接或外部资源关系；排版器只保留关系，不主动访问目标",
            scan["external_relationship_count"], "preserve_only"))
    if scan.get("broken_cross_references"):
        risks.append(_risk(
            "BROKEN_CROSS_REFERENCES", "warning",
            "检测到找不到书签目标的 REF/PAGEREF/NOTEREF 域，请人工修复交叉引用",
            len(scan["broken_cross_references"]), "manual_review"))
    if totals["tracked_changes"]:
        risks.append(_risk(
            "SOURCE_TRACKED_CHANGES", "blocker",
            "源文档含未处理的修订；必须先接受/拒绝修订，避免正文读取遗漏",
            totals["tracked_changes"], "stop"))
    if totals["alt_chunks"]:
        risks.append(_risk(
            "ALT_CHUNK_CONTENT", "blocker",
            "源文档含外部导入内容 altChunk，python-docx 无法安全处理",
            totals["alt_chunks"], "stop"))
    if scan.get("embedded_objects"):
        risks.append(_risk(
            "EMBEDDED_OBJECTS", "warning",
            "检测到嵌入对象；不会调整其内部内容",
            scan["embedded_objects"]))
    if scan.get("has_document_protection"):
        severity = "warning" if template_source else "blocker"
        risks.append(_risk(
            "DOCUMENT_PROTECTION", severity,
            "参考模板启用了编辑保护；只读提取可能遗漏受保护结构"
            if template_source else "目标文档启用了编辑保护，拒绝在未解除保护时写入",
            action="preserve_only" if template_source else "stop"))
    if scan.get("has_digital_signatures"):
        severity = "warning" if template_source else "blocker"
        risks.append(_risk(
            "DIGITAL_SIGNATURE", severity,
            "参考模板包含数字签名；模板保持只读，不复制签名"
            if template_source else "目标文档包含数字签名，任何保存都会使签名失效",
            action="preserve_only" if template_source else "stop"))
    if scan.get("has_macros"):
        severity = "warning" if template_source else "blocker"
        risks.append(_risk(
            "MACRO_CONTENT", severity,
            "参考模板包含宏工程；只提取常规 Word 格式，不复制宏"
            if template_source else "目标 DOCX 包含宏工程；当前流水线不保真处理宏",
            action="preserve_only" if template_source else "stop"))

    structure = spec.get("structure") or {}
    page = spec.get("page") or {}
    columns = page.get("columns")
    landscape_tables = (spec.get("table") or {}).get("landscape_table_indices")
    section_overrides = page.get("section_overrides") or []
    invalid_sections = sorted({
        item.get("section_index") for item in section_overrides
        if isinstance(item, dict) and isinstance(item.get("section_index"), int)
        and item["section_index"] >= scan.get("section_count", 1)
    })
    if invalid_sections:
        risks.append(_risk(
            "SECTION_OVERRIDE_OUT_OF_RANGE", "blocker",
            f"规范点名的节不存在：{invalid_sections}；不能把模板分节安全映射到目标稿",
            len(invalid_sections), "stop"))
    table_overrides = (spec.get("table") or {}).get("overrides") or []
    invalid_tables = sorted({
        item.get("table_index") for item in table_overrides
        if isinstance(item, dict) and isinstance(item.get("table_index"), int)
        and item["table_index"] >= scan.get("top_level_table_count", 0)
    })
    if invalid_tables:
        risks.append(_risk(
            "TABLE_OVERRIDE_OUT_OF_RANGE", "blocker",
            f"规范点名的顶层表格不存在：{invalid_tables}",
            len(invalid_tables), "stop"))
    if scan.get("section_count", 1) > 1 and (
        structure.get("enabled") or (isinstance(columns, int) and columns > 1)
        or landscape_tables
    ):
        severity = "warning" if allow_risky_structure else "blocker"
        risks.append(_risk(
            "MULTI_SECTION_STRUCTURAL_REBUILD", severity,
            "源文档已有多节，而规范要求重建论文结构、分栏或新增横向表格节；"
            "默认拒绝删除/重排原分节",
            scan["section_count"],
            "explicit_override" if not allow_risky_structure else "approved_override"))
    if (
        scan.get("section_count", 1) > 1
        and (page.get("size") or page.get("orientation") or page.get("margin"))
        and not allow_risky_structure
    ):
        risks.append(_risk(
            "MULTI_SECTION_GEOMETRY_PRESERVED", "warning",
            "源文档已有多节；全局纸张、方向和页边距不会覆盖各节，"
            "请用 page.section_overrides 点名需要调整的节",
            scan["section_count"], "preserve_sections"))

    report["risks"] = risks
    report["blockers"] = [item for item in risks if item["severity"] == "blocker"]
    report["warnings"] = [item for item in risks if item["severity"] == "warning"]
    report["ok"] = not report["blockers"]
    return report


def preflight_docx(
    docx_path, spec=None, allow_risky_structure=False, template_source=False,
):
    return assess_preflight(
        scan_docx(docx_path), spec=spec,
        allow_risky_structure=allow_risky_structure,
        template_source=template_source)


def merge_preflight_reports(target_report, template_report=None):
    """合并目标稿与可选模板的风险，同时保留两份原始扫描明细。"""
    result = deepcopy(target_report)
    target_risks = result.get("risks", [])
    for risk in target_risks:
        risk["source"] = "target"
    result["risks"] = target_risks
    result["blockers"] = [
        risk for risk in target_risks if risk.get("severity") == "blocker"
    ]
    result["warnings"] = [
        risk for risk in target_risks if risk.get("severity") == "warning"
    ]
    if template_report is not None:
        result["template"] = deepcopy(template_report)
        template_risks = deepcopy(template_report.get("risks", []))
        for risk in template_risks:
            risk["source"] = "template"
        result["risks"].extend(template_risks)
        result["blockers"].extend(
            risk for risk in template_risks
            if risk.get("severity") == "blocker"
        )
        result["warnings"].extend(
            risk for risk in template_risks
            if risk.get("severity") == "warning"
        )
    result["ok"] = not result["blockers"]
    return result


def raise_for_preflight(report):
    if not report.get("ok"):
        raise PreflightBlockedError(report)
