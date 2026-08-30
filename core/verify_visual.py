# 视觉验证（PLAN.md 6.3 + 第 9 节第 1 条）：
# render.py 渲染输出 docx → 6.3 prompt 送视觉模型 → 结构化问题清单
# [{role, field, pass, observed, expected}] → 代码只改 FormatSpec 对应字段，重跑一次。
# 不做开放循环：最多一轮定向修复。

import json
import os
import re

from core.render import render_docx_to_png
from core.schema import SpecValidationError, validate_spec

PROMPT_TEMPLATE = """你是排版质检员。对照检查清单逐条检查这份文档的渲染图。
本批图片对应整份文档的页码：{page_numbers}。
检查清单（来自排版规范）：
{checklist}
输出严格 JSON 对象（不得输出顶层数组）：
{{"status": "pass|fail|inconclusive", "pages_checked": [{page_numbers}],
  "reason": "无法判断时说明原因",
  "issues": [{{"role": "...", "field": "...", "pass": true,
               "observed": "图上看到的实际值",
               "expected": "清单要求值"}}]}}。
pass 必须是 JSON 布尔值 true/false，不能是字符串。
至少返回一个有把握的检查项；若图片空白、缺字或无法判断，
必须返回 status=inconclusive 并填写 reason，不得用空 issues 冒充通过。"""


class VisualVerificationError(RuntimeError):
    """视觉质检基类错误。"""


class VisualRenderError(VisualVerificationError):
    """DOCX 渲染或渲染质量检查失败。"""


class VisualModelError(VisualVerificationError):
    """图片已渲染，但多模态请求失败。"""


class VisualResponseError(VisualVerificationError):
    """模型返回的 JSON 结构不可用。"""


class VisualInconclusiveError(VisualVerificationError):
    """模型明确表示无法完成质检。"""

# 可自动修复的数值字段：issue.field -> (spec 内的取值路径, 合法范围)
_FIXABLE_NUMERIC = {
    "size_pt": (8, 72),
    "line_spacing.pt": (8, 72),
    "first_line_indent_chars": (0, 8),
}
_FIXABLE_ALIGN = {"left", "center", "right", "justify"}


def _build_checklist(spec):
    """把 FormatSpec 展开成人话检查清单。"""
    lines = []
    page = spec.get("page") or {}
    margin = page.get("margin") or {}
    if margin:
        lines.append(
            f"- 页面: 页边距 上{margin.get('top_mm')}/下{margin.get('bottom_mm')}"
            f"/左{margin.get('left_mm')}/右{margin.get('right_mm')} 毫米")
    lg = page.get("line_grid") or {}
    if lg.get("line_pt"):
        lines.append(f"- 页面: 行网格每行 {lg['line_pt']} 磅")
    for role, rule in (spec.get("roles") or {}).items():
        parts = []
        if rule.get("font_eastasia"):
            parts.append(f"中文字体 {rule['font_eastasia']}")
        if rule.get("size_pt"):
            parts.append(f"字号 {rule['size_pt']} 磅")
        if rule.get("bold") is not None:
            parts.append("加粗" if rule["bold"] else "不加粗")
        if rule.get("alignment"):
            parts.append(f"对齐 {rule['alignment']}")
        if rule.get("first_line_indent_chars"):
            parts.append(f"首行缩进 {rule['first_line_indent_chars']} 字符")
        if rule.get("space_before_pt"):
            parts.append(f"段前 {rule['space_before_pt']} 磅")
        if rule.get("space_after_pt"):
            parts.append(f"段后 {rule['space_after_pt']} 磅")
        ls = rule.get("line_spacing") or {}
        if ls.get("pt"):
            parts.append(f"行距 {'固定值' if ls.get('type') == 'exact' else '倍数'} {ls['pt']} 磅")
        if parts:
            lines.append(f"- 角色 {role}: " + "，".join(parts))
    return "\n".join(lines)


def _as_bool(value):
    if isinstance(value, bool):
        return value
    if isinstance(value, str) and value.strip().lower() in {"true", "false"}:
        return value.strip().lower() == "true"
    raise VisualResponseError(f"pass 必须是布尔值，收到 {value!r}")


def _validate_response(payload, expected_pages):
    """严格校验 VLM 对象；空结果绝不得被当成通过。"""
    if not isinstance(payload, dict):
        raise VisualResponseError("视觉模型必须返回 JSON 对象，不能返回顶层数组")
    status = str(payload.get("status") or "").strip().lower()
    if status not in {"pass", "fail", "inconclusive"}:
        raise VisualResponseError(f"非法的 status: {status!r}")

    raw_pages = payload.get("pages_checked")
    if not isinstance(raw_pages, list):
        raise VisualResponseError("缺少 pages_checked 数组")
    try:
        checked_pages = {int(page) for page in raw_pages}
    except (TypeError, ValueError) as exc:
        raise VisualResponseError("pages_checked 含有非整数页码") from exc
    missing_pages = set(expected_pages) - checked_pages
    if missing_pages:
        raise VisualResponseError(f"视觉响应未覆盖页码 {sorted(missing_pages)}")

    items = payload.get("issues")
    if not isinstance(items, list):
        raise VisualResponseError("缺少 issues 数组")
    out = []
    for it in items:
        if not isinstance(it, dict) or not {"role", "field", "pass"}.issubset(it):
            raise VisualResponseError(f"issues 条目结构不完整: {it!r}")
        out.append({
            "role": str(it["role"]),
            "field": str(it["field"]),
            "pass": _as_bool(it["pass"]),
            "observed": str(it.get("observed", "")),
            "expected": str(it.get("expected", "")),
        })

    reason = str(payload.get("reason") or "").strip()
    if status == "inconclusive":
        raise VisualInconclusiveError(reason or "视觉模型无法判断本批页面")
    if not out:
        raise VisualResponseError("视觉模型未返回任何检查项，不能视为通过")
    if status == "pass" and any(not item["pass"] for item in out):
        raise VisualResponseError("status=pass 但 issues 中包含失败项")
    if status == "fail" and all(item["pass"] for item in out):
        raise VisualResponseError("status=fail 但 issues 中没有失败项")
    return out


def _positive_int_env(name, default, minimum=1):
    try:
        return max(minimum, int(os.environ.get(name, str(default))))
    except ValueError:
        return default


def _image_batches(paths):
    """按页数和原始 PNG 字节数分批，避免长文档一次 base64 超限。"""
    max_pages = _positive_int_env("LLM_VISION_PAGES_PER_REQUEST", 3)
    max_bytes = _positive_int_env("LLM_VISION_MAX_BYTES", 5_000_000, minimum=100_000)
    batch = []
    batch_bytes = 0
    for page_no, path in enumerate(paths, 1):
        size = os.path.getsize(path)
        if batch and (len(batch) >= max_pages or batch_bytes + size > max_bytes):
            yield batch
            batch = []
            batch_bytes = 0
        batch.append((page_no, path))
        batch_bytes += size
    if batch:
        yield batch


def verify_visual(docx_path, spec, llm, png_dir, on_event=None):
    """渲染 + VLM 质检，返回结构化问题清单（全部检查项，含 pass=true）。"""
    on_event = on_event or (lambda _message: None)
    try:
        pages = render_docx_to_png(docx_path, png_dir, on_event=on_event)
    except Exception as exc:  # noqa: BLE001 —— 明确标记为渲染阶段
        raise VisualRenderError(str(exc)) from exc
    if not pages:
        raise VisualRenderError("文档渲染后没有生成任何页面")

    checklist = _build_checklist(spec)
    if not checklist.strip():
        raise VisualResponseError("FormatSpec 没有可供视觉检查的规则")
    all_items = []
    batches = list(_image_batches(pages))
    for batch_no, batch in enumerate(batches, 1):
        page_numbers = [page_no for page_no, _ in batch]
        prompt = PROMPT_TEMPLATE.format(
            checklist=checklist,
            page_numbers=", ".join(str(page) for page in page_numbers),
        )
        if len(batches) > 1:
            on_event(
                f"视觉质检第 {batch_no}/{len(batches)} 批（第 "
                f"{page_numbers[0]}-{page_numbers[-1]} 页）")
        try:
            payload = llm.chat_vision_json(prompt, [path for _, path in batch])
        except Exception as exc:  # noqa: BLE001
            raise VisualModelError(
                f"第 {page_numbers[0]}-{page_numbers[-1]} 页请求失败: {exc}") from exc
        all_items.extend(_validate_response(payload, page_numbers))
    return all_items


def _parse_number(text):
    m = re.search(r"-?\d+(?:\.\d+)?", str(text))
    return float(m.group()) if m else None


def apply_fixes(spec, issues):
    """定向修复：只改 FormatSpec 里与失败项对应的字段，且仅在 expected 可解析、
    数值在合法范围内时动手。返回 (修复后的 spec, 实际应用的修复列表)。
    修完过一遍 schema 校验，修坏了就放弃自动修复、返回原 spec。
    """
    fixed = json.loads(json.dumps(spec))  # deep copy
    applied = []
    for it in issues:
        if it["pass"]:
            continue
        role, field = it["role"], it["field"]
        rule = (fixed.get("roles") or {}).get(role)
        if rule is None:
            continue
        if field in _FIXABLE_NUMERIC:
            lo, hi = _FIXABLE_NUMERIC[field]
            # 只能向规范的 expected 方向修正；绝不得把观测到的
            # 错误值 observed 固化回 FormatSpec。
            v = _parse_number(it["expected"])
            if v is None or not (lo <= v <= hi):
                continue
            if field == "line_spacing.pt":
                ls = rule.get("line_spacing")
                if isinstance(ls, dict):
                    ls["pt"] = v
                else:
                    continue
            else:
                if rule.get(field) == v:
                    continue
                rule[field] = v
            applied.append(it)
        elif field == "alignment" and it["expected"] in _FIXABLE_ALIGN:
            if rule.get("alignment") == it["expected"]:
                continue
            rule["alignment"] = it["expected"]
            applied.append(it)
    if applied:
        try:
            validate_spec(fixed)
        except SpecValidationError:
            return spec, []  # 修坏了，回退
    return fixed, applied


if __name__ == "__main__":
    import sys
    from core.llm import LLMClient
    with open(sys.argv[2], encoding="utf-8") as f:
        spec = json.load(f)
    issues = verify_visual(sys.argv[1], spec, LLMClient(), png_dir="out/verify_render")
    print(json.dumps(issues, ensure_ascii=False, indent=2))
