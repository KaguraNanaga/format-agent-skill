# Agent 编排器 —— 把整条流水线包装成带"工作日志"的事件流，供演示界面实时展示。
# 事件: {"step": 步骤名, "message": 人话描述, "status": run|ok|warn|err, "data": 任意}
# 演示故事: 理解归 AI，动手归代码，中间用 JSON 交接 —— 日志把这个过程直播出来。

import json
import os
from contextlib import ExitStack
from copy import deepcopy

from core.apply import apply_format, write_report
from core.extract import extract_paragraphs
from core.safe_output import (
    AtomicOutputSet,
    IntegrityViolationError,
    validate_output_paths,
)
from core.schema import validate_spec
from core.style_set import style_name_for_role


class Agent:
    """任务式排版 Agent：给它格式来源 + 目标文档，它自主完成理解→执行→自检。"""

    def __init__(self, llm=None, on_event=None):
        # on_event(event_dict)；llm 为 None 时，需要 LLM 的步骤才会延迟构造
        self._llm = llm
        self.on_event = on_event or (lambda event: None)

    def _emit(self, step, message, status="run", data=None):
        self.on_event({"step": step, "message": message, "status": status, "data": data})

    def _get_llm(self):
        if self._llm is None:
            from core.llm import LLMClient
            self._llm = LLMClient(on_event=lambda msg: self._emit("llm", msg, status="warn"))
        return self._llm

    def run(self, target_path, out_path, spec_text=None, spec=None,
            template_path=None, template_rolemap=None, rolemap=None,
            style_pack=None, style_pack_options=None,
            verify=False, report_path=None, cleanup_mode=None,
            refresh_fields=False, allow_risky_structure=False,
            _conversion_context=None):
        """跑完整流程，返回结果 dict（spec/rolemap/changelog/issues/paths）。"""
        report_path = report_path or os.path.splitext(out_path)[0] + "_report.md"
        base = os.path.splitext(out_path)[0]
        report_docx_path = base + "_report.docx"
        tracked_path = base + "_tracked.docx"
        final_paths = {
            "out_path": out_path,
            "report_path": report_path,
            "report_docx_path": report_docx_path,
            "tracked_path": tracked_path,
        }
        validate_output_paths(target_path, final_paths)
        if os.path.splitext(out_path)[1].lower() != ".docx":
            raise ValueError("输出必须使用 .docx 扩展名；旧格式只作为输入转换源。")

        # 旧 Word/WPS/ODT/RTF 只作为输入：先转入临时 DOCX，再让完整流水线
        # 对转换结果重新做 Story 预检和文本一致性验收。递归仅发生一次。
        if _conversion_context is None:
            from core.input_conversion import converted_input
            with ExitStack() as stack:
                converted_target = stack.enter_context(converted_input(target_path))
                converted_template = (
                    stack.enter_context(converted_input(template_path))
                    if template_path else None
                )
                conversion_context = {
                    "target": converted_target.as_dict(),
                    "template": (
                        converted_template.as_dict() if converted_template else None),
                }
                converted = [
                    value for value in (converted_target, converted_template)
                    if value is not None and value.lossy
                ]
                if converted:
                    self._emit(
                        "输入转换",
                        "已将旧格式转换为临时 DOCX；将对转换结果执行完整预检与验收",
                        status="warn", data=conversion_context)
                result = self.run(
                    converted_target.docx_path, out_path,
                    spec_text=spec_text, spec=spec,
                    template_path=(
                        converted_template.docx_path if converted_template else None),
                    template_rolemap=template_rolemap, rolemap=rolemap,
                    style_pack=style_pack, style_pack_options=style_pack_options,
                    verify=verify, report_path=report_path,
                    cleanup_mode=cleanup_mode, refresh_fields=refresh_fields,
                    allow_risky_structure=allow_risky_structure,
                    _conversion_context=conversion_context,
                )
                result["input_conversion"] = conversion_context
                return result

        # ① 理解格式规范 → FormatSpec
        self._emit("理解规范", "开始理解格式来源，抽取格式规则 ...")
        if style_pack is not None:
            if spec is not None or spec_text is not None or template_path is not None:
                raise ValueError("Style Pack 不能与 spec/spec_text/template 同时提供")
            from core.style_packs import get_style_pack
            spec = get_style_pack(style_pack, **(style_pack_options or {}))
            self._emit(
                "理解规范", f"已加载 Style Pack：{style_pack}", status="ok")
        elif spec is not None:
            if cleanup_mode is not None:
                spec = dict(spec)
                spec["cleanup"] = {"mode": cleanup_mode}
            validate_spec(spec)
            self._emit("理解规范", "FormatSpec 由用户直接给定（JSON），校验通过", status="ok")
        elif template_path is not None:
            from core.rules_from_template import extract_rules_from_template
            if template_rolemap is None:
                self._emit("理解规范", "正在解析模板文档结构，标注模板段落角色 ...")
                tpl_paras = extract_paragraphs(template_path)
                from core.label_roles import label_roles
                template_rolemap = label_roles(
                    tpl_paras, self._get_llm(),
                    on_event=lambda m: self._emit("理解规范", m))
            spec = extract_rules_from_template(template_path, template_rolemap)
            self._emit("理解规范",
                       f"已从模板确定性读取出 {len(spec['roles'])} 个角色的格式规则",
                       status="ok")
        elif spec_text is not None:
            from core.rules_from_text import extract_rules
            spec = extract_rules(
                spec_text, self._get_llm(),
                on_event=lambda m: self._emit("理解规范", m, status="warn"))
            self._emit("理解规范",
                       f"规范理解完成：识别出 {len(spec['roles'])} 个角色的格式规则",
                       status="ok")
        else:
            raise ValueError(
                "必须提供 spec_text / spec / template_path / style_pack 之一")

        if cleanup_mode is not None and (spec.get("cleanup") or {}).get("mode") != cleanup_mode:
            spec = dict(spec)
            spec["cleanup"] = {"mode": cleanup_mode}
            validate_spec(spec)

        if template_path and (spec.get("structure") or {}).get("enabled"):
            from core.thesis_structure import extract_cover_metadata
            metadata = extract_cover_metadata(target_path)
            if metadata:
                spec = deepcopy(spec)
                cover = spec.setdefault("structure", {}).setdefault("cover", {})
                cover.setdefault("metadata", {}).update(metadata)
                validate_spec(spec)
                self._emit(
                    "理解规范",
                    f"已从目标稿显式标签回填 {len(metadata)} 项封面元数据",
                    status="ok", data=metadata)

        # ①.5 全 Story/复杂结构能力预检：在任何输出写入前阻断已知危险输入。
        from core.preflight import (
            merge_preflight_reports, preflight_docx, raise_for_preflight,
        )
        self._emit("能力预检", "正在扫描目标稿与模板的正文、页眉页脚、脚注、文本框、修订和分节 ...")
        target_preflight = preflight_docx(
            target_path, spec=spec,
            allow_risky_structure=allow_risky_structure)
        template_preflight = (
            preflight_docx(template_path, template_source=True)
            if template_path else None
        )
        preflight = merge_preflight_reports(
            target_preflight, template_report=template_preflight)
        if preflight["warnings"]:
            self._emit(
                "能力预检",
                f"发现 {len(preflight['warnings'])} 类只保留、不重排的结构；将记录到预检结果",
                status="warn", data=preflight)
        if preflight["blockers"]:
            self._emit(
                "能力预检",
                f"发现 {len(preflight['blockers'])} 个硬阻断项，未创建或覆盖任何正式产物",
                status="err", data=preflight)
        raise_for_preflight(preflight)
        self._emit(
            "能力预检",
            f"预检通过：{preflight['section_count']} 节，"
            f"{len(preflight['story_parts'])} 个 Story 部件",
            status="ok", data=preflight)

        # ② 解析目标文档结构
        self._emit("解析文档", "正在解析目标文档结构 ...")
        paragraphs = extract_paragraphs(target_path)
        n_table = sum(1 for p in paragraphs if p["in_table"])
        n_protected = sum(
            1 for p in paragraphs if p.get("story", "main") != "main")
        self._emit("解析文档",
                   f"发现 {len(paragraphs)} 个段落记录（{n_table} 段在表格内，"
                   f"{n_protected} 段来自受保护 Story，均不参与正文重排）",
                   status="ok", data=paragraphs)

        # ③ 标注段落角色 → RoleMap
        if rolemap is not None:
            # 外部给定（宿主 Agent 自标）的 RoleMap 也要过校验：角色合法、非表格段全覆盖
            from core.schema import BASE_ROLES
            if not isinstance(rolemap, dict):
                raise ValueError("RoleMap 必须是 idx → role 的 JSON object")
            non_integer_keys = [
                key for key in rolemap
                if not isinstance(key, int) or isinstance(key, bool)
            ]
            if non_integer_keys:
                raise ValueError(
                    f"RoleMap 键必须是整数段落 idx：{non_integer_keys[:10]}")
            valid_indices = {p["idx"] for p in paragraphs}
            expected = {
                p["idx"] for p in paragraphs
                if p.get("editable", True) and not p.get("in_table")
            }
            bad_roles = [
                (key, value) for key, value in rolemap.items()
                if not isinstance(value, str) or value not in BASE_ROLES
            ]
            if bad_roles:
                raise ValueError(
                    f"RoleMap 含非法角色 {bad_roles[:10]}，合法枚举：{BASE_ROLES}")
            missing = expected - set(rolemap)
            if missing:
                raise ValueError(f"RoleMap 未覆盖这些非表格段落：{sorted(missing)}")
            extra = set(rolemap) - valid_indices
            if extra:
                raise ValueError(f"RoleMap 含不存在的段落 idx：{sorted(extra)}")
            self._emit("标注角色", "RoleMap 由外部给定（JSON），校验通过，跳过标注", status="ok")
        else:
            self._emit("标注角色", "正在逐段判断角色（标题/正文/落款/日期 ...）")
            from core.label_roles import label_roles
            rolemap = label_roles(
                paragraphs, self._get_llm(),
                on_event=lambda m: self._emit("标注角色", m),
                profile=spec.get("profile"))
            counts = {}
            for r in rolemap.values():
                counts[r] = counts.get(r, 0) + 1
            summary = "、".join(f"{k}×{v}" for k, v in sorted(counts.items()))
            self._emit("标注角色", f"角色标注完成：{summary}", status="ok", data=rolemap)

        # ④ 确定性执行排版
        transaction = AtomicOutputSet(final_paths)
        candidate_out = transaction.temp("out_path")
        candidate_report = transaction.temp("report_path")
        candidate_report_docx = transaction.temp("report_docx_path")
        candidate_tracked = transaction.temp("tracked_path")
        self._emit("执行排版", "正在按 FormatSpec × RoleMap 逐段改写文档（确定性代码，AI 不碰 docx）...")
        changelog = apply_format(
            target_path, spec, rolemap, candidate_out,
            template_path=template_path,
            allow_risky_structure=allow_risky_structure)
        write_report(changelog, spec, candidate_report)
        # 同步产出：docx 检测报告 + 修订模式文档（Word 审阅视图可见改动）
        from core.report_docx import build_report_docx
        build_report_docx(changelog, spec, candidate_report_docx)
        apply_format(
            target_path, spec, rolemap, candidate_tracked, track=True,
            template_path=template_path,
            allow_risky_structure=allow_risky_structure)
        n_changed = sum(1 for c in changelog if c["changed_fields"])
        n_styles = len({c.get("style_name") for c in changelog if c.get("style_name")})
        self._emit("执行排版",
                   f"排版完成：已创建/更新 {n_styles} 个 Word 命名样式，"
                   f"并应用到 {n_changed} 个段落，候选稿等待完整性验收",
                   status="ok", data=changelog)

        # ④.5 文本一致性校验：排版只许改格式，正文一个字都不能动
        from core.text_integrity import check_text_integrity
        allowed_additions = []
        if ((spec.get("toc") or {}).get("enabled")
                and not (spec.get("structure") or {}).get("enabled")):
            allowed_additions = ["目录", "（在 Word 中右键此处选择「更新域」生成目录）"]
        # 手工编号被自动编号替换而剥掉的前缀属于预期内的文字变化
        expected_prefixes = []
        changed_idxs = {c["idx"] for c in changelog
                        if "manual_number_prefix_removed" in c.get("changed_fields", [])}
        for p in paragraphs:
            if p["idx"] in changed_idxs and p.get("manual_number"):
                expected_prefixes.append(str(p["manual_number"]))
        for change in changelog:
            allowed_additions.extend(change.get("allowed_additions") or [])
            expected_prefixes.extend(change.get("stripped_prefixes") or [])
        page_rules = spec.get("page") or {}
        allowed_story_changes = set()
        if (spec.get("structure") or {}).get("enabled"):
            allowed_story_changes.update({"headers", "footers"})
        for key in ("header", "even_header", "first_header"):
            rule = page_rules.get(key) or {}
            if rule.get("text") is not None and not rule.get("preserve_text"):
                allowed_story_changes.add("headers")
        for key in ("footer", "even_footer", "first_footer"):
            rule = page_rules.get(key) or {}
            if rule.get("text") is not None and not rule.get("preserve_text"):
                allowed_story_changes.add("footers")
        for override in page_rules.get("section_overrides") or []:
            for key in ("header", "even_header", "first_header"):
                rule = override.get(key) or {}
                if rule.get("text") is not None and not rule.get("preserve_text"):
                    allowed_story_changes.add("headers")
            for key in ("footer", "even_footer", "first_footer"):
                rule = override.get(key) or {}
                if rule.get("text") is not None and not rule.get("preserve_text"):
                    allowed_story_changes.add("footers")
        integrity = check_text_integrity(
            target_path, candidate_out,
            allowed_additions=allowed_additions,
            expected_stripped_prefixes=expected_prefixes,
            allowed_story_changes=allowed_story_changes)
        if integrity["ok"]:
            self._emit(
                "执行排版", "文本一致性校验通过：正文与受保护 Story 文字零改动",
                status="ok")
        else:
            self._emit("执行排版",
                       f"文本一致性校验发现差异：新增 {len(integrity['added'])} 段、"
                       f"缺失 {len(integrity['removed'])} 段、"
                       f"Story 差异 {len(integrity.get('story_differences', []))} 类，"
                       "请人工核对",
                       status="err", data=integrity)
            raise IntegrityViolationError(integrity)

        field_refresh = None
        if refresh_fields:
            try:
                from core.field_refresh import refresh_fields_word
                field_refresh = refresh_fields_word(candidate_out)
                self._emit(
                    "刷新域",
                    "已用 Microsoft Word 刷新目录、动态页眉和页码并保存",
                    status="ok", data=field_refresh)
            except RuntimeError as exc:
                self._emit(
                    "刷新域",
                    f"字段未能预刷新：{exc}；文档将在 Word 打开时自动更新",
                    status="warn")

        # ⑤ 视觉自检（可选，一轮定向修复，不做开放循环）
        # 注意：自检是加分项，失败（如模型不支持图片）不能拖垮已完成的排版结果。
        issues, applied = [], []
        if verify:
            try:
                self._emit("视觉自检", "正在把排版结果渲染成图，交给视觉模型对照规范质检 ...")
                from core.verify_visual import apply_fixes, verify_visual
                png_dir = os.path.splitext(candidate_out)[0] + "_verify_render"
                issues = verify_visual(
                    candidate_out, spec, self._get_llm(), png_dir,
                    on_event=lambda message: self._emit(
                        "视觉自检", message, status="warn"))
                failed = [i for i in issues if not i["pass"]]
                if not failed:
                    self._emit("视觉自检", f"自检通过：{len(issues)} 项检查全部符合规范", status="ok")
                else:
                    self._emit("视觉自检",
                               f"发现 {len(failed)} 项不符："
                               + "、".join(f"{i['role']}.{i['field']}" for i in failed),
                               status="warn", data=issues)
                    spec, applied = apply_fixes(spec, failed)
                    if applied:
                        self._emit("视觉自检",
                                   f"已定向修复 {len(applied)} 项，正在重排 ...", status="warn")
                        changelog = apply_format(
                            target_path, spec, rolemap, candidate_out,
                            template_path=template_path,
                            allow_risky_structure=allow_risky_structure)
                        write_report(changelog, spec, candidate_report)
                        from core.report_docx import build_report_docx
                        build_report_docx(changelog, spec, candidate_report_docx)
                        apply_format(
                            target_path, spec, rolemap, candidate_tracked, track=True,
                            template_path=template_path,
                            allow_risky_structure=allow_risky_structure)
                        integrity = check_text_integrity(
                            target_path, candidate_out,
                            allowed_additions=allowed_additions,
                            expected_stripped_prefixes=expected_prefixes,
                            allowed_story_changes=allowed_story_changes)
                        self._emit("视觉自检", "修复后重排完成"
                                   + ("，文本一致性校验通过" if integrity["ok"]
                                      else "，但文本一致性校验发现差异，请人工核对"),
                                   status="ok" if integrity["ok"] else "err")
                        if not integrity["ok"]:
                            raise IntegrityViolationError(integrity)
                        if refresh_fields:
                            try:
                                from core.field_refresh import refresh_fields_word
                                field_refresh = refresh_fields_word(candidate_out)
                                self._emit(
                                    "刷新域", "修复后已重新刷新 Word 域",
                                    status="ok", data=field_refresh)
                            except RuntimeError as exc:
                                self._emit(
                                    "刷新域", f"修复后字段未能预刷新：{exc}",
                                    status="warn")
                    else:
                        self._emit("视觉自检",
                                   "这些问题无法安全自动修复，已保留在问题清单中供人工处理",
                                   status="warn")
            except IntegrityViolationError:
                raise  # 内容变化属于硬失败，不能被视觉自检的降级逻辑吞掉。
            except Exception as e:  # noqa: BLE001 —— 自检失败降级为警告
                from core.verify_visual import (
                    VisualInconclusiveError,
                    VisualModelError,
                    VisualRenderError,
                    VisualResponseError,
                )
                if isinstance(e, VisualRenderError):
                    detail = f"渲染阶段失败：{e}"
                elif isinstance(e, VisualModelError):
                    detail = f"多模态请求失败：{e}"
                elif isinstance(e, VisualResponseError):
                    detail = f"模型 JSON/结构校验失败：{e}"
                elif isinstance(e, VisualInconclusiveError):
                    detail = f"模型无法下结论：{e}"
                else:
                    detail = f"未预期错误：{e}"
                self._emit(
                    "视觉自检",
                    f"自检未完成（{detail}）。候选排版稿仍将按结构校验结果提交，"
                    "本次不会被误报为“0 项全部通过”",
                    status="err")

        transaction.commit(main_key="out_path")
        self._emit("完成", "全部流程结束", status="ok")
        stylemap = {
            role: style_name_for_role(role, rule)
            for role, rule in (spec.get("roles") or {}).items()
        }
        return {
            "spec": spec, "paragraphs": paragraphs, "rolemap": rolemap,
            "stylemap": stylemap, "changelog": changelog,
            "issues": issues, "applied_fixes": applied,
            "text_integrity": integrity,
            "field_refresh": field_refresh,
            "preflight": preflight,
            "input_conversion": _conversion_context,
            "out_path": out_path, "report_path": report_path,
            "report_docx_path": report_docx_path, "tracked_path": tracked_path,
        }
