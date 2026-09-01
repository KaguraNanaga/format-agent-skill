# Agent 编排器 —— 把整条流水线包装成带"工作日志"的事件流，供演示界面实时展示。
# 事件: {"step": 步骤名, "message": 人话描述, "status": run|ok|warn|err, "data": 任意}
# 演示故事: 理解归 AI，动手归代码，中间用 JSON 交接 —— 日志把这个过程直播出来。

import json
import os

from core.apply import apply_format, write_report
from core.extract import extract_paragraphs
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
            verify=False, report_path=None, cleanup_mode=None,
            refresh_fields=False):
        """跑完整流程，返回结果 dict（spec/rolemap/changelog/issues/paths）。"""
        report_path = report_path or os.path.splitext(out_path)[0] + "_report.md"

        # ① 理解格式规范 → FormatSpec
        self._emit("理解规范", "开始理解格式来源，抽取格式规则 ...")
        if spec is not None:
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
            raise ValueError("必须提供 spec_text / spec / template_path 之一")

        if cleanup_mode is not None and (spec.get("cleanup") or {}).get("mode") != cleanup_mode:
            spec = dict(spec)
            spec["cleanup"] = {"mode": cleanup_mode}
            validate_spec(spec)

        # ② 解析目标文档结构
        self._emit("解析文档", "正在解析目标文档结构 ...")
        paragraphs = extract_paragraphs(target_path)
        n_table = sum(1 for p in paragraphs if p["in_table"])
        self._emit("解析文档",
                   f"发现 {len(paragraphs)} 个段落（{n_table} 段在表格内，不参与重排）",
                   status="ok", data=paragraphs)

        # ③ 标注段落角色 → RoleMap
        if rolemap is not None:
            # 外部给定（宿主 Agent 自标）的 RoleMap 也要过校验：角色合法、非表格段全覆盖
            from core.schema import BASE_ROLES
            expected = {p["idx"] for p in paragraphs if not p.get("in_table")}
            bad_roles = {r for r in rolemap.values() if r not in BASE_ROLES}
            if bad_roles:
                raise ValueError(f"RoleMap 含非法角色 {sorted(bad_roles)}，合法枚举：{BASE_ROLES}")
            missing = expected - set(rolemap)
            if missing:
                raise ValueError(f"RoleMap 未覆盖这些非表格段落：{sorted(missing)}")
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
        self._emit("执行排版", "正在按 FormatSpec × RoleMap 逐段改写文档（确定性代码，AI 不碰 docx）...")
        changelog = apply_format(
            target_path, spec, rolemap, out_path,
            template_path=template_path)
        write_report(changelog, spec, report_path)
        # 同步产出：docx 检测报告 + 修订模式文档（Word 审阅视图可见改动）
        from core.report_docx import build_report_docx
        base = os.path.splitext(out_path)[0]
        report_docx_path = base + "_report.docx"
        build_report_docx(changelog, spec, report_docx_path)
        tracked_path = base + "_tracked.docx"
        apply_format(
            target_path, spec, rolemap, tracked_path, track=True,
            template_path=template_path)
        n_changed = sum(1 for c in changelog if c["changed_fields"])
        n_styles = len({c.get("style_name") for c in changelog if c.get("style_name")})
        self._emit("执行排版",
                   f"排版完成：已创建/更新 {n_styles} 个 Word 命名样式，"
                   f"并应用到 {n_changed} 个段落，输出 {os.path.basename(out_path)}",
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
        structure_changes = [
            c for c in changelog if c.get("role") == "structure"]
        for change in structure_changes:
            allowed_additions.extend(change.get("allowed_additions") or [])
            expected_prefixes.extend(change.get("stripped_prefixes") or [])
        integrity = check_text_integrity(
            target_path, out_path,
            allowed_additions=allowed_additions,
            expected_stripped_prefixes=expected_prefixes)
        if integrity["ok"]:
            self._emit("执行排版", "文本一致性校验通过：正文内容零改动", status="ok")
        else:
            self._emit("执行排版",
                       f"文本一致性校验发现差异：新增 {len(integrity['added'])} 段、"
                       f"缺失 {len(integrity['removed'])} 段，请人工核对",
                       status="err", data=integrity)

        field_refresh = None
        if refresh_fields:
            try:
                from core.field_refresh import refresh_fields_word
                field_refresh = refresh_fields_word(out_path)
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
                png_dir = os.path.splitext(out_path)[0] + "_verify_render"
                issues = verify_visual(
                    out_path, spec, self._get_llm(), png_dir,
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
                            target_path, spec, rolemap, out_path,
                            template_path=template_path)
                        write_report(changelog, spec, report_path)
                        from core.report_docx import build_report_docx
                        build_report_docx(changelog, spec, report_docx_path)
                        apply_format(
                            target_path, spec, rolemap, tracked_path, track=True,
                            template_path=template_path)
                        integrity = check_text_integrity(
                            target_path, out_path,
                            allowed_additions=allowed_additions,
                            expected_stripped_prefixes=expected_prefixes)
                        self._emit("视觉自检", "修复后重排完成"
                                   + ("，文本一致性校验通过" if integrity["ok"]
                                      else "，但文本一致性校验发现差异，请人工核对"),
                                   status="ok" if integrity["ok"] else "err")
                        if refresh_fields:
                            try:
                                from core.field_refresh import refresh_fields_word
                                field_refresh = refresh_fields_word(out_path)
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
                    f"自检未完成（{detail}）。排版 DOCX 已保留，"
                    "本次不会被误报为“0 项全部通过”",
                    status="err")

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
            "out_path": out_path, "report_path": report_path,
            "report_docx_path": report_docx_path, "tracked_path": tracked_path,
        }
