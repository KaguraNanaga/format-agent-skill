# CLI 串全流程（PLAN.md 第 7 节）：
#   python main.py --spec assets/spec.txt --target assets/messy.docx --out out/formatted.docx
# 编排逻辑统一在 core/agent.py（与演示界面 app.py 共用，保证两条入口行为一致），
# CLI 只是把 Agent 的工作日志事件打印到终端。
# 降级：--spec-json 直接喂人肉 FormatSpec JSON；--rolemap-json 直接喂人肉 RoleMap。

import argparse
import json
import os
import sys
from contextlib import ExitStack

from core.agent import Agent

_STATUS_ICON = {"run": "…", "ok": "✓", "warn": "!", "err": "✗"}


def main():
    from core.style_packs import list_style_packs
    ap = argparse.ArgumentParser(description="通用格式排版 Agent：规范/模板 + 目标文档 → 排版后 docx + 对照报告")
    ap.add_argument("--spec", help="格式规范文字（txt）路径")
    ap.add_argument("--template", help="格式模板 docx 路径（格式源第二种，确定性读规则）")
    ap.add_argument("--template-rolemap-json", help="模板的角色标注 JSON（不给则用 LLM 标注模板）")
    ap.add_argument("--spec-json", help="直接给 FormatSpec JSON（跳过规则抽取）")
    ap.add_argument(
        "--style-pack", choices=tuple(list_style_packs()),
        help="使用内置版式基线包；机关/学校/法院/投稿机构模板仍优先")
    ap.add_argument(
        "--running-head", help="MLA 等 Style Pack 的页眉文字（通常为作者姓氏）")
    ap.add_argument("--rolemap-json", help="直接给 RoleMap JSON（跳过 LLM 角色标注）")
    ap.add_argument(
        "--target", required=True,
        help="待排版输入：.docx/.doc/.wps/.odt/.rtf（统一输出 DOCX）")
    ap.add_argument("--out", required=True, help="输出 docx 路径")
    ap.add_argument("--report", help="对照报告路径（默认 <out去掉扩展名>_report.md）")
    ap.add_argument("--verify", action="store_true",
                    help="排版后用同一个多模态模型做一轮视觉验证并定向修复")
    ap.add_argument(
        "--refresh-fields", action="store_true",
        help="Windows 下调用 Microsoft Word 刷新目录、动态页眉和页码并保存")
    ap.add_argument(
        "--allow-risky-structure", action="store_true",
        help="显式允许论文结构、分栏或横向表格等分节调整（默认安全阻断）")
    ap.add_argument(
        "--preflight-only", action="store_true",
        help="只扫描 Story、分节、修订、文本框等能力风险并输出 JSON，不做排版")
    ap.add_argument(
        "--cleanup-mode", choices=("controlled", "strict", "preserve_emphasis"),
        help="直接格式清洗策略：受控字段/严格克隆/保留正文粗斜体强调")
    ap.add_argument("--extract-only", action="store_true",
                    help="（Agent 内置智能模式）只抽取段落清单 JSON，不做排版；"
                         "宿主 Agent 读清单后自行产出 RoleMap/FormatSpec 再回调本程序")
    args = ap.parse_args()

    if args.preflight_only:
        from core.preflight import (
            PreflightError, merge_preflight_reports, preflight_docx,
        )
        from core.safe_output import write_json_atomic
        from core.schema import SpecValidationError, validate_spec
        from core.input_conversion import InputConversionError, converted_input
        try:
            preflight_spec = None
            if args.spec_json:
                with open(args.spec_json, encoding="utf-8") as handle:
                    preflight_spec = json.load(handle)
                validate_spec(preflight_spec)
            elif args.style_pack:
                from core.style_packs import get_style_pack
                preflight_spec = get_style_pack(
                    args.style_pack, running_head=args.running_head)
            with ExitStack() as stack:
                converted_target = stack.enter_context(converted_input(args.target))
                converted_template = (
                    stack.enter_context(converted_input(args.template))
                    if args.template else None
                )
                target_report = preflight_docx(
                    converted_target.docx_path,
                    spec=preflight_spec,
                    allow_risky_structure=args.allow_risky_structure,
                )
                template_report = (
                    preflight_docx(
                        converted_template.docx_path, template_source=True)
                    if converted_template else None
                )
                report = merge_preflight_reports(target_report, template_report)
                report["input_conversion"] = {
                    "target": converted_target.as_dict(),
                    "template": (
                        converted_template.as_dict() if converted_template else None),
                }
        except (PreflightError, SpecValidationError, InputConversionError,
                OSError, json.JSONDecodeError) as exc:
            print(f"能力预检失败：{exc}", file=sys.stderr)
            return 2
        preflight_path = os.path.splitext(args.out)[0] + "_preflight.json"
        write_json_atomic(preflight_path, report)
        print(f"能力预检: {preflight_path}")
        print(f"阻断 {len(report['blockers'])} 项 / 警告 {len(report['warnings'])} 项")
        return 0 if report["ok"] else 2

    if args.extract_only:
        from core.extract import extract_paragraphs
        from core.input_conversion import InputConversionError, converted_input
        try:
            with converted_input(args.target) as converted:
                paragraphs = extract_paragraphs(converted.docx_path)
        except (InputConversionError, OSError) as exc:
            print(f"输入转换失败：{exc}", file=sys.stderr)
            return 5
        out_json = os.path.splitext(args.out)[0] + "_paragraphs.json"
        os.makedirs(os.path.dirname(os.path.abspath(out_json)), exist_ok=True)
        with open(out_json, "w", encoding="utf-8") as f:
            json.dump(paragraphs, f, ensure_ascii=False, indent=2)
        print(f"段落清单已写出: {out_json}（{len(paragraphs)} 段）")
        return

    format_sources = [
        bool(args.spec), bool(args.spec_json), bool(args.template),
        bool(args.style_pack),
    ]
    if sum(format_sources) != 1:
        ap.error(
            "必须且只能提供 --spec、--template、--spec-json、--style-pack 之一")

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)

    kwargs = {"target_path": args.target, "out_path": args.out,
              "verify": args.verify, "refresh_fields": args.refresh_fields,
              "allow_risky_structure": args.allow_risky_structure}
    if args.cleanup_mode:
        kwargs["cleanup_mode"] = args.cleanup_mode
    if args.report:
        kwargs["report_path"] = args.report
    if args.style_pack:
        kwargs["style_pack"] = args.style_pack
        kwargs["style_pack_options"] = {"running_head": args.running_head}
    elif args.spec_json:
        with open(args.spec_json, encoding="utf-8") as f:
            kwargs["spec"] = json.load(f)
    elif args.template:
        kwargs["template_path"] = args.template
        if args.template_rolemap_json:
            with open(args.template_rolemap_json, encoding="utf-8") as f:
                kwargs["template_rolemap"] = {int(k): v for k, v in json.load(f).items()}
    else:
        with open(args.spec, encoding="utf-8") as f:
            kwargs["spec_text"] = f.read()
    if args.rolemap_json:
        with open(args.rolemap_json, encoding="utf-8") as f:
            kwargs["rolemap"] = {int(k): v for k, v in json.load(f).items()}

    def print_event(e):
        icon = _STATUS_ICON.get(e["status"], " ")
        print(f"{icon} [{e['step']}] {e['message']}")

    base = os.path.splitext(args.out)[0]
    from core.safe_output import (
        IntegrityViolationError, UnsafeOutputPathError,
        validate_output_paths, write_json_atomic,
    )
    from core.preflight import PreflightBlockedError, PreflightError
    from core.input_conversion import InputConversionError
    try:
        # CLI 还会写中间 JSON；把它们也纳入碰撞检测，避免自定义 report
        # 路径在成功后又被 metadata 静默覆盖。
        validate_output_paths(args.target, {
            "out_path": args.out,
            "report_path": args.report or base + "_report.md",
            "report_docx_path": base + "_report.docx",
            "tracked_path": base + "_tracked.docx",
            "formatspec_path": base + "_formatspec.json",
            "rolemap_path": base + "_rolemap.json",
            "stylemap_path": base + "_stylemap.json",
            "preflight_path": base + "_preflight.json",
            "issues_path": base + "_issues.json",
        })
        result = Agent(on_event=print_event).run(**kwargs)
    except PreflightBlockedError as exc:
        write_json_atomic(base + "_preflight.json", exc.report)
        print(f"\n能力预检阻断：{base}_preflight.json", file=sys.stderr)
        return 2
    except IntegrityViolationError as exc:
        write_json_atomic(base + "_integrity_failure.json", exc.integrity)
        print(f"\n文本一致性失败，未提交终稿：{base}_integrity_failure.json", file=sys.stderr)
        return 3
    except UnsafeOutputPathError as exc:
        print(f"\n输出路径不安全，未执行：{exc}", file=sys.stderr)
        return 4
    except PreflightError as exc:
        print(f"\n能力预检失败，未执行：{exc}", file=sys.stderr)
        return 2
    except InputConversionError as exc:
        print(f"\n输入转换失败，未执行：{exc}", file=sys.stderr)
        return 5

    # 归档中间产物（演示时要展示两个 JSON）
    write_json_atomic(base + "_formatspec.json", result["spec"])
    write_json_atomic(
        base + "_rolemap.json",
        {str(k): v for k, v in sorted(result["rolemap"].items())})
    write_json_atomic(base + "_stylemap.json", result["stylemap"])
    write_json_atomic(base + "_preflight.json", result["preflight"])
    if result["issues"]:
        write_json_atomic(base + "_issues.json", result["issues"])

    print(f"\n输出: {result['out_path']}")
    print(f"对照报告: {result['report_path']}")
    print(f"中间产物: {base}_formatspec.json / {base}_rolemap.json / {base}_stylemap.json")
    return 0


if __name__ == "__main__":
    sys.exit(main())
