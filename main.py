# CLI 串全流程（PLAN.md 第 7 节）：
#   python main.py --spec assets/spec.txt --target assets/messy.docx --out out/formatted.docx
# 编排逻辑统一在 core/agent.py（与演示界面 app.py 共用，保证两条入口行为一致），
# CLI 只是把 Agent 的工作日志事件打印到终端。
# 降级：--spec-json 直接喂人肉 FormatSpec JSON；--rolemap-json 直接喂人肉 RoleMap。

import argparse
import json
import os
import sys

from core.agent import Agent

_STATUS_ICON = {"run": "…", "ok": "✓", "warn": "!", "err": "✗"}


def main():
    ap = argparse.ArgumentParser(description="通用格式排版 Agent：规范/模板 + 目标文档 → 排版后 docx + 对照报告")
    ap.add_argument("--spec", help="格式规范文字（txt）路径")
    ap.add_argument("--template", help="格式模板 docx 路径（格式源第二种，确定性读规则）")
    ap.add_argument("--template-rolemap-json", help="模板的角色标注 JSON（不给则用 LLM 标注模板）")
    ap.add_argument("--spec-json", help="直接给 FormatSpec JSON（跳过规则抽取）")
    ap.add_argument("--rolemap-json", help="直接给 RoleMap JSON（跳过 LLM 角色标注）")
    ap.add_argument("--target", required=True, help="待排版的目标 docx")
    ap.add_argument("--out", required=True, help="输出 docx 路径")
    ap.add_argument("--report", help="对照报告路径（默认 <out去掉扩展名>_report.md）")
    ap.add_argument("--verify", action="store_true",
                    help="排版后用同一个多模态模型做一轮视觉验证并定向修复")
    ap.add_argument("--extract-only", action="store_true",
                    help="（Agent 内置智能模式）只抽取段落清单 JSON，不做排版；"
                         "宿主 Agent 读清单后自行产出 RoleMap/FormatSpec 再回调本程序")
    args = ap.parse_args()

    if args.extract_only:
        from core.extract import extract_paragraphs
        paragraphs = extract_paragraphs(args.target)
        out_json = os.path.splitext(args.out)[0] + "_paragraphs.json"
        os.makedirs(os.path.dirname(os.path.abspath(out_json)), exist_ok=True)
        with open(out_json, "w", encoding="utf-8") as f:
            json.dump(paragraphs, f, ensure_ascii=False, indent=2)
        print(f"段落清单已写出: {out_json}（{len(paragraphs)} 段）")
        return

    if not args.spec and not args.spec_json and not args.template:
        ap.error("必须提供 --spec（规范文字）、--template（模板 docx）或 --spec-json（FormatSpec JSON）之一")

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)

    kwargs = {"target_path": args.target, "out_path": args.out,
              "verify": args.verify}
    if args.report:
        kwargs["report_path"] = args.report
    if args.spec_json:
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

    result = Agent(on_event=print_event).run(**kwargs)

    # 归档中间产物（演示时要展示两个 JSON）
    base = os.path.splitext(args.out)[0]
    with open(base + "_formatspec.json", "w", encoding="utf-8") as f:
        json.dump(result["spec"], f, ensure_ascii=False, indent=2)
    with open(base + "_rolemap.json", "w", encoding="utf-8") as f:
        json.dump({str(k): v for k, v in sorted(result["rolemap"].items())},
                  f, ensure_ascii=False, indent=2)
    with open(base + "_stylemap.json", "w", encoding="utf-8") as f:
        json.dump(result["stylemap"], f, ensure_ascii=False, indent=2)
    if result["issues"]:
        with open(base + "_issues.json", "w", encoding="utf-8") as f:
            json.dump(result["issues"], f, ensure_ascii=False, indent=2)

    print(f"\n输出: {result['out_path']}")
    print(f"对照报告: {result['report_path']}")
    print(f"中间产物: {base}_formatspec.json / {base}_rolemap.json / {base}_stylemap.json")


if __name__ == "__main__":
    sys.exit(main())
