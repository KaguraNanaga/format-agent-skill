"""单个 Microsoft Word/WPS COM 候选的 DOCX → PDF 工作进程。"""

from __future__ import annotations

import argparse
import json
from pathlib import Path


def convert(source: Path, destination: Path, prog_id: str, display_name: str) -> dict:
    try:
        import pythoncom
        import win32com.client
    except ModuleNotFoundError as exc:
        return {
            "ok": False,
            "converter": display_name,
            "error": "缺少 pywin32，无法调用 Word/WPS 导出 PDF。",
            "exception_type": type(exc).__name__,
        }

    app = document = None
    pythoncom.CoInitialize()
    try:
        app = win32com.client.DispatchEx(prog_id)
        try:
            app.Visible = False
            app.DisplayAlerts = 0
            app.AutomationSecurity = 3
            app.Options.UpdateLinksAtOpen = False
        except Exception:
            pass
        try:
            document = app.Documents.Open(
                str(source), ConfirmConversions=False, ReadOnly=True,
                AddToRecentFiles=False, Visible=False,
                OpenAndRepair=False, NoEncodingDialog=True,
            )
        except Exception:
            document = app.Documents.Open(str(source), False, True)
        try:
            document.SaveAs2(str(destination), FileFormat=17)
        except Exception:
            document.SaveAs2(str(destination), 17)
        if not destination.is_file() or destination.read_bytes()[:5] != b"%PDF-":
            raise RuntimeError(f"转换器没有生成有效 PDF：{destination}")
        return {"ok": True, "converter": display_name}
    except Exception as exc:  # noqa: BLE001 - 把 Office 诊断传回父进程
        return {
            "ok": False,
            "converter": display_name,
            "error": str(exc),
            "exception_type": type(exc).__name__,
        }
    finally:
        if document is not None:
            try:
                document.Close(SaveChanges=False)
            except Exception:
                pass
        if app is not None:
            try:
                app.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("source", type=Path)
    parser.add_argument("destination", type=Path)
    parser.add_argument("prog_id")
    parser.add_argument("display_name")
    args = parser.parse_args()
    payload = convert(
        args.source.resolve(), args.destination.resolve(),
        args.prog_id, args.display_name,
    )
    print(json.dumps(payload, ensure_ascii=False))
    return 0 if payload["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
