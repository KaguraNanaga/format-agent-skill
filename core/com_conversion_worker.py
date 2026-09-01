"""单个 Office COM 候选转换工作进程。

不要从业务代码直接导入这个模块。父进程为每个 Word/WPS 候选启动一次，
因此某个 COM 服务器无响应时可以限时终止并继续下一个候选。
"""

from __future__ import annotations

import argparse
import json
import sys
import zipfile
from pathlib import Path


def _valid_docx(path: Path) -> bool:
    return path.is_file() and zipfile.is_zipfile(path)


def convert(source: Path, destination: Path, prog_id: str, display_name: str) -> dict:
    try:
        import pythoncom
        import win32com.client
    except ModuleNotFoundError as exc:
        return {
            "ok": False,
            "converter": display_name,
            "error": (
                "缺少 pywin32，无法调用 Word/WPS 转换；"
                "请在运行 Format Agent 的 Python 环境安装 pywin32。"
            ),
            "exception_type": type(exc).__name__,
        }

    app = document = None
    pythoncom.CoInitialize()
    try:
        # 只用 DispatchEx，避免 Dispatch 静默附着到用户已打开的 Office，
        # 随后的 Quit() 会关闭用户会话并可能造成未保存内容丢失。
        app = win32com.client.DispatchEx(prog_id)
        try:
            app.Visible = False
        except Exception:
            pass
        try:
            app.DisplayAlerts = 0
        except Exception:
            pass
        try:
            app.AutomationSecurity = 3
        except Exception:
            pass
        try:
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
            # WPS 的 COM 参数表在不同版本间不完全一致。
            document = app.Documents.Open(str(source), False, True)
        # Word 接受 wdFormatDocumentDefault=16；WPS 12.x 对 16 可能无响应，
        # 但兼容 wdFormatXMLDocument=12。二者都生成标准 OOXML DOCX。
        file_format = (
            12 if prog_id.lower().startswith(("kwps", "wps")) else 16
        )
        try:
            document.SaveAs2(str(destination), FileFormat=file_format)
        except Exception:
            document.SaveAs2(str(destination), file_format)
        if not _valid_docx(destination):
            raise RuntimeError(f"转换器没有生成有效 DOCX：{destination}")
        return {"ok": True, "converter": display_name}
    except Exception as exc:  # noqa: BLE001 - 必须把 COM 诊断传回父进程
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
