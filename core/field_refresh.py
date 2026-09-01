"""用 Windows Word 刷新 TOC/STYLEREF/PAGE 域并保存缓存结果。"""

import json
import os
import re
import subprocess
import sys
from pathlib import Path


SAFE_FIELD_KINDS = {
    "PAGE", "NUMPAGES", "SECTION", "SECTIONPAGES", "DATE", "TIME",
    "AUTHOR", "TITLE", "SUBJECT", "FILENAME", "FILESIZE", "DOCPROPERTY",
    "DOCVARIABLE", "REF", "PAGEREF", "NOTEREF", "STYLEREF", "SEQ",
    "LISTNUM", "AUTONUM", "AUTONUMLGL", "AUTONUMOUT", "TOC", "TOA",
    "INDEX", "XE", "TA", "CITATION", "BIBLIOGRAPHY", "EQ", "SYMBOL",
}


def _field_kind(field):
    try:
        code = str(field.Code.Text or "")
    except Exception:
        return "UNKNOWN"
    match = re.match(r"\s*([A-Za-z][A-Za-z0-9]*)", code)
    return match.group(1).upper() if match else "UNKNOWN"


def _update_safe_fields(fields):
    result = {"total": int(fields.Count), "updated": 0, "skipped": {}}
    for index in range(1, int(fields.Count) + 1):
        field = fields(index)
        kind = _field_kind(field)
        if kind not in SAFE_FIELD_KINDS:
            result["skipped"][kind] = result["skipped"].get(kind, 0) + 1
            continue
        field.Update()
        result["updated"] += 1
    return result


def _refresh_fields_word_in_process(docx_path):
    """工作进程内部实现；公开入口必须通过限时父进程调用。"""
    if sys.platform != "win32":
        raise RuntimeError("字段落盘刷新仅支持安装了 Microsoft Word 的 Windows")
    try:
        import win32com.client
    except ImportError as exc:
        raise RuntimeError("缺少 pywin32，无法调用 Microsoft Word 刷新域") from exc

    # Word 正在编辑其他文档（尤其存在未保存文档或模态对话框）时，新的
    # 自动化调用可能长时间阻塞，也可能打扰用户当前会话。此时安全降级，
    # 依靠文档内已写入的 w:updateFields 在下次打开时更新。
    try:
        active_word = win32com.client.GetActiveObject("Word.Application")
        open_documents = int(active_word.Documents.Count)
    except Exception:
        open_documents = 0
    if open_documents:
        raise RuntimeError(
            f"检测到 Microsoft Word 正在编辑 {open_documents} 个文档；"
            "为避免干扰当前会话，本次不启动后台刷新")

    absolute_path = os.path.abspath(docx_path)
    word = None
    document = None
    counts = {
        "body_fields": 0, "body_fields_updated": 0, "toc": 0,
        "toc_skipped_external": 0,
        "header_footer_fields": 0, "header_footer_fields_updated": 0,
        "skipped_field_kinds": {},
    }
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            word.AutomationSecurity = 3
        except Exception:
            pass
        word.Options.UpdateLinksAtOpen = False
        document = word.Documents.Open(
            absolute_path,
            ConfirmConversions=False,
            ReadOnly=False,
            AddToRecentFiles=False,
        )

        body_result = _update_safe_fields(document.Fields)
        counts["body_fields"] = body_result["total"]
        counts["body_fields_updated"] = body_result["updated"]
        counts["skipped_field_kinds"].update(body_result["skipped"])

        counts["toc"] = int(document.TablesOfContents.Count)
        if body_result["skipped"].get("RD"):
            # RD 域可让目录引用其他文件；不应因一次排版而访问外部文档。
            counts["toc_skipped_external"] = counts["toc"]
        else:
            for index in range(1, document.TablesOfContents.Count + 1):
                toc = document.TablesOfContents(index)
                toc.Update()
                toc.UpdatePageNumbers()

        for section in document.Sections:
            for story_name in ("Headers", "Footers"):
                collection = getattr(section, story_name)
                for story_type in (1, 2, 3):
                    try:
                        story = collection(story_type)
                        if story.Exists:
                            story_result = _update_safe_fields(
                                story.Range.Fields)
                            counts["header_footer_fields"] += story_result["total"]
                            counts["header_footer_fields_updated"] += story_result["updated"]
                            for kind, number in story_result["skipped"].items():
                                counts["skipped_field_kinds"][kind] = (
                                    counts["skipped_field_kinds"].get(kind, 0)
                                    + number
                                )
                    except Exception:  # 某些节未创建首页/奇偶页 story
                        continue

        document.Repaginate()
        document.Save()
        return counts
    except Exception as exc:
        raise RuntimeError(f"Microsoft Word 刷新域失败：{exc}") from exc
    finally:
        if document is not None:
            try:
                document.Close(SaveChanges=False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass


def refresh_fields_word(docx_path):
    """在限时隔离进程中刷新正文、目录、页眉和页脚域。

    非 Windows、缺少 Word/pywin32、Word 正在被用户使用或 COM 超时时抛出
    RuntimeError；调用方应降级为 w:updateFields（下次在 Word 打开时更新）。
    """
    if sys.platform != "win32":
        raise RuntimeError("字段落盘刷新仅支持安装了 Microsoft Word 的 Windows")
    from core.input_conversion import (
        _com_timeout_seconds,
        _office_pids,
        _terminate_new_office_processes,
    )

    absolute_path = str(Path(docx_path).expanduser().resolve())
    process_names = ("winword.exe",)
    before = _office_pids(process_names)
    command = [sys.executable, str(Path(__file__).resolve()), "--worker", absolute_path]
    timeout = _com_timeout_seconds()
    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        result = subprocess.run(
            command, capture_output=True, text=True, timeout=timeout,
            check=False, creationflags=creationflags,
        )
    except subprocess.TimeoutExpired as exc:
        terminated = _terminate_new_office_processes(process_names, before)
        cleanup = f"；已终止新建 Word 进程 {terminated}" if terminated else ""
        raise RuntimeError(
            f"Microsoft Word 刷新域在 {timeout:g} 秒后超时{cleanup}") from exc
    lines = (result.stdout or "").strip().splitlines()
    try:
        payload = json.loads(lines[-1]) if lines else {}
    except json.JSONDecodeError:
        payload = {}
    if result.returncode or not payload.get("ok"):
        detail = payload.get("error") or (result.stderr or result.stdout or "无诊断输出").strip()
        raise RuntimeError(f"Microsoft Word 刷新域失败：{detail}")
    return payload["counts"]


def _worker_main(path):
    try:
        counts = _refresh_fields_word_in_process(path)
        payload = {"ok": True, "counts": counts}
        code = 0
    except Exception as exc:  # noqa: BLE001 - 把 Word 诊断传回父进程
        payload = {
            "ok": False, "error": str(exc),
            "exception_type": type(exc).__name__,
        }
        code = 1
    print(json.dumps(payload, ensure_ascii=False))
    return code


if __name__ == "__main__" and len(sys.argv) == 3 and sys.argv[1] == "--worker":
    raise SystemExit(_worker_main(sys.argv[2]))
