"""用 Windows Word 刷新 TOC/STYLEREF/PAGE 域并保存缓存结果。"""

import os
import sys


def refresh_fields_word(docx_path):
    """在独立 Word 进程中刷新正文、目录、页眉和页脚域。

    非 Windows 或未安装 Word/pywin32 时抛出 RuntimeError；调用方应降级为
    文档内的 w:updateFields（下次在 Word 打开时自动更新）。
    """
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
    counts = {"body_fields": 0, "toc": 0, "header_footer_fields": 0}
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(
            absolute_path,
            ConfirmConversions=False,
            ReadOnly=False,
            AddToRecentFiles=False,
        )

        counts["body_fields"] = int(document.Fields.Count)
        if document.Fields.Count:
            document.Fields.Update()

        counts["toc"] = int(document.TablesOfContents.Count)
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
                            counts["header_footer_fields"] += int(
                                story.Range.Fields.Count)
                            if story.Range.Fields.Count:
                                story.Range.Fields.Update()
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
