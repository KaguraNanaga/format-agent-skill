"""受支持输入格式到临时 DOCX 的显式、可审计转换层。"""

import contextlib
import os
import shutil
import subprocess
import tempfile
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path


SUPPORTED_INPUT_EXTENSIONS = {".docx", ".doc", ".wps", ".odt", ".rtf"}
UNSUPPORTED_INPUT_EXTENSIONS = {".pdf"}


class InputConversionError(RuntimeError):
    pass


@dataclass(frozen=True)
class ConvertedInput:
    source_path: str
    docx_path: str
    converter: str
    lossy: bool
    warnings: tuple

    def as_dict(self):
        value = asdict(self)
        value["warnings"] = list(self.warnings)
        return value


def _validate_docx(path):
    if not Path(path).is_file() or not zipfile.is_zipfile(path):
        raise InputConversionError(f"转换器没有生成有效 DOCX：{path}")


def _office_candidates(extension):
    wps = (("Kwps.Application", "WPS"), ("wps.Application", "WPS"))
    word = (("Word.Application", "Microsoft Word"),)
    return wps + word if extension == ".wps" else word + wps


def _convert_with_com(source, destination):
    if os.name != "nt":
        raise InputConversionError("COM 转换只支持 Windows")
    try:
        import pythoncom
        import win32com.client
    except ModuleNotFoundError as exc:
        raise InputConversionError(
            "缺少 pywin32，无法调用 Word/WPS 转换；请安装 pywin32 或 LibreOffice。") from exc

    errors = []
    pythoncom.CoInitialize()
    try:
        for prog_id, display_name in _office_candidates(source.suffix.lower()):
            app = document = None
            try:
                try:
                    app = win32com.client.DispatchEx(prog_id)
                except Exception:
                    app = win32com.client.Dispatch(prog_id)
                try:
                    app.Visible = False
                except Exception:
                    pass
                try:
                    app.DisplayAlerts = 0
                except Exception:
                    pass
                try:
                    # msoAutomationSecurityForceDisable：转换旧格式时禁止自动宏。
                    app.AutomationSecurity = 3
                except Exception:
                    pass
                try:
                    document = app.Documents.Open(
                        str(source), ConfirmConversions=False, ReadOnly=True,
                        AddToRecentFiles=False, Visible=False,
                        OpenAndRepair=False, NoEncodingDialog=True,
                    )
                except Exception:
                    # WPS 的 COM 参数表在不同版本间不完全一致；位置参数 2
                    # 对应 ConfirmConversions，参数 3 对应 ReadOnly。
                    document = app.Documents.Open(str(source), False, True)
                try:
                    document.SaveAs2(str(destination), FileFormat=16)
                except Exception:
                    document.SaveAs2(str(destination), 16)
                _validate_docx(destination)
                return display_name
            except Exception as exc:  # 逐一回退到下一套 Office
                errors.append(f"{display_name}: {exc}")
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
    finally:
        pythoncom.CoUninitialize()
    raise InputConversionError(
        "未能使用 Microsoft Word/WPS 转换输入文件。" + "；".join(errors))


def _libreoffice_binary():
    candidates = [shutil.which("soffice"), shutil.which("libreoffice")]
    if os.name == "nt":
        candidates.extend((
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ))
    return next((value for value in candidates if value and Path(value).is_file()), None)


def _convert_with_libreoffice(source, destination_dir):
    binary = _libreoffice_binary()
    if not binary:
        raise InputConversionError("未检测到 LibreOffice/soffice")
    command = [
        binary, "--headless", "--convert-to", "docx", "--outdir",
        str(destination_dir), str(source),
    ]
    try:
        result = subprocess.run(
            command, capture_output=True, text=True, timeout=120, check=False)
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise InputConversionError(f"LibreOffice 转换失败：{exc}") from exc
    destination = destination_dir / f"{source.stem}.docx"
    if result.returncode or not destination.exists():
        detail = (result.stderr or result.stdout or "无诊断输出").strip()
        raise InputConversionError(f"LibreOffice 转换失败：{detail}")
    _validate_docx(destination)
    return destination


@contextlib.contextmanager
def converted_input(path):
    """将输入临时转换为 DOCX；上下文退出后清理中间文件。"""
    source = Path(path).expanduser().resolve()
    if not source.is_file():
        raise FileNotFoundError(f"输入文件不存在：{source}")
    extension = source.suffix.lower()
    if extension == ".docx":
        _validate_docx(source)
        yield ConvertedInput(
            str(source), str(source), "native-docx", False, tuple())
        return
    if extension in UNSUPPORTED_INPUT_EXTENSIONS:
        raise InputConversionError(
            "PDF 不是可编辑 Word 源文件，本转换层拒绝猜测版面还原；"
            "请先提供可靠 OCR/转换后的 DOCX，再运行排版。")
    if extension not in SUPPORTED_INPUT_EXTENSIONS:
        raise InputConversionError(
            f"不支持输入扩展名 {extension!r}；支持 .docx/.doc/.wps/.odt/.rtf。")

    with tempfile.TemporaryDirectory(prefix="format-agent-input-") as temp_dir:
        directory = Path(temp_dir)
        destination = directory / f"{source.stem}.docx"
        warnings = (
            "格式转换可能改变分页、字体替代、浮动对象或域；转换后的 DOCX 会重新执行能力预检与文本一致性校验。",
        )
        converter = None
        errors = []
        if extension in {".odt", ".rtf"}:
            try:
                generated = _convert_with_libreoffice(source, directory)
                if generated != destination:
                    shutil.copy2(generated, destination)
                converter = "LibreOffice"
            except InputConversionError as exc:
                errors.append(str(exc))
        if converter is None:
            try:
                converter = _convert_with_com(source, destination)
            except InputConversionError as exc:
                errors.append(str(exc))
        if converter is None:
            raise InputConversionError("；".join(errors))
        _validate_docx(destination)
        yield ConvertedInput(
            str(source), str(destination), converter, True, warnings)
