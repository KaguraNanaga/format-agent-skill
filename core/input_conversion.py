"""受支持输入格式到临时 DOCX 的显式、可审计转换层。"""

import contextlib
import ctypes
import json
import os
import signal
import shutil
import subprocess
import sys
import tempfile
import time
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


def _com_timeout_seconds():
    raw = os.environ.get("FORMAT_AGENT_COM_TIMEOUT_SECONDS", "45")
    try:
        value = float(raw)
    except ValueError:
        value = 45.0
    return min(300.0, max(5.0, value))


def _office_process_names(prog_id):
    return ("wps.exe",) if prog_id.lower().startswith(("kwps", "wps")) else ("winword.exe",)


def _office_pids(process_names):
    """用 Win32 Toolhelp 枚举 Office 进程；不依赖 tasklist 的权限/语言。"""
    if os.name != "nt":
        return set()
    from ctypes import wintypes

    class PROCESSENTRY32W(ctypes.Structure):
        _fields_ = (
            ("dwSize", wintypes.DWORD),
            ("cntUsage", wintypes.DWORD),
            ("th32ProcessID", wintypes.DWORD),
            ("th32DefaultHeapID", ctypes.c_size_t),
            ("th32ModuleID", wintypes.DWORD),
            ("cntThreads", wintypes.DWORD),
            ("th32ParentProcessID", wintypes.DWORD),
            ("pcPriClassBase", wintypes.LONG),
            ("dwFlags", wintypes.DWORD),
            ("szExeFile", wintypes.WCHAR * 260),
        )

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateToolhelp32Snapshot.argtypes = (wintypes.DWORD, wintypes.DWORD)
    kernel32.CreateToolhelp32Snapshot.restype = wintypes.HANDLE
    kernel32.Process32FirstW.argtypes = (
        wintypes.HANDLE, ctypes.POINTER(PROCESSENTRY32W))
    kernel32.Process32FirstW.restype = wintypes.BOOL
    kernel32.Process32NextW.argtypes = (
        wintypes.HANDLE, ctypes.POINTER(PROCESSENTRY32W))
    kernel32.Process32NextW.restype = wintypes.BOOL
    kernel32.CloseHandle.argtypes = (wintypes.HANDLE,)
    kernel32.CloseHandle.restype = wintypes.BOOL
    snapshot = kernel32.CreateToolhelp32Snapshot(0x00000002, 0)
    invalid_handle = ctypes.c_void_p(-1).value
    if snapshot == invalid_handle:
        return set()
    wanted = {name.casefold() for name in process_names}
    result = set()
    entry = PROCESSENTRY32W()
    entry.dwSize = ctypes.sizeof(entry)
    try:
        success = kernel32.Process32FirstW(snapshot, ctypes.byref(entry))
        while success:
            if entry.szExeFile.casefold() in wanted:
                result.add(int(entry.th32ProcessID))
            success = kernel32.Process32NextW(snapshot, ctypes.byref(entry))
    finally:
        kernel32.CloseHandle(snapshot)
    return result


def _terminate_new_office_processes(process_names, before):
    """超时后只结束本次新出现的 Office 进程，不触碰用户原有会话。"""
    if os.name != "nt":
        return []
    terminated = []
    # COM 激活进程可能在工作进程被终止后才迟到启动；短暂轮询，防止
    # WINWORD/WPS 成为永久后台孤儿。before 中的用户原有进程永不触碰。
    for _ in range(10):
        time.sleep(0.5)
        new_pids = _office_pids(process_names) - set(before) - set(terminated)
        for pid in sorted(new_pids):
            try:
                kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
                kernel32.OpenProcess.argtypes = (
                    ctypes.c_uint32, ctypes.c_int, ctypes.c_uint32)
                kernel32.OpenProcess.restype = ctypes.c_void_p
                kernel32.TerminateProcess.argtypes = (ctypes.c_void_p, ctypes.c_uint32)
                kernel32.TerminateProcess.restype = ctypes.c_int
                kernel32.CloseHandle.argtypes = (ctypes.c_void_p,)
                handle = kernel32.OpenProcess(0x0001, False, pid)
                if handle:
                    try:
                        if kernel32.TerminateProcess(handle, 1):
                            terminated.append(pid)
                            continue
                    finally:
                        kernel32.CloseHandle(handle)
                os.kill(pid, signal.SIGTERM)
                terminated.append(pid)
            except OSError:
                pass
    return terminated


def _convert_with_com_candidate(source, destination, prog_id, display_name):
    process_names = _office_process_names(prog_id)
    before = _office_pids(process_names)
    worker = Path(__file__).with_name("com_conversion_worker.py")
    command = [
        sys.executable, str(worker), str(source), str(destination),
        prog_id, display_name,
    ]
    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    timeout = _com_timeout_seconds()
    try:
        result = subprocess.run(
            command, capture_output=True, text=True, timeout=timeout,
            check=False, creationflags=creationflags,
        )
    except subprocess.TimeoutExpired as exc:
        terminated = _terminate_new_office_processes(process_names, before)
        cleanup = f"；已终止新建 Office 进程 {terminated}" if terminated else ""
        raise InputConversionError(
            f"{display_name} COM 转换在 {timeout:g} 秒后超时{cleanup}") from exc
    output = (result.stdout or "").strip().splitlines()
    try:
        payload = json.loads(output[-1]) if output else {}
    except json.JSONDecodeError:
        payload = {}
    if result.returncode or not payload.get("ok"):
        detail = payload.get("error") or (result.stderr or result.stdout or "无诊断输出").strip()
        raise InputConversionError(f"{display_name}: {detail}")
    _validate_docx(destination)
    return payload.get("converter") or display_name


def _convert_with_com(source, destination):
    if os.name != "nt":
        raise InputConversionError("COM 转换只支持 Windows")
    errors = []
    for prog_id, display_name in _office_candidates(source.suffix.lower()):
        try:
            return _convert_with_com_candidate(
                source, destination, prog_id, display_name)
        except InputConversionError as exc:
            errors.append(str(exc))
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
