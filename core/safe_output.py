"""输出路径保护、同目录临时文件和原子提交。"""

import json
import os
import tempfile
import weakref


class UnsafeOutputPathError(ValueError):
    pass


class IntegrityViolationError(RuntimeError):
    def __init__(self, integrity):
        self.integrity = integrity
        super().__init__(
            "文本一致性校验失败："
            f"新增 {len(integrity.get('added', []))} 段，"
            f"缺失 {len(integrity.get('removed', []))} 段，"
            f"Story 差异 {len(integrity.get('story_differences', []))} 类")


def _normalized(path):
    return os.path.normcase(os.path.realpath(os.path.abspath(os.fspath(path))))


def validate_output_paths(target_path, output_paths):
    """禁止覆盖源稿、非 DOCX 主稿，以及多个产物互相覆盖。"""
    target = _normalized(target_path)
    normalized = []
    for label, path in output_paths.items():
        value = _normalized(path)
        if value == target:
            raise UnsafeOutputPathError(f"{label} 不能与源文档使用同一路径")
        normalized.append((label, value))
    seen = {}
    for label, value in normalized:
        if value in seen:
            raise UnsafeOutputPathError(
                f"{label} 与 {seen[value]} 指向同一输出路径")
        seen[value] = label
    main_path = output_paths.get("out_path")
    if main_path and os.path.splitext(os.fspath(main_path))[1].lower() != ".docx":
        raise UnsafeOutputPathError("主输出必须使用 .docx 扩展名")


def allocate_temp_sibling(final_path):
    """在最终文件同目录分配临时路径，确保 os.replace 属于同卷原子替换。"""
    final_path = os.path.abspath(os.fspath(final_path))
    directory = os.path.dirname(final_path)
    os.makedirs(directory, exist_ok=True)
    stem, extension = os.path.splitext(os.path.basename(final_path))
    fd, temp_path = tempfile.mkstemp(
        prefix=f".{stem}-", suffix=f".tmp{extension}", dir=directory)
    os.close(fd)
    os.unlink(temp_path)
    return temp_path


def commit_atomic(pairs):
    """逐个原子替换；调用方应把主 DOCX 放最后，使它成为完成标志。"""
    for temp_path, final_path in pairs:
        os.replace(temp_path, final_path)


class AtomicOutputSet:
    """为一组最终产物分配同目录临时文件，并在成功时统一提交。"""

    def __init__(self, final_paths):
        self.final_paths = dict(final_paths)
        self.temp_paths = {
            key: allocate_temp_sibling(path)
            for key, path in self.final_paths.items()
        }
        self._finalizer = weakref.finalize(
            self, cleanup_temps, list(self.temp_paths.values()))

    def temp(self, key):
        return self.temp_paths[key]

    def commit(self, main_key="out_path"):
        order = [key for key in self.final_paths if key != main_key] + [main_key]
        commit_atomic([
            (self.temp_paths[key], self.final_paths[key]) for key in order
        ])
        self._finalizer.detach()


def cleanup_temps(paths):
    for path in paths:
        try:
            if path and os.path.exists(path):
                os.unlink(path)
        except OSError:
            pass


def write_json_atomic(path, value):
    temp_path = allocate_temp_sibling(path)
    try:
        with open(temp_path, "w", encoding="utf-8") as handle:
            json.dump(value, handle, ensure_ascii=False, indent=2)
        os.replace(temp_path, path)
    finally:
        cleanup_temps([temp_path])
