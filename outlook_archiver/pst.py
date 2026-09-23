from __future__ import annotations

import ctypes
import hashlib
import os
import shutil
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

from .utils import read_json, safe_name, sha256_file, sha256_text, write_json_atomic


@dataclass(slots=True)
class PstCopy:
    source: Path
    copy: Path
    reused: bool
    sha256: str


def default_work_root() -> Path:
    local = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local"))
    return local / "OutlookEmailArchiver" / "PSTWork"


def work_root_for_output(output_root: str | Path) -> Path:
    """Place reusable PST copies beside the archive root, on the same drive."""
    output = Path(output_root).expanduser().resolve()
    parent = output if output.parent == output else output.parent
    return parent / "PSTWork"


@contextmanager
def _exclusive_reader(path: Path):
    if os.name != "nt":
        with path.open("rb") as stream:
            yield stream
        return
    import msvcrt

    GENERIC_READ = 0x80000000
    OPEN_EXISTING = 3
    FILE_ATTRIBUTE_NORMAL = 0x80
    invalid = ctypes.c_void_p(-1).value
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateFileW.argtypes = [
        ctypes.c_wchar_p, ctypes.c_uint32, ctypes.c_uint32, ctypes.c_void_p,
        ctypes.c_uint32, ctypes.c_uint32, ctypes.c_void_p,
    ]
    kernel32.CreateFileW.restype = ctypes.c_void_p
    handle = kernel32.CreateFileW(str(path), GENERIC_READ, 0, None, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, None)
    if handle == invalid:
        error = ctypes.get_last_error()
        if error in (32, 33):
            raise RuntimeError(
                f"源 PST 正被其他进程占用，且没有可复用的已验证工作副本：{path}。"
                "程序尚未启动 Outlook。请检查任务管理器中的 Outlook 及复制/备份程序；"
                "如果 Outlook 已运行，可在其中关闭该数据文件后重试。"
                f"（Windows 错误 {error}）"
            )
        raise OSError(error, f"无法读取源 PST：{ctypes.FormatError(error).strip()}", str(path))
    try:
        descriptor = msvcrt.open_osfhandle(handle, os.O_RDONLY | os.O_BINARY)
        handle = None
        with os.fdopen(descriptor, "rb") as stream:
            yield stream
    finally:
        if handle not in (None, invalid):
            kernel32.CloseHandle(ctypes.c_void_p(handle))


def prepare_pst_copy(
    source_path: str | Path,
    *,
    work_root: Path | None = None,
    reuse_roots: Iterable[str | Path] = (),
    source_open_in_outlook: bool = False,
) -> PstCopy:
    source = Path(source_path).expanduser().resolve()
    if not source.is_file():
        raise FileNotFoundError(f"找不到 PST 文件：{source}")
    if source.suffix.lower() != ".pst":
        raise ValueError("请选择 .pst 文件。")
    stat = source.stat()
    root = (work_root or default_work_root()).resolve()
    identity = sha256_text(os.path.normcase(str(source)))[:16]
    directory = root / f"{safe_name(source.stem, 48)}_{identity}"
    directory.mkdir(parents=True, exist_ok=True)
    copy_path = directory / f"working_{safe_name(source.name, 80, 'archive.pst')}"
    manifest_path = directory / "copy-manifest.json"

    def verified_copy(candidate_directory: Path) -> PstCopy | None:
        candidate_copy = candidate_directory / copy_path.name
        candidate_manifest = candidate_directory / "copy-manifest.json"
        if not candidate_manifest.exists() or not candidate_copy.exists():
            return None
        try:
            manifest = read_json(candidate_manifest)
            same_source = (
                manifest.get("source_path_norm") == os.path.normcase(str(source))
                and int(manifest.get("source_size", -1)) == stat.st_size
                and int(manifest.get("source_mtime_ns", -1)) == stat.st_mtime_ns
            )
            same_copy_path = os.path.normcase(str(candidate_copy)) == os.path.normcase(
                str(Path(manifest.get("copy_path", "")).expanduser().resolve())
            )
            initial_hashes_match = bool(manifest.get("verified")) and (
                manifest.get("copy_sha256") == manifest.get("source_sha256")
            )
            copy_size = candidate_copy.stat().st_size
            reasonable_size = copy_size >= max(4, int(stat.st_size * 0.80))
            with candidate_copy.open("rb") as stream:
                pst_header = stream.read(4) == b"!BDN"
            if same_source and same_copy_path and initial_hashes_match and reasonable_size and pst_header:
                # Outlook may legitimately update indexes/internal state in the disposable work copy.
                # The strict source==copy SHA-256 check is therefore performed only when the copy is created.
                return PstCopy(source, candidate_copy, True, str(manifest.get("source_sha256", "")))
        except (OSError, ValueError, KeyError):
            return None
        return None

    candidate_directories = [directory]
    for reuse_root in reuse_roots:
        alternate = Path(reuse_root).expanduser().resolve() / directory.name
        if alternate not in candidate_directories:
            candidate_directories.append(alternate)
    for candidate_directory in candidate_directories:
        reused = verified_copy(candidate_directory)
        if reused is not None:
            return reused

    if source_open_in_outlook:
        raise RuntimeError(
            "源 PST 当前正被 Outlook 或其他程序使用，且没有可复用的已验证工作副本。"
            "请在经典 Outlook 中关闭该数据文件（不要退出 Outlook 也可以），确认没有复制/备份程序占用后重试。"
        )

    free = shutil.disk_usage(directory).free
    required = int(stat.st_size * 1.10) + 128 * 1024 * 1024
    if free < required:
        raise OSError(
            f"PST 工作副本空间不足：工作副本位置为 {directory}；"
            f"需要约 {required / 1024**3:.2f} GB（含安全余量），"
            f"当前可用 {free / 1024**3:.2f} GB。请在窗口中改选空间充足的工作副本位置。"
        )
    temporary = directory / f".{copy_path.name}.{os.getpid()}.partial"
    try:
        digest = hashlib.sha256()
        try:
            with _exclusive_reader(source) as input_stream, temporary.open("wb") as output_stream:
                for block in iter(lambda: input_stream.read(1024 * 1024), b""):
                    output_stream.write(block)
                    digest.update(block)
        except OSError as exc:
            raise RuntimeError(
                f"PST 工作副本复制失败。源文件：{source}；目标：{temporary}。"
                f"请检查文件权限、目标磁盘空间或磁盘状态。系统详情：{exc}"
            ) from exc
        shutil.copystat(source, temporary)
        source_hash = digest.hexdigest()
        copy_hash = sha256_file(temporary)
        if source_hash != copy_hash or temporary.stat().st_size != stat.st_size:
            raise OSError("PST 工作副本校验失败；临时副本未投入使用。")
        os.replace(temporary, copy_path)
        write_json_atomic(
            manifest_path,
            {
                "version": 2,
                "source_path": str(source),
                "source_path_norm": os.path.normcase(str(source)),
                "source_size": stat.st_size,
                "source_mtime_ns": stat.st_mtime_ns,
                "source_sha256": source_hash,
                "copy_path": str(copy_path),
                "copy_sha256": copy_hash,
                "verified": True,
                "reuse_note": "首次复制已逐字节校验；Outlook 挂载后工作副本内部状态可变化。",
            },
        )
        return PstCopy(source, copy_path, False, copy_hash)
    finally:
        if temporary.exists():
            temporary.unlink()
