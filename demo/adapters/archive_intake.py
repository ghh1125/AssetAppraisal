"""Safe extraction helpers for uploaded audit-material archives."""

from __future__ import annotations

from pathlib import Path, PurePosixPath
import re
import shutil
import stat
import subprocess
from typing import Any, Iterable
from zipfile import ZipFile, ZipInfo


MAX_ARCHIVE_MEMBERS = 5_000
MAX_ARCHIVE_UNCOMPRESSED_BYTES = 4 * 1024 * 1024 * 1024


def _safe_relative_path(name: str) -> Path:
    normalized = str(name or "").replace("\\", "/")
    if not normalized or "\x00" in normalized:
        raise ValueError("压缩包包含无效文件名")
    pure = PurePosixPath(normalized)
    if pure.is_absolute() or ".." in pure.parts or re.match(r"^[A-Za-z]:", normalized):
        raise ValueError(f"压缩包包含越界路径：{name}")
    safe_parts = [part for part in pure.parts if part not in {"", "."}]
    if not safe_parts:
        raise ValueError("压缩包包含无效空路径")
    return Path(*safe_parts)


def _validate_member_set(members: Iterable[tuple[str, int, bool]]) -> list[tuple[Path, int, bool]]:
    checked: list[tuple[Path, int, bool]] = []
    total_size = 0
    for index, (name, size, is_link) in enumerate(members, start=1):
        if index > MAX_ARCHIVE_MEMBERS:
            raise ValueError(f"压缩包文件数超过上限 {MAX_ARCHIVE_MEMBERS}")
        if is_link:
            raise ValueError(f"压缩包不允许符号链接：{name}")
        relative = _safe_relative_path(name)
        total_size += max(0, int(size or 0))
        if total_size > MAX_ARCHIVE_UNCOMPRESSED_BYTES:
            raise ValueError("压缩包解压后总大小超过 4GB 上限")
        checked.append((relative, int(size or 0), is_link))
    return checked


def _extract_zip(archive_path: Path, destination: Path) -> list[Path]:
    with ZipFile(archive_path) as archive:
        infos: list[ZipInfo] = archive.infolist()
        checked = _validate_member_set(
            (
                info.filename,
                info.file_size,
                stat.S_IFMT(info.external_attr >> 16) == stat.S_IFLNK,
            )
            for info in infos
        )
        extracted: list[Path] = []
        for info, (relative, _size, _is_link) in zip(infos, checked):
            target = destination / relative
            if info.is_dir():
                target.mkdir(parents=True, exist_ok=True)
                continue
            target.parent.mkdir(parents=True, exist_ok=True)
            with archive.open(info) as source, target.open("wb") as sink:
                shutil.copyfileobj(source, sink, length=1024 * 1024)
            extracted.append(target)
        return extracted


def _extract_rar(archive_path: Path, destination: Path) -> list[Path]:
    try:
        import rarfile
    except ImportError as exc:  # pragma: no cover - depends on deployment extras
        raise RuntimeError("RAR 解压组件未安装，请安装项目依赖 rarfile，或改传 ZIP") from exc

    # Windows ships libarchive as ``tar.exe``.  rarfile knows how to drive
    # bsdtar, but its default executable name is ``bsdtar``; point the public
    # tool setting at the Windows binary when no dedicated bsdtar is present.
    if not shutil.which(str(getattr(rarfile, "BSDTAR_TOOL", "bsdtar"))) and shutil.which("tar"):
        rarfile.BSDTAR_TOOL = shutil.which("tar")
        rarfile.tool_setup(
            unrar=False,
            unar=False,
            bsdtar=True,
            sevenzip=False,
            sevenzip2=False,
            force=True,
        )

    with rarfile.RarFile(archive_path) as archive:
        infos: list[Any] = archive.infolist()
        checked = _validate_member_set(
            (
                info.filename,
                info.file_size,
                bool(getattr(info, "is_symlink", lambda: False)()) or bool(getattr(info, "file_redir", None)),
            )
            for info in infos
        )
        bsdtar = shutil.which("bsdtar") or shutil.which("tar")
        if bsdtar:
            completed = subprocess.run(
                [bsdtar, "-xf", str(archive_path), "-C", str(destination)],
                check=False,
                capture_output=True,
                text=True,
            )
            if completed.returncode != 0:
                detail = (completed.stderr or completed.stdout or "未知错误").strip()
                raise RuntimeError(f"RAR 解压失败：{detail[:500]}")
            return [
                destination / relative
                for info, (relative, _size, _is_link) in zip(infos, checked)
                if not info.isdir()
            ]
        extracted: list[Path] = []
        for info, (relative, _size, _is_link) in zip(infos, checked):
            target = destination / relative
            if info.isdir():
                target.mkdir(parents=True, exist_ok=True)
                continue
            target.parent.mkdir(parents=True, exist_ok=True)
            with archive.open(info) as source, target.open("wb") as sink:
                shutil.copyfileobj(source, sink, length=1024 * 1024)
            extracted.append(target)
        return extracted


def safe_extract_archive(archive_path: Path, destination: Path) -> list[Path]:
    """Extract ZIP/RAR without trusting member paths or symbolic links."""
    suffix = archive_path.suffix.lower()
    if suffix not in {".zip", ".rar"}:
        raise ValueError("材料包仅支持 .rar 或 .zip 格式")
    destination.mkdir(parents=True, exist_ok=True)
    if suffix == ".zip":
        return _extract_zip(archive_path, destination)
    return _extract_rar(archive_path, destination)
