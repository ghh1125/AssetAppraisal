from pathlib import Path
from zipfile import ZipFile

import pytest

from demo.adapters.archive_intake import safe_extract_archive


def test_safe_extract_zip_preserves_nested_materials(tmp_path: Path):
    archive_path = tmp_path / "materials.zip"
    with ZipFile(archive_path, "w") as archive:
        archive.writestr("审计报告/报告.pdf", b"%PDF")
        archive.writestr("营业执照/执照.png", b"png")

    output = tmp_path / "extracted"
    paths = safe_extract_archive(archive_path, output)

    assert {path.relative_to(output).as_posix() for path in paths} == {
        "审计报告/报告.pdf",
        "营业执照/执照.png",
    }


def test_safe_extract_zip_rejects_path_traversal(tmp_path: Path):
    archive_path = tmp_path / "unsafe.zip"
    with ZipFile(archive_path, "w") as archive:
        archive.writestr("../outside.txt", b"unsafe")

    with pytest.raises(ValueError, match="越界路径"):
        safe_extract_archive(archive_path, tmp_path / "extracted")
    assert not (tmp_path / "outside.txt").exists()
