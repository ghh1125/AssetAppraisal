from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path


class LibreOfficeTemplatePageReader:
    """将只读 Word 模板渲染为 PDF，并提取每个原模板页的文字。"""

    def __init__(self, executable: str | None = None):
        self.executable = executable or os.environ.get("LIBREOFFICE_BIN") or shutil.which("soffice")

    def extract(self, template_path: Path) -> tuple[list[str], list[str]]:
        if not self.executable:
            return [], ["原模板页码提取失败：未找到 LibreOffice/soffice"]
        try:
            import fitz
        except ImportError as exc:
            return [], [f"原模板页码提取失败：未安装 PyMuPDF（{exc}）"]
        try:
            with tempfile.TemporaryDirectory(prefix="template-pages-") as temporary:
                root = Path(temporary)
                profile = root / "profile"
                output = root / "output"
                profile.mkdir()
                output.mkdir()
                command = [
                    self.executable,
                    f"-env:UserInstallation={profile.as_uri()}",
                    "--invisible",
                    "--headless",
                    "--norestore",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    str(output),
                    str(template_path),
                ]
                environment = os.environ.copy()
                environment["HOME"] = str(profile)
                environment["TMPDIR"] = "/private/tmp" if Path("/private/tmp").is_dir() else temporary
                completed = subprocess.run(
                    command,
                    check=False,
                    capture_output=True,
                    text=True,
                    env=environment,
                    timeout=120,
                )
                pdfs = sorted(output.glob("*.pdf"))
                if completed.returncode or not pdfs:
                    detail = (completed.stderr or completed.stdout or "未生成 PDF").strip()
                    return [], [f"原模板页码提取失败：{detail}"]
                with fitz.open(pdfs[0]) as document:
                    return [page.get_text("text") for page in document], []
        except Exception as exc:
            return [], [f"原模板页码提取失败：{exc}"]
