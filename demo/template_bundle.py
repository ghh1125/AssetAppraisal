"""Resolve the versioned, company-neutral workbook template bundle.

The two appraisal workbooks are application resources.  They must not be
loaded from a customer-material directory or selected by a customer filename.
"""

from __future__ import annotations

from dataclasses import dataclass
import hashlib
import json
import os
from pathlib import Path
import re
from typing import Any
from xml.etree import ElementTree
from zipfile import BadZipFile, ZipFile


DEFAULT_TEMPLATE_BUNDLE_DIR = Path(__file__).resolve().parent / "workbook_templates"
TEMPLATE_BUNDLE_ENV = "APPRAISAL_WORKBOOK_TEMPLATE_DIR"
TEMPLATE_BUNDLE_MANIFEST = "workbook_templates.json"


@dataclass(frozen=True)
class WorkbookTemplateBundle:
    version: str
    root: Path
    asset: Path
    income: Path


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _role_path(root: Path, role: str, item: Any) -> Path:
    if not isinstance(item, dict):
        raise RuntimeError(f"工作簿模板清单缺少 {role} 角色")
    filename = str(item.get("file") or "").strip()
    if not filename:
        raise RuntimeError(f"工作簿模板清单未配置 {role} 文件")
    path = (root / filename).resolve()
    if path.parent != root or path.suffix.lower() != ".xlsx":
        raise RuntimeError(f"工作簿模板清单中的 {role} 路径不安全或格式不正确")
    if not path.is_file():
        raise RuntimeError(f"工作簿模板缺失：{path}")
    expected_hash = str(item.get("sha256") or "").strip().lower()
    if not expected_hash:
        raise RuntimeError(f"工作簿模板清单未登记 {role} 的 SHA-256")
    actual_hash = _sha256(path)
    if actual_hash != expected_hash:
        raise RuntimeError(f"工作簿模板完整性校验失败：{path.name}")
    _validate_structure(path, role, item)
    return path


def _validate_structure(path: Path, role: str, item: dict[str, Any]) -> None:
    try:
        with ZipFile(path) as archive:
            workbook_xml = archive.read("xl/workbook.xml")
            worksheet_names = [
                name
                for name in archive.namelist()
                if name.startswith("xl/worksheets/") and name.endswith(".xml")
            ]
            worksheet_xml = [archive.read(name) for name in worksheet_names]
    except (BadZipFile, KeyError, OSError) as exc:
        raise RuntimeError(f"工作簿模板格式无法打开：{path.name}") from exc
    try:
        root = ElementTree.fromstring(workbook_xml)
    except ElementTree.ParseError as exc:
        raise RuntimeError(f"工作簿模板核心结构无法解析：{path.name}") from exc
    namespace = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    sheet_names = [str(sheet.attrib.get("name") or "") for sheet in root.findall(".//x:sheet", namespace)]
    try:
        expected_sheet_count = int(item.get("sheet_count") or 0)
        expected_formula_count = int(item.get("formula_count") or 0)
    except (TypeError, ValueError) as exc:
        raise RuntimeError(f"工作簿模板清单的结构计数不正确：{role}/{path.name}") from exc
    if expected_sheet_count <= 0 or len(sheet_names) != expected_sheet_count:
        raise RuntimeError(f"工作簿模板 Sheet 数量不符合清单：{role}/{path.name}")
    required_sheets = item.get("required_sheets")
    if not isinstance(required_sheets, list) or not required_sheets:
        raise RuntimeError(f"工作簿模板清单未登记 {role} 的必要 Sheet")
    missing_sheets = sorted(set(map(str, required_sheets)).difference(sheet_names))
    if missing_sheets:
        raise RuntimeError(f"工作簿模板缺少必要 Sheet：{role}/{'、'.join(missing_sheets)}")
    formula_count = sum(len(re.findall(br"<f(?:\s|>)", payload)) for payload in worksheet_xml)
    if expected_formula_count <= 0 or formula_count != expected_formula_count:
        raise RuntimeError(f"工作簿模板公式数量不符合清单：{role}/{path.name}")
    required_placeholders = item.get("required_placeholders")
    if not isinstance(required_placeholders, list) or not required_placeholders:
        raise RuntimeError(f"工作簿模板清单未登记 {role} 的身份占位符")
    searchable = b"\n".join(worksheet_xml).decode("utf-8", errors="ignore")
    missing_placeholders = [str(token) for token in required_placeholders if str(token) not in searchable]
    if missing_placeholders:
        raise RuntimeError(f"工作簿模板缺少身份占位符：{role}/{'、'.join(missing_placeholders)}")


def load_workbook_template_bundle(template_dir: Path | None = None) -> WorkbookTemplateBundle:
    """Load one explicit bundle, an environment override, or the built-in bundle."""
    configured = template_dir
    if configured is None:
        env_value = str(os.environ.get(TEMPLATE_BUNDLE_ENV) or "").strip()
        configured = Path(env_value) if env_value else DEFAULT_TEMPLATE_BUNDLE_DIR
    root = configured.expanduser().resolve()
    manifest_path = root / TEMPLATE_BUNDLE_MANIFEST
    if not manifest_path.is_file():
        raise RuntimeError(
            f"工作簿模板包缺少 {TEMPLATE_BUNDLE_MANIFEST}：{root}。"
            "请使用仓库内置模板包，或配置一个带完整性清单的自定义模板包。"
        )
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"工作簿模板清单无法读取：{manifest_path}") from exc
    version = str(manifest.get("version") or "").strip()
    if not version:
        raise RuntimeError("工作簿模板清单缺少 version")
    roles = manifest.get("roles")
    if not isinstance(roles, dict):
        raise RuntimeError("工作簿模板清单缺少 roles")
    return WorkbookTemplateBundle(
        version=version,
        root=root,
        asset=_role_path(root, "asset", roles.get("asset")),
        income=_role_path(root, "income", roles.get("income")),
    )
