"""Prepare generated XLSX files for delivery without desktop Excel.

The reference templates can contain stale cached error values even when their
formulas are valid. This module validates OOXML, rejects structural formula
errors, clears only stale cached formula errors, and requests a full
calculation when the reviewer opens the file. It never launches Excel.
"""

from __future__ import annotations

import html
import os
from pathlib import Path
import re
import tempfile
from typing import Any
from zipfile import ZIP_DEFLATED, BadZipFile, ZipFile


_WORKSHEET_PREFIX = "xl/worksheets/"
_FORMULA_ERROR_VALUES = (
    "#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!",
)
_STRUCTURAL_FORMULA_ERRORS = ("#REF!", "#NAME?")
_CELL_RE = re.compile(r"(<c\b[^>]*>)(.*?)(</c>)", re.DOTALL)
_FORMULA_RE = re.compile(r"<f(?:\s[^>]*)?>(.*?)</f>", re.DOTALL)
_VALUE_RE = re.compile(r"<v>(.*?)</v>", re.DOTALL)
_CALC_PR_RE = re.compile(r"<calcPr\b[^>]*/>")


def _sheet_count(workbook_xml: str) -> int:
    return len(re.findall(r"<sheet\b", workbook_xml))


def workbook_sheet_count(path: Path) -> int:
    """Read the workbook sheet count without mutating the package."""
    try:
        with ZipFile(path, "r") as workbook:
            return _sheet_count(workbook.read("xl/workbook.xml").decode("utf-8"))
    except (BadZipFile, KeyError, UnicodeDecodeError) as exc:
        raise RuntimeError(f"工作簿格式无法打开：{path.name}") from exc


def _force_full_calculation(workbook_xml: str) -> str:
    calc_pr = '<calcPr calcId="0" calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>'
    if _CALC_PR_RE.search(workbook_xml):
        return _CALC_PR_RE.sub(calc_pr, workbook_xml, count=1)
    return workbook_xml.replace("</workbook>", f"{calc_pr}</workbook>")


def _sanitize_worksheet_xml(xml: str, *, sheet_path: str) -> tuple[str, int]:
    cleared = 0

    def replace_cell(match: re.Match[str]) -> str:
        nonlocal cleared
        opening, body, closing = match.groups()
        formula_match = _FORMULA_RE.search(body)
        value_match = _VALUE_RE.search(body)
        if formula_match:
            formula = html.unescape(formula_match.group(1)).upper()
            for token in _STRUCTURAL_FORMULA_ERRORS:
                if token in formula:
                    raise RuntimeError(f"公式存在结构性错误：{sheet_path} 包含 {token}")
        if not value_match:
            return match.group(0)
        cached_value = html.unescape(value_match.group(1)).strip().upper()
        if cached_value not in _FORMULA_ERROR_VALUES:
            return match.group(0)
        if not formula_match:
            raise RuntimeError(f"单元格保存了非公式错误值：{sheet_path} 包含 {cached_value}")
        opening = re.sub(r'\s+t="e"', "", opening)
        body = _VALUE_RE.sub("", body, count=1)
        cleared += 1
        return f"{opening}{body}{closing}"

    return _CELL_RE.sub(replace_cell, xml), cleared


def prepare_workbook_for_delivery(path: Path) -> dict[str, Any]:
    """Validate and update one generated workbook in place."""
    path = path.resolve()
    if not path.is_file():
        raise RuntimeError(f"生成的工作簿不存在：{path.name}")
    try:
        with ZipFile(path, "r") as source:
            bad_member = source.testzip()
            if bad_member:
                raise RuntimeError(f"工作簿压缩包校验失败：{path.name}/{bad_member}")
            members = source.infolist()
            payloads = {member.filename: source.read(member.filename) for member in members}
    except BadZipFile as exc:
        raise RuntimeError(f"工作簿格式无法打开：{path.name}") from exc

    workbook_key = "xl/workbook.xml"
    if workbook_key not in payloads:
        raise RuntimeError(f"工作簿缺少核心结构：{path.name}")
    workbook_xml = payloads[workbook_key].decode("utf-8")
    count = _sheet_count(workbook_xml)
    payloads[workbook_key] = _force_full_calculation(workbook_xml).encode("utf-8")

    cleared = 0
    for member_name, payload in list(payloads.items()):
        if not member_name.startswith(_WORKSHEET_PREFIX) or not member_name.endswith(".xml"):
            continue
        sanitized, member_cleared = _sanitize_worksheet_xml(
            payload.decode("utf-8"),
            sheet_path=member_name,
        )
        payloads[member_name] = sanitized.encode("utf-8")
        cleared += member_cleared

    with tempfile.NamedTemporaryFile(
        prefix=f".{path.stem}-",
        suffix=".xlsx",
        dir=path.parent,
        delete=False,
    ) as temporary:
        temporary_path = Path(temporary.name)
    try:
        with ZipFile(temporary_path, "w", compression=ZIP_DEFLATED) as target:
            for member in members:
                target.writestr(member, payloads[member.filename])
        os.replace(temporary_path, path)
    finally:
        temporary_path.unlink(missing_ok=True)

    return {
        "path": str(path),
        "sheet_count": count,
        "cached_formula_errors_cleared": cleared,
        "full_calculation_on_open": True,
    }


def prepare_generated_workbooks(output_dir: Path) -> dict[str, Any]:
    """Prepare exactly the two client workbooks in a generated case folder."""
    results = []
    for path in sorted(output_dir.resolve().glob("*.xlsx"), key=lambda item: item.stat().st_size):
        if workbook_sheet_count(path) not in {29, 44}:
            continue
        results.append(prepare_workbook_for_delivery(path))
    if len(results) != 2:
        raise RuntimeError(f"应处理两份客户工作簿，实际识别到 {len(results)} 份")
    return {
        "status": "completed",
        "engine": "backend_ooxml",
        "workbook_count": len(results),
        "cached_formula_errors_cleared": sum(
            item["cached_formula_errors_cleared"] for item in results
        ),
        "workbooks": results,
    }
