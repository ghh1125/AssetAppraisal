"""Field-level source selection and cross-material reconciliation.

The report mapping defines the authority per field.  This module keeps that
decision separate from extraction: adapters may discover candidates from any
material, but only the configured authoritative source may supply the value
written into Word.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from copy import deepcopy
import re
from typing import Any, Iterable


@dataclass(frozen=True)
class SourceCandidate:
    value: Any
    source_kind: str
    source_file: str
    source_locator: str


@dataclass(frozen=True)
class SourceDiscrepancy:
    field_key: str
    selected_source_file: str
    selected_source_locator: str
    selected_value: Any
    other_source_file: str
    other_source_locator: str
    other_value: Any


@dataclass
class ReconciledField:
    field_key: str
    value: Any = None
    source_kind: str = "missing"
    source_file: str = ""
    source_locator: str = ""
    candidates: list[SourceCandidate] = field(default_factory=list)
    discrepancies: list[SourceDiscrepancy] = field(default_factory=list)
    issues: list[str] = field(default_factory=list)


def _number(value: Any) -> Decimal | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float, Decimal)):
        return Decimal(str(value))
    text = str(value or "").strip().replace(",", "").replace("，", "")
    if not text or not re.fullmatch(r"[-+]?\d+(?:\.\d+)?", text):
        return None
    try:
        return Decimal(text)
    except InvalidOperation:
        return None


def _same_value(left: Any, right: Any) -> bool:
    if isinstance(left, dict) and isinstance(right, dict):
        return set(left) == set(right) and all(
            _same_value(left[key], right[key]) for key in left
        )
    if isinstance(left, (list, tuple)) and isinstance(right, (list, tuple)):
        return len(left) == len(right) and all(
            _same_value(left_item, right_item)
            for left_item, right_item in zip(left, right, strict=True)
        )
    left_number, right_number = _number(left), _number(right)
    if left_number is not None and right_number is not None:
        return left_number == right_number
    return str(left or "").strip() == str(right or "").strip()


def _has_material_value(value: Any) -> bool:
    """Reject template-shaped matrices containing only unresolved markers."""
    if isinstance(value, dict):
        rows = value.get("rows")
        if isinstance(rows, list):
            # First row commonly contains only headers; first column in later
            # rows is a label.  Neither is evidence that an amount was found.
            return any(
                _has_material_value(cell)
                for row in rows[1:]
                if isinstance(row, list)
                for cell in row[1:]
            )
        return any(_has_material_value(item) for item in value.values())
    if isinstance(value, (list, tuple)):
        return any(_has_material_value(item) for item in value)
    text = str(value if value is not None else "").strip()
    return text not in {"", "XXX", "XX", "20XX"}


def _row_key(row: list[Any]) -> str:
    if not row:
        return ""
    return re.sub(r"[：:\s]", "", str(row[0] or "")).strip()


def _merge_missing_table_cells(primary: Any, fallback: Any) -> Any:
    """Use Excel only for cells the authoritative PDF OCR did not capture.

    A single OCR hit makes a table a valid PDF candidate, but it does not mean
    every row was recognised.  Preserve every material PDF cell and fill only
    marker/blank cells from the same labelled Excel row.
    """
    if not isinstance(primary, dict) or not isinstance(fallback, dict):
        return primary
    primary_rows = primary.get("rows")
    fallback_rows = fallback.get("rows")
    if not isinstance(primary_rows, list) or not isinstance(fallback_rows, list):
        return primary
    first_row_is_header = bool(
        fallback_rows
        and isinstance(fallback_rows[0], list)
        and _row_key(fallback_rows[0]) in {"项目", "项目报表日", "项目报表年度"}
    )
    data_start = 1 if first_row_is_header else 0
    fallback_by_label = {
        _row_key(row): row
        for row in fallback_rows[data_start:]
        if isinstance(row, list) and _row_key(row)
    }
    if not fallback_by_label:
        return primary
    merged = deepcopy(primary)
    merged_rows = merged.get("rows", [])
    for row in merged_rows[data_start:]:
        if not isinstance(row, list):
            continue
        fallback_row = fallback_by_label.get(_row_key(row))
        if not fallback_row:
            continue
        for index in range(1, min(len(row), len(fallback_row))):
            if not _has_material_value(row[index]) and _has_material_value(fallback_row[index]):
                row[index] = fallback_row[index]
    return merged


def reconcile_field(
    *,
    field_key: str,
    candidates: Iterable[SourceCandidate],
    source_priority: tuple[str, ...],
    require_primary_source: bool,
    allow_fallback_when_primary_missing: bool = False,
    primary_source_supplied: bool = False,
) -> ReconciledField:
    """Select the authoritative candidate and retain every disagreement.

    ``require_primary_source`` is used for audit-derived fields.  It prevents
    a valuation workbook from silently filling an audit-report field when no
    audit PDF was supplied.
    """
    def missing_primary_issue() -> str:
        if primary_source_supplied:
            return f"{field_key}：审计PDF已上传，但OCR未识别到该字段"
        return f"{field_key}：未上传审计PDF，无法获取审计数据"

    usable = [item for item in candidates if _has_material_value(item.value)]
    result = ReconciledField(field_key=field_key, candidates=usable)
    if not usable:
        if require_primary_source and source_priority and source_priority[0] == "pdf_ocr":
            result.issues.append(missing_primary_issue())
        return result

    selected = next(
        (
            item
            for source_kind in source_priority
            for item in usable
            if item.source_kind == source_kind
        ),
        None,
    )
    if selected is None:
        if require_primary_source and source_priority and source_priority[0] == "pdf_ocr":
            result.issues.append(missing_primary_issue())
        return result

    if require_primary_source and source_priority and selected.source_kind != source_priority[0]:
        if source_priority[0] == "pdf_ocr":
            if not allow_fallback_when_primary_missing:
                result.issues.append(missing_primary_issue())
                return result
            result.issues.append(
                f"{missing_primary_issue()}；"
                f"已采用{selected.source_file}（{selected.source_locator}），暂未完成PDF对照"
            )

    result.value = selected.value
    result.source_kind = selected.source_kind
    result.source_file = selected.source_file
    result.source_locator = selected.source_locator
    for candidate in usable:
        if candidate == selected or _same_value(selected.value, candidate.value):
            continue
        result.discrepancies.append(
            SourceDiscrepancy(
                field_key=field_key,
                selected_source_file=selected.source_file,
                selected_source_locator=selected.source_locator,
                selected_value=selected.value,
                other_source_file=candidate.source_file,
                other_source_locator=candidate.source_locator,
                other_value=candidate.value,
            )
        )
    if isinstance(selected.value, dict) and isinstance(selected.value.get("rows"), list):
        fallback = next(
            (
                item
                for item in usable
                if item != selected and isinstance(item.value, dict)
                and isinstance(item.value.get("rows"), list)
            ),
            None,
        )
        if fallback is not None:
            result.value = _merge_missing_table_cells(selected.value, fallback.value)
    return result
