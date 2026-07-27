from __future__ import annotations

from copy import deepcopy
import re
from typing import Any


_UNRESOLVED = re.compile(r"^X{2,}$", re.IGNORECASE)
_CALENDAR_PERIOD = re.compile(r"20\d{2}")


def _is_unresolved(value: Any) -> bool:
    text = str(value or "").strip()
    return not text or bool(_UNRESOLVED.fullmatch(text))


def _rows(table: dict[str, Any]) -> list[list[Any]]:
    rows = table.get("rows", [])
    return rows if isinstance(rows, list) else []


def _headers(table: dict[str, Any]) -> list[str]:
    rows = _rows(table)
    if not rows or not isinstance(rows[0], list):
        return []
    return [str(value or "").strip() for value in rows[0][1:]]


def _coverage(table: dict[str, Any]) -> int:
    return sum(
        not _is_unresolved(value)
        for row in _rows(table)[1:]
        if isinstance(row, list)
        for value in row[1:]
    )


def _period_specificity(table: dict[str, Any]) -> int:
    return sum(bool(_CALENDAR_PERIOD.search(header)) for header in _headers(table))


def _quality(table: dict[str, Any]) -> tuple[int, int]:
    return _coverage(table), _period_specificity(table)


def merge_historical_tables(
    existing: dict[str, Any],
    candidate: dict[str, Any],
) -> dict[str, Any]:
    """Merge compatible history tables or keep the better evidenced layout."""
    if not _rows(existing):
        return deepcopy(candidate)
    if not _rows(candidate):
        return deepcopy(existing)
    if _headers(existing) != _headers(candidate):
        return deepcopy(candidate if _quality(candidate) > _quality(existing) else existing)

    merged = deepcopy(existing)
    merged_rows = _rows(merged)
    candidate_by_label = {
        str(row[0] or "").strip(): row
        for row in _rows(candidate)[1:]
        if isinstance(row, list) and row
    }
    known_labels: set[str] = set()
    for row in merged_rows[1:]:
        if not isinstance(row, list) or not row:
            continue
        label = str(row[0] or "").strip()
        known_labels.add(label)
        other = candidate_by_label.get(label)
        if not other:
            continue
        width = max(len(row), len(other))
        row.extend(["XXX"] * (width - len(row)))
        for index in range(1, width):
            other_value = other[index] if index < len(other) else "XXX"
            if _is_unresolved(row[index]) and not _is_unresolved(other_value):
                row[index] = other_value
    for label, row in candidate_by_label.items():
        if label not in known_labels:
            merged_rows.append(deepcopy(row))
    return merged
