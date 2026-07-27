from __future__ import annotations

import re
from collections.abc import Mapping, Sequence
from datetime import date, datetime
from typing import TypeAlias


CanonicalPeriod: TypeAlias = tuple[int | None, int | None, int | None, str]

_NON_HISTORICAL_TOKENS = (
    "增长",
    "增幅",
    "占比",
    "比例",
    "预测",
    "预算",
    "预计",
    "估算",
    "序号",
    "%",
)
_RELATIVE_PERIODS = {
    "期初数",
    "期末数",
    "年初数",
    "年末数",
    "本期数",
    "上期数",
    "本年累计",
    "上年同期",
    "评估基准期",
    "评估基准日",
    "基准期",
    "基准日",
}


def canonical_period(value: object) -> CanonicalPeriod | None:
    """Return a sortable business period, or ``None`` for non-period headers."""
    if isinstance(value, datetime):
        return value.year, value.month, value.day, value.strftime("%Y年%m月%d日")
    if isinstance(value, date):
        return value.year, value.month, value.day, value.strftime("%Y年%m月%d日")
    if isinstance(value, (int, float)) or value is None:
        return None

    text = re.sub(r"\s+", "", str(value).strip())
    if not text or any(token in text for token in _NON_HISTORICAL_TOKENS):
        return None
    if text in _RELATIVE_PERIODS or any(
        token in text for token in ("评估基准期", "评估基准日", "基准期", "基准日")
    ):
        return None, None, None, text

    match = re.search(
        r"(?P<year>20\d{2})(?:年|[./-])"
        r"(?:(?P<month>\d{1,2})(?:月|[./-]))?"
        r"(?:(?P<day>\d{1,2})日?)?",
        text,
    )
    if match:
        year = int(match.group("year"))
        month = int(match.group("month") or 12)
        day = int(match.group("day") or 31)
        return year, month, day, text

    annual = re.fullmatch(r"(20\d{2})(?:年度|年)?", text)
    if annual:
        return int(annual.group(1)), 12, 31, text
    return None


def choose_historical_columns(
    headers: Mapping[int, Sequence[object]],
    *,
    valuation_year: int | None = None,
    candidate_columns: Sequence[int] | None = None,
    limit: int = 3,
) -> list[int]:
    """Choose actual historical amount columns from reconstructed headers."""
    allowed = set(candidate_columns) if candidate_columns is not None else None
    dated: list[tuple[tuple[int, int, int], int]] = []
    relative: list[int] = []
    for column, parts in headers.items():
        if allowed is not None and column not in allowed:
            continue
        combined = "".join(str(part or "").strip() for part in parts)
        if not combined or any(token in combined for token in _NON_HISTORICAL_TOKENS):
            continue
        periods = [period for part in parts if (period := canonical_period(part))]
        if not periods:
            period = canonical_period(combined)
            periods = [period] if period else []
        if not periods:
            continue
        period = next((item for item in periods if item[0] is not None), periods[-1])
        year, month, day, _ = period
        if year is None:
            relative.append(column)
            continue
        if valuation_year is not None and year > valuation_year:
            continue
        dated.append(((year, month or 12, day or 31), column))

    if dated:
        ordered = [column for _, column in sorted(dated)]
        if relative:
            ordered = sorted(set(ordered) | set(relative))
        return ordered[-limit:]
    return sorted(relative)[-limit:]
