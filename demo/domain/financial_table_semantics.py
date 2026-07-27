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
_ELECTRONIC_EQUIPMENT_ALIASES = (
    "电子设备",
    "办公电子设备",
    "电脑及电子设备",
    "电力电子设备",
    "工器具及电子设备",
)


def appraisal_zero_is_unfinished(
    *,
    book_value: float | None,
    appraised_value: float | None,
    appraisal_column_values: Sequence[object],
) -> bool:
    """Treat an isolated all-zero appraisal column as not yet appraised."""
    if book_value in (None, 0) or appraised_value != 0:
        return False
    numeric = [
        float(value)
        for value in appraisal_column_values
        if isinstance(value, (int, float)) and not isinstance(value, bool)
    ]
    return not any(value != 0 for value in numeric)


def _normalized_label(value: object) -> str:
    return re.sub(
        r"[\s：:()（）一二三四五六七八九十、．.]+",
        "",
        str(value or "").replace("帳", "账").replace("帐", "账"),
    )


def canonical_long_term_asset_category(value: object) -> str | None:
    """Normalize only categories that unambiguously mean electronic equipment."""
    label = _normalized_label(value)
    if any(alias in label for alias in _ELECTRONIC_EQUIPMENT_ALIASES):
        return "电子设备"
    return None


def detail_header_role(value: object) -> str | None:
    """Return the business role of a fixed-asset detail column."""
    label = _normalized_label(value)
    if not label:
        return None
    if label in {"资产编号", "固定资产编号", "设备编号", "卡片编号"}:
        return "asset_id"
    if label in {"资产名称", "固定资产名称", "设备名称"}:
        return "asset_name"
    if label in {"资产类别", "固定资产类别", "设备类别", "资产分类", "分类"}:
        return "category"
    if "评估" in label and ("价值" in label or label.endswith("值")):
        return "appraised"
    if "账面" in label and any(token in label for token in ("净值", "净额")):
        return "book_net"
    if "账面" in label and ("原值" in label or "原始价值" in label):
        return "book_cost"
    if label in {"账面价值", "账面金额"}:
        return "book_net"
    if "累计折旧" in label:
        return "depreciation"
    return None


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
