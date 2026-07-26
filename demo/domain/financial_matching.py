from __future__ import annotations

import re
import unicodedata
from collections import defaultdict
from datetime import date
from decimal import Decimal, InvalidOperation
from typing import Any


def blank_configured_table(spec: dict[str, Any]) -> list[list[str]]:
    """Build a blank matrix from a configured financial table shape."""
    matrix: list[list[str]] = []
    header = [str(value) for value in spec.get("header", [])]
    if spec.get("include_header", True) and header:
        matrix.append(header)
    for row in spec.get("rows", []):
        cells = row.get("cells", [])
        matrix.append([str(row.get("label", "")), *([""] * len(cells))])
    return matrix


def normalize_label(value: str) -> str:
    text = unicodedata.normalize("NFKC", str(value)).strip()
    text = re.sub(r"^[（(]?[一二三四五六七八九十0-9]+[)）、.．:：]+", "", text)
    return re.sub(r"[\s:：,，。；;（）()]+", "", text)


def parse_period(value: str) -> str | None:
    text = unicodedata.normalize("NFKC", str(value)).strip()
    half_year = re.search(r"((?:19|20)\d{2})年(?:上半年|上半年度)", text)
    if half_year:
        return f"{half_year.group(1)}-06-30"
    full_year = re.search(r"((?:19|20)\d{2})年(?:度|全年)?", text)
    if full_year and "月" not in text and "日" not in text:
        return full_year.group(1)
    full_date = re.search(r"((?:19|20)\d{2})年(\d{1,2})月(\d{1,2})日", text)
    if full_date is None:
        full_date = re.search(r"((?:19|20)\d{2})[/.-](\d{1,2})[/.-](\d{1,2})", text)
    if full_date:
        year, month, day = map(int, full_date.groups())
        try:
            return date(year, month, day).isoformat()
        except ValueError:
            return None
    range_month = re.search(r"((?:19|20)\d{2})年\s*\d{1,2}\s*[-至~—]\s*(\d{1,2})月", text)
    if range_month:
        year, month = map(int, range_month.groups())
        if 1 <= month <= 12:
            next_month = date(year + (month == 12), 1 if month == 12 else month + 1, 1)
            return date.fromordinal(next_month.toordinal() - 1).isoformat()
    year_month = re.search(r"((?:19|20)\d{2})年\s*(\d{1,2})月", text)
    if year_month:
        year, month = map(int, year_month.groups())
        if 1 <= month <= 12:
            next_month = date(year + (month == 12), 1 if month == 12 else month + 1, 1)
            return date.fromordinal(next_month.toordinal() - 1).isoformat()
    year = re.fullmatch(r"\s*((?:19|20)\d{2})(?:年|年度)?\s*", text)
    return year.group(1) if year else None


def parse_number(value: Any) -> float | int | None:
    if value is None:
        return None
    text = unicodedata.normalize("NFKC", str(value)).strip()
    if text in {"", "-", "--", "—", "/", "不适用"}:
        return None
    negative = (
        bool(re.fullmatch(r"[（(].*[）)]", text))
        or bool(re.match(r"^[\-−—]", text))
        or bool(re.search(r"[\-−—]$", text))
    )
    text = text.strip("（）()")
    text = re.sub(r"^[\-−—]\s*", "", text)
    text = re.sub(r"\s*[\-−—]$", "", text)
    text = re.sub(r"[\s,，￥¥]", "", text)
    text = re.sub(r"(?:人民币)?(?:万|千|百)?元|%$", "", text)
    try:
        number = Decimal(text)
    except InvalidOperation:
        return None
    if negative:
        number = -abs(number)
    if number == number.to_integral_value():
        return int(number)
    return float(number)


def unit_multiplier(unit: str | None) -> int:
    normalized = normalize_label(unit or "元")
    if "万元" in normalized:
        return 10_000
    if "千元" in normalized:
        return 1_000
    if "百万元" in normalized:
        return 1_000_000
    return 1


def _detect_unit(cells: list[dict[str, Any]], explicit: str | None) -> str:
    if explicit:
        return explicit
    for cell in cells:
        text = normalize_label(str(cell.get("text", "")))
        match = re.search(r"单位(百万元|万元|千元|元)", text)
        if match:
            return match.group(1)
    return "元"


def _scaled(value: float | int, multiplier: int) -> float | int:
    result = Decimal(str(value)) * multiplier
    if result == result.to_integral_value():
        return int(result)
    return float(result)


def match_financial_table(
    cells: list[dict[str, Any]],
    aliases: dict[str, list[str]],
    unit: str | None = None,
    period_aliases: dict[str, str] | None = None,
) -> dict[str, Any]:
    """按字段语义和期间匹配二维 OCR 表格，不依赖固定单元格坐标。"""

    matrix = {(int(cell["row"]), int(cell["column"])): cell for cell in cells}
    alias_index = {
        normalize_label(alias): field_key
        for field_key, names in aliases.items()
        for alias in names
    }
    normalized_period_aliases = {
        normalize_label(key): value for key, value in (period_aliases or {}).items()
    }
    labels = []
    periods = []
    for cell in cells:
        label = normalize_label(str(cell.get("text", "")))
        if label in alias_index:
            labels.append((cell, alias_index[label]))
        period = normalized_period_aliases.get(label) or parse_period(str(cell.get("text", "")))
        if period:
            periods.append((cell, period))

    multiplier = unit_multiplier(_detect_unit(cells, unit))
    candidates: dict[tuple[str, str], list[float | int]] = defaultdict(list)
    for label_cell, field_key in labels:
        label_row, label_column = int(label_cell["row"]), int(label_cell["column"])
        for period_cell, period in periods:
            period_row, period_column = int(period_cell["row"]), int(period_cell["column"])
            coordinates = []
            if period_column != label_column:
                coordinates.append((label_row, period_column))
            if period_row != label_row:
                coordinates.append((period_row, label_column))
            for coordinate in dict.fromkeys(coordinates):
                value_cell = matrix.get(coordinate)
                if value_cell is None or value_cell in (label_cell, period_cell):
                    continue
                value = parse_number(value_cell.get("text"))
                if value is not None:
                    candidates[(field_key, period)].append(_scaled(value, multiplier))

    result: dict[str, Any] = {}
    conflicts = []
    for (field_key, period), values in sorted(candidates.items()):
        unique = list(dict.fromkeys(values))
        if len(unique) == 1:
            result.setdefault(field_key, {})[period] = unique[0]
        else:
            conflicts.append({"field_key": field_key, "period": period, "values": unique})
    if conflicts:
        result["_conflicts"] = conflicts
    return result
