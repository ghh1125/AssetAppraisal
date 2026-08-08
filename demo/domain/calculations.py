from __future__ import annotations

from datetime import date, timedelta
from decimal import Decimal
import re
from typing import Any


def _decimal(value: str | int | float | Decimal) -> Decimal:
    return Decimal(str(value))


def chinese_upper_integer(value: int) -> str:
    """Convert a non-negative integer to Chinese financial uppercase text."""
    if value < 0:
        raise ValueError("人民币大写金额不能为负数")
    if value == 0:
        return "零"
    digits = "零壹贰叁肆伍陆柒捌玖"
    inner_units = ("", "拾", "佰", "仟")
    group_units = ("", "万", "亿", "兆")

    def group_text(group: int) -> str:
        result = ""
        pending_zero = False
        for position in range(3, -1, -1):
            digit = group // (10**position) % 10
            if digit == 0:
                pending_zero = bool(result)
            else:
                if pending_zero:
                    result += "零"
                result += digits[digit] + inner_units[position]
                pending_zero = False
        return result

    groups: list[int] = []
    while value:
        groups.append(value % 10000)
        value //= 10000
    result = ""
    need_zero = False
    for index in range(len(groups) - 1, -1, -1):
        group = groups[index]
        if group == 0:
            need_zero = bool(result)
            continue
        if result and (need_zero or group < 1000):
            result += "零"
        result += group_text(group) + group_units[index]
        need_zero = False
    return result


def split_date(value: str) -> dict[str, int]:
    parsed = date.fromisoformat(value)
    return {"year": parsed.year, "month": parsed.month, "day": parsed.day}


def flexible_date_parts(value: Any) -> dict[str, int] | None:
    """Read a date from ISO, Chinese-formatted, or Excel-style display text."""
    if isinstance(value, date):
        parsed = value
    else:
        match = re.search(r"(?<!\d)(\d{4})\D+(\d{1,2})\D+(\d{1,2})(?!\d)", str(value))
        if not match:
            return None
        try:
            parsed = date(*(int(part) for part in match.groups()))
        except ValueError:
            return None
    return {"year": parsed.year, "month": parsed.month, "day": parsed.day}


def derive_system_fields(
    fields: dict[str, Any],
    report_date: str | date,
    *,
    final_value_field: str | None = None,
) -> dict[str, Any]:
    """Fill deterministic report-date and validity fields without infrastructure access."""
    result = dict(fields)
    report_parts = flexible_date_parts(report_date)
    if report_parts is None:
        raise ValueError(f"无法识别报告日期：{report_date}")

    def put_default(key: str, value: Any) -> None:
        current = result.get(key)
        if current in (None, "", []) or str(current).startswith("【待人工补充："):
            result[key] = value

    for part in ("year", "month", "day"):
        put_default(f"report_date_{part}", report_parts[part])
    put_default("report_number_year", report_parts["year"])

    valuation_parts = None
    for key in ("valuation_date_year", "valuation_date_month", "valuation_date_day"):
        valuation_parts = flexible_date_parts(result.get(key))
        if valuation_parts:
            break
    if valuation_parts:
        start = date(valuation_parts["year"], valuation_parts["month"], valuation_parts["day"])
        end_text = validity_period(start.isoformat())["end"]
        end_parts = split_date(end_text)
        for part in ("year", "month", "day"):
            result[f"valuation_date_{part}"] = valuation_parts[part]
            put_default(f"validity_start_{part}", valuation_parts[part])
            put_default(f"validity_end_{part}", end_parts[part])

    book = result.get("book_net_assets")
    if book not in (None, "", []) and _decimal(book) != 0:
        for prefix in ("income", "asset"):
            value = result.get(f"{prefix}_approach_value")
            if value in (None, "", []):
                continue
            difference = _decimal(value) - _decimal(book)
            put_default(f"{prefix}_increment", float(difference))
            put_default(f"{prefix}_increment_rate", float(difference / _decimal(book) * 100))

    if final_value_field and result.get(final_value_field) not in (None, "", []):
        put_default("final_appraisal_value", result[final_value_field])
        final_value = _decimal(result[final_value_field])
        final_value_yuan = final_value * Decimal("10000")
        # Assessment conclusions are stored in 万元, while one template
        # location asks for a Chinese RMB amount in 元.  A decimal 万元 value
        # can still be an exact integer yuan amount (e.g. 10023.56 万元).
        if final_value_yuan == final_value_yuan.to_integral():
            put_default(
                "final_value_chinese",
                chinese_upper_integer(int(final_value_yuan)),
            )
        if final_value == final_value.to_integral():
            put_default("final_value_chinese_wan", chinese_upper_integer(int(final_value)))
        elif final_value_yuan == final_value_yuan.to_integral():
            # The template labels this location as 万元, but a decimal 万元
            # conclusion is an exact integer amount in 元.  Mark the unit in
            # the value so the Word adapter can switch the nearby suffix to
            # 元整 instead of producing the misleading 万元整.
            put_default(
                "final_value_chinese_wan",
                f"{chinese_upper_integer(int(final_value_yuan))}元",
            )
        if book not in (None, "", []):
            difference = _decimal(result[final_value_field]) - _decimal(book)
            put_default("appraisal_increment", float(difference))
            if _decimal(book) != 0:
                put_default("appraisal_increment_rate", float(difference / _decimal(book) * 100))
    return result


def convert_amount(value: str, from_unit: str, to_unit: str) -> str:
    amount = _decimal(value)
    factors = {"元": Decimal("1"), "万元": Decimal("10000")}
    return format(amount * factors[from_unit] / factors[to_unit], "f")


def increment(appraised: str, book: str) -> str:
    return format(_decimal(appraised) - _decimal(book), "f")


def increment_rate(appraised: str, book: str) -> str | None:
    base = _decimal(book)
    return None if base == 0 else format((_decimal(appraised) - base) / base, "f")


def validity_period(value: str) -> dict[str, str]:
    start = date.fromisoformat(value)
    try:
        end = start.replace(year=start.year + 1) - timedelta(days=1)
    except ValueError:
        end = start.replace(year=start.year + 1, day=28) - timedelta(days=1)
    return {"start": start.isoformat(), "end": end.isoformat()}


def normalize_company_name(value: str) -> str:
    suffix = "有限公司"
    while value.endswith(suffix + suffix):
        value = value[: -len(suffix)]
    return value.strip()


def format_methods(values: list[str]) -> str:
    return "、".join(dict.fromkeys(values))
