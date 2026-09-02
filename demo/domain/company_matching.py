"""Conservative company-name matching for mixed material packages.

The preferred signal is always an exact recognised legal name.  A source-file
name is only used as a fallback when OCR did not recover the full legal name,
and only when a distinctive prefix of at least four characters matches.
"""

from __future__ import annotations

import re
from typing import Any, Callable, Iterable


_LEGAL_SUFFIXES = ("股份有限公司", "有限责任公司", "有限公司", "股份公司", "公司")


def normalize_company_name(value: Any) -> str:
    """Normalize punctuation and equivalent legal suffix wording."""
    return re.sub(r"[^0-9A-Za-z\u4e00-\u9fff]", "", str(value or "")).replace(
        "有限责任公司", "有限公司"
    )


def company_name_core(value: Any) -> str:
    """Return a normalized legal name without the terminal company suffix."""
    normalized = normalize_company_name(value)
    for suffix in _LEGAL_SUFFIXES:
        if normalized.endswith(suffix):
            return normalized[: -len(suffix)]
    return normalized


def _source_prefix_score(target_company_name: Any, source_file: Any) -> int:
    """Score a filename abbreviation such as 上海明悦 for a full legal name."""
    target_core = company_name_core(target_company_name)
    source_key = normalize_company_name(source_file)
    if len(target_core) < 4 or not source_key:
        return 0
    for length in range(len(target_core), 3, -1):
        if target_core[:length] in source_key:
            return 5_000 + length
    return 0


def company_match_score(
    target_company_name: Any,
    candidate_company_name: Any,
    source_file: Any = "",
) -> int:
    """Return a conservative match score; zero means no reliable match."""
    target_key = normalize_company_name(target_company_name)
    candidate_key = normalize_company_name(candidate_company_name)
    if not target_key:
        return 0
    if candidate_key == target_key:
        return 10_000

    target_core = company_name_core(target_company_name)
    candidate_core = company_name_core(candidate_company_name)
    if target_core and candidate_core:
        if target_core == candidate_core:
            return 9_000
        shorter = min(len(target_core), len(candidate_core))
        if shorter >= 4 and (
            target_core.startswith(candidate_core) or candidate_core.startswith(target_core)
        ):
            return 8_000 + shorter

    return _source_prefix_score(target_company_name, source_file)


def matching_company_records(
    records: Iterable[dict[str, Any]],
    target_company_name: Any,
    *,
    name_getter: Callable[[dict[str, Any]], Any],
    source_getter: Callable[[dict[str, Any]], Any],
) -> list[dict[str, Any]]:
    """Return only records tied at the highest reliable subject-match score."""
    scored = [
        (
            company_match_score(target_company_name, name_getter(record), source_getter(record)),
            record,
        )
        for record in records
    ]
    best = max((score for score, _ in scored), default=0)
    if best <= 0:
        return []
    return [record for score, record in scored if score == best]
