from __future__ import annotations

from decimal import Decimal, InvalidOperation
from typing import Any

from demo.domain.financial_matching import normalize_label, parse_number


def _amount(value: Any) -> str:
    if value in (None, "", "-", "—", "="):
        return "0.00"
    text = str(value).replace(",", "").replace("，", "").replace(" ", "")
    try:
        return f"{Decimal(text):,.2f}"
    except InvalidOperation:
        return str(value).strip()


def _cell_index(normalized: dict[str, list[dict[str, Any]]]) -> dict[tuple[int, str, int, int], str]:
    return {
        (int(cell["page_number"]), str(cell.get("table_id", "")), int(cell["row"]), int(cell["column"])): str(cell.get("text", ""))
        for cell in normalized.get("table_cells", [])
    }


def _table_groups(normalized: dict[str, list[dict[str, Any]]]) -> list[tuple[int, str, list[dict[str, Any]]]]:
    grouped: dict[tuple[int, str], list[dict[str, Any]]] = {}
    for cell in normalized.get("table_cells", []):
        key = (int(cell.get("page_number", 0)), str(cell.get("table_id", "")))
        grouped.setdefault(key, []).append(cell)
    return [(page, table_id, cells) for (page, table_id), cells in grouped.items()]


def _table_context(normalized: dict[str, list[dict[str, Any]]]) -> dict[int, str]:
    context: dict[int, list[str]] = {}
    for block in normalized.get("text_blocks", []):
        text = str(block.get("text") or "").strip()
        if text:
            context.setdefault(int(block.get("page_number", 0)), []).append(text)
    return {page: " ".join(values) for page, values in context.items()}


def _find_semantic_table(
    normalized: dict[str, list[dict[str, Any]]], locator: dict[str, Any]
) -> list[dict[str, Any]]:
    table_markers = [normalize_label(value) for value in locator.get("table_markers", []) if value]
    page_markers = [normalize_label(value) for value in locator.get("page_markers", []) if value]
    row_aliases = [normalize_label(value) for value in locator.get("row_aliases", []) if value]
    column_aliases = [normalize_label(value) for value in locator.get("column_aliases", []) if value]
    context = _table_context(normalized)
    candidates = []
    for page, table_id, cells in _table_groups(normalized):
        max_columns = locator.get("max_columns")
        if max_columns is not None:
            column_count = len({int(cell.get("column", 0)) for cell in cells})
            if column_count > int(max_columns):
                continue
        table_text = normalize_label(" ".join(str(cell.get("text") or "") for cell in cells))
        page_text = normalize_label(context.get(page, ""))
        rows: dict[int, str] = {}
        for cell in cells:
            row_number = int(cell.get("row", 0))
            rows[row_number] = rows.get(row_number, "") + str(cell.get("text") or "")
        normalized_rows = [normalize_label(value) for value in rows.values()]
        row_matches = not row_aliases or any(
            alias in row_text for alias in row_aliases for row_text in normalized_rows
        )
        column_matches = not column_aliases or any(
            alias in normalize_label(str(cell.get("text") or ""))
            for alias in column_aliases
            for cell in cells
        )
        if not row_matches or not column_matches:
            continue
        score = 4 * sum(marker in table_text for marker in table_markers)
        score += 2 * sum(marker in page_text for marker in page_markers)
        score += 4 if row_aliases else 0
        score += 3 if column_aliases else 0
        if score:
            candidates.append((score, page, table_id, cells))
    if not candidates:
        return []
    candidates.sort(key=lambda item: (item[0], item[1]), reverse=True)
    return candidates[0][3]


def find_ocr_table(
    normalized: dict[str, list[dict[str, Any]]],
    *,
    table_id: str = "",
    table_markers: list[str] | None = None,
    page_markers: list[str] | None = None,
) -> list[dict[str, Any]]:
    """Find an OCR table by stable semantic markers, with optional legacy ID."""
    if table_id:
        cells = [
            cell for cell in normalized.get("table_cells", [])
            if str(cell.get("table_id", "")) == table_id
        ]
        if cells:
            return cells
    return _find_semantic_table(
        normalized,
        {"table_markers": table_markers or [], "page_markers": page_markers or []},
    )


def _semantic_value(normalized: dict[str, list[dict[str, Any]]], locator: dict[str, Any]) -> str:
    cells = _find_semantic_table(normalized, locator)
    if not cells:
        return ""
    rows: dict[int, list[dict[str, Any]]] = {}
    for cell in cells:
        rows.setdefault(int(cell["row"]), []).append(cell)
    row_aliases = [normalize_label(value) for value in locator.get("row_aliases", []) if value]
    target_rows = [
        row
        for row in rows.values()
        if any(alias in normalize_label("".join(str(cell.get("text") or "") for cell in row)) for alias in row_aliases)
    ]
    if not target_rows:
        return ""
    target = max(target_rows, key=lambda row: int(row[0].get("row", 0)))
    column_aliases = [normalize_label(value) for value in locator.get("column_aliases", []) if value]
    target_columns: list[int] = []
    if column_aliases:
        header_cells = [
            cell
            for row_number in sorted(rows)[:3]
            for cell in rows[row_number]
        ]
        for alias in column_aliases:
            exact = [
                int(cell["column"])
                for cell in header_cells
                if normalize_label(str(cell.get("text") or "")) == alias
            ]
            partial = [
                int(cell["column"])
                for cell in header_cells
                if alias in normalize_label(str(cell.get("text") or ""))
            ]
            for column in [*exact, *partial]:
                if column not in target_columns:
                    target_columns.append(column)
    values = [cell for cell in target if parse_number(cell.get("text")) is not None]
    if column_aliases:
        for target_column in target_columns:
            narrowed = [cell for cell in values if int(cell["column"]) == target_column]
            if narrowed:
                values = narrowed
                break
        else:
            return ""
    if not values:
        return ""
    selected = values[0] if locator.get("prefer_first_numeric") else values[-1]
    return _amount(selected.get("text")) if locator.get("format") == "amount" else str(selected.get("text") or "").strip()


def resolve_semantic_locator(
    normalized: dict[str, list[dict[str, Any]]], locator: dict[str, Any]
) -> str:
    """Resolve one layout-independent OCR field for a configured consumer."""
    return _semantic_value(normalized, locator)


def resolve_ocr_aux_fields(
    normalized: dict[str, list[dict[str, Any]]], config: dict[str, Any]
) -> dict[str, str]:
    """Resolve non-yellow helper fields used to populate fixed Word tables."""
    result: dict[str, str] = {}
    for field_key, locator in config.get("ocr_aux_fields", {}).items():
        if isinstance(locator, dict) and locator.get("kind") == "semantic":
            value = _semantic_value(normalized, locator)
            if value not in (None, ""):
                result[str(field_key)] = value
    return result


def resolve_configured_ocr_fields(
    normalized: dict[str, list[dict[str, Any]]], config: dict[str, Any]
) -> tuple[dict[str, Any], list[str]]:
    """Resolve fields whose values are explicitly anchored to PDF OCR table cells.

    Coordinates are project configuration, not business logic: a new report layout can
    supply a new page/table/row/column rule without changing the reusable pipeline.
    """
    index = _cell_index(normalized)
    values: dict[str, Any] = {}
    issues: list[str] = []
    for rule in config.get("ocr_field_rules", []):
        field_key = str(rule["field_key"])
        inputs = rule.get("inputs", {})
        rendered: dict[str, str] = {}
        missing = []
        for name, locator in inputs.items():
            if locator.get("kind") == "semantic":
                raw = _semantic_value(normalized, locator)
                if raw in (None, ""):
                    missing.append(f"{name}@semantic")
                rendered[name] = raw
                continue
            key = (
                int(locator["page"]),
                str(locator["table_id"]),
                int(locator["row"]),
                int(locator["column"]),
            )
            raw = index.get(key, "")
            if raw in (None, ""):
                missing.append(f"{name}@{key}")
            rendered[name] = _amount(raw) if locator.get("format") == "amount" else str(raw).strip()
        if missing:
            issues.append(f"{field_key}：PDF OCR 表格单元格缺失：{'、'.join(missing)}")
            continue
        try:
            values[field_key] = str(rule["template"]).format_map(rendered)
        except KeyError as exc:
            issues.append(f"{field_key}：OCR 字段模板缺少输入：{exc.args[0]}")
    return values, issues
