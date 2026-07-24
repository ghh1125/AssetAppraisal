from __future__ import annotations

from decimal import Decimal, InvalidOperation
from pathlib import Path
import re
from typing import Any

from .document import read_paragraph_containing, read_table_cell
from .excel import read_cells, read_range_values


def _format_value(value: Any, style: str | None, scale: Any = 1) -> str:
    if value in (None, ""):
        return ""
    if style == "year":
        match = re.search(r"(?:19|20)\d{2}", str(value))
        return match.group(0) if match else ""
    if style == "percent":
        try:
            number = Decimal(str(value).replace("%", ""))
            if "%" not in str(value) and abs(number) <= 1:
                number *= 100
            return f"{number.normalize()}%"
        except InvalidOperation:
            return str(value).strip()
    if style == "amount":
        try:
            return f"{Decimal(str(value)) * Decimal(str(scale)):,.2f}"
        except InvalidOperation:
            return str(value)
    return str(value).strip()


def resolve_material_field(
    spec: dict[str, Any],
    sources: dict[str, Path],
    source_lineage: dict[str, dict[str, str]] | None = None,
) -> tuple[str, dict[str, str]]:
    """Resolve one config-driven narrative field from Excel or a reference DOCX."""
    values: dict[str, str] = {}
    files: list[str] = []
    locators: list[str] = []
    kinds: list[str] = []
    source_lineage = source_lineage or {}
    for name, item in spec.get("inputs", {}).items():
        source_name = item["source"]
        path = sources[source_name]
        kind = item.get("kind", "excel_cell")
        if kind == "excel_cell":
            locator = item["locator"]
            raw = read_cells(path, [locator])[locator]
            value = _format_value(raw, item.get("format"), item.get("scale", 1))
        elif kind == "excel_range_unique":
            locator = item["locator"]
            unique: list[str] = []
            for raw in read_range_values(path, locator):
                value = _format_value(raw, item.get("format"), item.get("scale", 1))
                if value and value not in unique:
                    unique.append(value)
            value = item.get("joiner", "、").join(unique)
        elif kind == "excel_range_zero_status":
            locator = item["locator"]
            raw_values = [value for value in read_range_values(path, locator) if value not in (None, "")]
            all_zero = bool(raw_values) and all(Decimal(str(value)) == 0 for value in raw_values)
            value = item["zero_text"] if all_zero else item.get("nonzero_text", "")
        elif kind == "excel_cells_summary":
            parts: list[str] = []
            for label, cell_spec in item.get("cells", {}).items():
                raw = read_cells(path, [cell_spec["locator"]])[cell_spec["locator"]]
                formatted = _format_value(raw, cell_spec.get("format", "amount"), cell_spec.get("scale", 1))
                if formatted:
                    parts.append(f"{label}{formatted}元")
            value = item.get("joiner", "；").join(parts)
            locator = "、".join(str(cell_spec["locator"]) for cell_spec in item.get("cells", {}).values())
        elif kind == "document_paragraph":
            locator = f"段落包含：{item['contains']}"
            value = read_paragraph_containing(path, item["contains"])
        elif kind == "document_presence_status":
            locator = f"段落包含：{item['contains']}"
            found = read_paragraph_containing(path, item["contains"])
            value = item["found_text"] if found else item.get("missing_text", "")
        elif kind == "document_table_cell":
            locator = f"表格包含：{item['table_contains']}；行包含：{item['row_contains']}；列{item['column']}"
            value = read_table_cell(
                path,
                table_contains=item["table_contains"],
                row_contains=item["row_contains"],
                column=int(item["column"]),
            )
        elif kind == "file_stem":
            locator = "文件名"
            value = path.stem
        else:
            raise ValueError(f"不支持的材料字段类型：{kind}")
        suffix = item.get("strip_suffix")
        if suffix and value.endswith(suffix):
            value = value[: -len(suffix)] + item.get("suffix_replacement", "。")
        values[name] = value
        files.append(path.name)
        locators.append(locator)
        lineage = source_lineage.get(source_name, {})
        origin_source = lineage.get("origin_source")
        if origin_source and origin_source in sources:
            files.append(sources[origin_source].name)
            kinds.append(lineage.get("kind", "ocr_xlsx"))
        else:
            kinds.append(source_name)
    if any(value == "" for value in values.values()):
        return "", {"kind": "blank", "file": "；".join(dict.fromkeys(files)), "locator": "；".join(locators)}
    unique_kinds = list(dict.fromkeys(kinds))
    evidence_kind = unique_kinds[0] if unique_kinds == ["ocr_xlsx"] else "material:" + "+".join(unique_kinds)
    return spec["template"].format(**values), {
        "kind": evidence_kind,
        "file": "；".join(dict.fromkeys(files)),
        "locator": "；".join(locators),
    }
