from __future__ import annotations

import hashlib
import inspect
import json
import re
import tempfile
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

from openpyxl import load_workbook

from demo import schemas
from demo.adapters.json_io import write_json
from demo.adapters.document import read_narrative_evidence, read_table_matrix
from demo.adapters.excel import (
    read_cells,
    try_read_cells,
    try_read_configured_table,
)
from demo.adapters.ocr_workbook import export_ocr_workbook, normalized_from_ocr_workbook
from demo.adapters.workflow_trace import (
    WorkflowTraceRecorder,
    trace_candidates,
    trace_evidence,
    trace_mappings,
    trace_resolved_fields,
)
from demo.adapters.word import (
    annotate_traceable_content,
    annotate_source_conflicts,
    document_paragraph_texts,
    fill_template,
    highlight_unresolved_placeholders,
    extract_word_comments,
    inventory_template,
    replace_image_markers,
    replace_report_number_year,
    replace_transaction_type_literals,
)
from demo.domain.generation_issues import (
    apply_page_locations,
    issues_from_special_evidence,
    issues_from_word_findings,
    organize_generation_issues,
)
from demo.domain.comment_mapping import (
    align_locations_to_output_template,
    build_comment_aware_locations,
)
from demo.domain.field_validation import (
    apply_missing_field_policy,
    normalize_narrative_modules,
    normalize_valuation_methods,
    require_financial_fields,
)
from demo.domain.field_validation import validate_valuation_subject_type
from demo.domain.financial_matching import blank_configured_table
from demo.domain.mapping import validate_mapping
from demo.domain.narrative_policy import (
    NARRATIVE_MODULE_LABELS,
    SELECTABLE_LLM_TEMPLATE_FIELDS,
    compose_company_profile_narrative,
    select_narrative_fields,
    select_llm_candidates,
)
from demo.domain.ocr_normalization import normalize_ocr_pages
from demo.domain.pdf_ocr_fields import find_ocr_table, resolve_configured_ocr_fields, resolve_ocr_aux_fields
from demo.domain.replacement import build_replacements
from demo.domain.source_reconciliation import SourceCandidate, reconcile_field
from demo.domain.source_labels import human_source_locator as render_source_locator
from demo.domain.source_labels import pdf_field_locators
from demo.domain.template_pagination import map_location_pages
from demo.domain.traceability_comments import TRACE_TITLES
from demo.domain.workflow_contracts import validate_workflow_contract
from demo.domain.yellow_routing import (
    RouteKind,
    fields_for_route,
    load_yellow_routes,
    validate_yellow_routes,
)
from demo.run import run_project


def _semantic_source_kind(source: dict[str, Any], source_overrides: dict[str, Path | None] | None) -> str:
    """Name a workbook role without relying on a sheet coordinate.

    The semantic extractor retains the file and sheet/cell evidence.  This
    helper only translates the selected file back to its uploaded role so the
    reconciliation record can explain a PDF-versus-workbook difference.
    """
    source_file = str(source.get("file", ""))
    for role, path in (source_overrides or {}).items():
        if path is not None and path.name == source_file:
            return {
                "reporting_workbook": "asset_workbook",
                "income_workbook": "income_workbook",
                "audited_financials": "audit_workbook",
            }.get(role, role)
    return "semantic_workbook"


def _reconcile_pdf_authoritative_fields(
    *,
    fields: dict[str, Any],
    evidence: dict[str, dict[str, Any]],
    resolved_ocr: dict[str, Any],
    field_keys: set[str],
    pdf_name: str,
    ocr_source_name: str,
    ocr_locators: dict[str, str] | None,
    source_overrides: dict[str, Path | None] | None,
) -> tuple[dict[str, Any], dict[str, dict[str, Any]], list[str], list[dict[str, Any]]]:
    """Use PDF/OCR for PDF-mandated fields and only compare workbooks.

    A workbook candidate may be useful evidence, but it must never silently
    replace a field that the Word comment requires to come from the audit
    report.  If an audit PDF/OCR candidate is absent, the field is deliberately
    left unresolved for the yellow placeholder policy.
    """
    updated_fields = dict(fields)
    updated_evidence = {key: dict(value) for key, value in evidence.items() if isinstance(value, dict)}
    issues: list[str] = []
    audit_rows: list[dict[str, Any]] = []
    for field_key in sorted(field_keys):
        candidates: list[SourceCandidate] = []
        current_value = updated_fields.get(field_key)
        current_source = updated_evidence.get(field_key, {})
        if current_value not in (None, "", [], {}):
            candidates.append(
                SourceCandidate(
                    value=current_value,
                    source_kind=_semantic_source_kind(current_source, source_overrides),
                    source_file=str(current_source.get("file", "")),
                    source_locator=str(current_source.get("locator", "")),
                )
            )
        # Only values actually produced by the OCR resolver count as PDF
        # evidence.  Do not treat a later Excel fallback as if it were PDF.
        if field_key in resolved_ocr and resolved_ocr[field_key] not in (None, "", [], {}):
            candidates.append(
                SourceCandidate(
                    value=resolved_ocr[field_key],
                    source_kind="pdf_ocr",
                    source_file=pdf_name or ocr_source_name,
                    source_locator=(ocr_locators or {}).get(
                        field_key,
                        f"审计 PDF：{field_key}",
                    ),
                )
            )
        # The citation below audit-derived tables is itself an audit-PDF
        # field.  Its value is the submitted PDF filename, not a financial
        # number requiring OCR.  It must still stay blank when no PDF exists.
        if field_key in {"audit_report_name", "financial_data_source_name"} and pdf_name:
            candidates.append(
                SourceCandidate(
                    value=Path(pdf_name).stem,
                    source_kind="pdf_ocr",
                    source_file=pdf_name,
                    source_locator="审计报告文件名",
                )
            )
        reconciled = reconcile_field(
            field_key=field_key,
            candidates=candidates,
            source_priority=("pdf_ocr", "audit_workbook", "asset_workbook", "income_workbook", "semantic_workbook"),
            require_primary_source=True,
            allow_fallback_when_primary_missing=True,
            primary_source_supplied=bool(pdf_name),
        )
        audit_rows.append(
            {
                "field_key": field_key,
                "selected": {
                    "value": reconciled.value,
                    "source_kind": reconciled.source_kind,
                    "source_file": reconciled.source_file,
                    "source_locator": reconciled.source_locator,
                },
                "candidates": [item.__dict__ for item in reconciled.candidates],
                "discrepancies": [item.__dict__ for item in reconciled.discrepancies],
                "issues": reconciled.issues,
                "pdf_uploaded": bool(pdf_name),
            }
        )
        if reconciled.value in (None, "", [], {}):
            updated_fields[field_key] = ""
            updated_evidence[field_key] = {"kind": "missing", "file": "", "locator": "无对应PDF文件，无法获取审计数据"}
        else:
            updated_fields[field_key] = reconciled.value
            updated_evidence[field_key] = {
                "kind": (
                    "pdf_ocr_xlsx"
                    if reconciled.source_kind == "pdf_ocr"
                    else reconciled.source_kind
                ),
                "file": reconciled.source_file,
                "locator": reconciled.source_locator,
            }
        issues.extend(reconciled.issues)
        for difference in reconciled.discrepancies:
            issues.append(
                f"{field_key}：数据不一致；审计PDF {difference.selected_source_file}"
                f"（{difference.selected_source_locator}）={difference.selected_value}；"
                f"其他材料 {difference.other_source_file}（{difference.other_source_locator}）={difference.other_value}。已按审计PDF填入"
            )
    return updated_fields, updated_evidence, issues, audit_rows


_RECONCILE_NUMBER = re.compile(r"[-+]?\d[\d,]*(?:\.\d+)?%?")


def _word_table_indices(config: dict[str, Any]) -> dict[str, int]:
    """Map logical fields to the 1-based table indexes visible in Word XML."""
    result: dict[str, int] = {}
    specs = list(config.get("financial_tables", []))
    for key in ("asset_scope_summary_table", "long_term_assets_table"):
        item = config.get(key)
        if isinstance(item, dict):
            specs.append(item)
    for spec in specs:
        if not isinstance(spec, dict):
            continue
        field_key = str(spec.get("field_key", "")).strip()
        if not field_key and spec is config.get("long_term_assets_table"):
            field_key = "long_term_assets_table"
        try:
            # Project/table replacement config is zero based; OOXML inventory
            # and reviewer-facing table identity are one based.
            table_index = int(spec["target_table_index"]) + 1
        except (KeyError, TypeError, ValueError):
            continue
        if field_key:
            result[field_key] = table_index
    return result


def _word_field_anchor_hints(
    locations: list[dict[str, Any]],
) -> dict[str, list[dict[str, str]]]:
    """Build stable before/after text anchors for mapped Word placeholders."""
    result: dict[str, list[dict[str, str]]] = defaultdict(list)
    marker_pattern = re.compile(r"20XX|X{2,}", re.I)
    for location in locations:
        field_key = str(location.get("field_key", "")).strip()
        context = str(location.get("context", ""))
        if not field_key or not context:
            continue
        markers = list(marker_pattern.finditer(context))
        try:
            occurrence = int(location.get("occurrence_index", 1)) - 1
        except (TypeError, ValueError):
            occurrence = 0
        if occurrence < 0 or occurrence >= len(markers):
            continue
        marker = markers[occurrence]
        previous_end = markers[occurrence - 1].end() if occurrence > 0 else 0
        next_start = markers[occurrence + 1].start() if occurrence + 1 < len(markers) else len(context)
        prefix = context[previous_end:marker.start()].strip()
        suffix = context[marker.end():next_start].strip()
        if len(_compact_anchor(prefix)) < 2 and len(_compact_anchor(suffix)) < 2:
            continue
        anchor = {
            "location_id": str(location.get("location_id", "")),
            "prefix": prefix[-80:],
            "suffix": suffix[:80],
        }
        if anchor not in result[field_key]:
            result[field_key].append(anchor)
    return dict(result)


def _compact_anchor(value: Any) -> str:
    return re.sub(r"[\s：:，,、（）()]", "", str(value or ""))


def _attach_word_anchor_hints(
    annotations: list[dict[str, Any]],
    anchors: dict[str, list[dict[str, str]]],
) -> list[dict[str, Any]]:
    for item in annotations:
        field_anchors = anchors.get(str(item.get("field_key", "")), [])
        if field_anchors:
            item["word_anchor_hints"] = field_anchors
    return annotations


def _numeric_value(value: Any) -> str | None:
    """Return a display value only for a standalone amount/rate cell."""
    if isinstance(value, bool):
        return None
    text = str(value if value is not None else "").strip()
    if not text or not _RECONCILE_NUMBER.fullmatch(text):
        return None
    return text


def _matrix(value: Any) -> list[list[Any]] | None:
    if isinstance(value, dict) and isinstance(value.get("rows"), list):
        value = value["rows"]
    if not isinstance(value, list) or not value or not all(isinstance(row, list) for row in value):
        return None
    return value


def _table_value_conflicts(
    selected: Any,
    other: Any,
    *,
    field_name: str,
    word_table_index: int | None = None,
) -> list[dict[str, str]]:
    """Return only differing numeric cells from two similarly shaped tables.

    A row label is retained as a Word lookup hint.  It prevents a common value
    such as ``0.00`` in another table from being coloured red by mistake.
    """
    selected_rows, other_rows = _matrix(selected), _matrix(other)
    if selected_rows is None or other_rows is None:
        return []
    differences: list[dict[str, str]] = []
    for row_index in range(1, min(len(selected_rows), len(other_rows))):
        left_row, right_row = selected_rows[row_index], other_rows[row_index]
        row_label = str(left_row[0]).strip() if left_row else ""
        for column_index in range(1, min(len(left_row), len(right_row))):
            left = _numeric_value(left_row[column_index])
            right = _numeric_value(right_row[column_index])
            if left is None or right is None:
                continue
            if left.replace(",", "") == right.replace(",", ""):
                continue
            header = ""
            if selected_rows and column_index < len(selected_rows[0]):
                header = str(selected_rows[0][column_index]).strip()
            label = " / ".join(item for item in (field_name, row_label, header) if item)
            item = {
                "pdf_value": left,
                "excel_value": right,
                "field_name": label or field_name,
                "word_context_hint": row_label,
                "word_period_hint": header,
            }
            if word_table_index is not None:
                item["word_table_index"] = word_table_index
            differences.append(item)
    return differences


def _human_source_locator(source_file: str, locator: str, field_name: str) -> str:
    """Turn internal OCR field locators into reviewer-readable locations."""
    return render_source_locator(source_file, locator, field_name)


def _resolve_unresolved_pdf_page_locators(
    field_names: dict[str, str],
    locators: dict[str, str],
    normalized: dict[str, Any],
    llm_adapter: Any | None,
) -> tuple[dict[str, str], list[str]]:
    """Complete PDF page labels with a closed-world LLM locator when needed.

    Deterministic OCR/table evidence remains the primary path.  The optional
    model receives only numbered OCR pages and may return only an existing
    page number; it cannot supply values or alter source selection.
    """
    updated = dict(locators)
    unresolved = [
        key
        for key in sorted(field_names)
        if key in updated and not re.search(r"第\d+页", str(updated.get(key, "")))
    ]
    if not unresolved or llm_adapter is None or not hasattr(llm_adapter, "locate_pdf_pages"):
        return updated, []
    try:
        page_map, issues = llm_adapter.locate_pdf_pages(
            normalized,
            field_names,
            unresolved,
        )
    except Exception as exc:
        return updated, [f"PDF页码定位调用失败：{exc}"]
    for field_key, page in (page_map or {}).items():
        if field_key not in unresolved:
            continue
        try:
            page_number = int(page)
        except (TypeError, ValueError):
            continue
        if page_number > 0:
            updated[field_key] = f"审计 PDF 第{page_number}页：{field_names.get(field_key, field_key)}"
    return updated, list(issues or [])


def _reviewable_source_conflicts(
    reconciliation_rows: list[dict[str, Any]],
    field_names: dict[str, str],
    word_table_indices: dict[str, int] | None = None,
) -> list[dict[str, str]]:
    """Turn reconciliation records into precise red-font/Word-comment tasks."""
    findings: list[dict[str, str]] = []
    seen: set[tuple[str, str, str, str]] = set()
    for row in reconciliation_rows:
        field_key = str(row.get("field_key", ""))
        display_name = field_names.get(field_key, field_key)
        word_table_index = (word_table_indices or {}).get(field_key)
        for discrepancy in row.get("discrepancies", []):
            if not isinstance(discrepancy, dict):
                continue
            common = {
                "field_key": field_key,
                "pdf_file": str(discrepancy.get("selected_source_file", "")),
                "pdf_locator": _human_source_locator(
                    str(discrepancy.get("selected_source_file", "")),
                    str(discrepancy.get("selected_source_locator", "")),
                    display_name,
                ),
                "excel_file": str(discrepancy.get("other_source_file", "")),
                "excel_locator": _human_source_locator(
                    str(discrepancy.get("other_source_file", "")),
                    str(discrepancy.get("other_source_locator", "")),
                    display_name,
                ),
            }
            pdf_value = _numeric_value(discrepancy.get("selected_value"))
            excel_value = _numeric_value(discrepancy.get("other_value"))
            if pdf_value is not None and excel_value is not None:
                scalar_candidate = {
                    **common,
                    "field_name": display_name,
                    "pdf_value": pdf_value,
                    "excel_value": excel_value,
                    "word_context_hint": "",
                    "word_period_hint": "",
                    "word_paragraph_hint": display_name,
                }
                if word_table_index is not None:
                    scalar_candidate["word_table_index"] = word_table_index
                candidates = [scalar_candidate]
            else:
                candidates = [
                    {**common, **item}
                    for item in _table_value_conflicts(
                        discrepancy.get("selected_value"),
                        discrepancy.get("other_value"),
                        field_name=display_name,
                        word_table_index=word_table_index,
                    )
                ]
            for candidate in candidates:
                key = (
                    candidate["field_key"],
                    candidate["field_name"],
                    candidate["pdf_value"],
                    candidate["excel_value"],
                )
                if key not in seen:
                    findings.append(candidate)
                    seen.add(key)
    return findings


def _first_table_numeric_value(value: Any) -> dict[str, str] | None:
    """Select one anchored cell to note an Excel-only table source in Word."""
    rows = _matrix(value)
    if rows is None:
        return None
    for row_index in range(1, len(rows)):
        row = rows[row_index]
        row_label = str(row[0]).strip() if row else ""
        for column_index in range(1, len(row)):
            numeric = _numeric_value(row[column_index])
            if numeric is None:
                continue
            header = str(rows[0][column_index]).strip() if column_index < len(rows[0]) else ""
            return {
                "excel_value": numeric,
                "word_context_hint": row_label,
                "word_period_hint": header,
                "field_suffix": " / ".join(item for item in (row_label, header) if item),
            }
    return None


def _reviewable_source_fallbacks(
    reconciliation_rows: list[dict[str, Any]],
    field_names: dict[str, str],
    word_table_indices: dict[str, int] | None = None,
) -> list[dict[str, str]]:
    """Record one Word note whenever a PDF-required field falls back to Excel.

    This is deliberately a source-availability note, not a mismatch: there is
    no PDF value to compare and the adopted Excel number stays black.
    """
    findings: list[dict[str, str]] = []
    for row in reconciliation_rows:
        selected = row.get("selected", {})
        if not isinstance(selected, dict) or selected.get("source_kind") == "pdf_ocr":
            continue
        source_file = str(selected.get("source_file", ""))
        value = selected.get("value")
        if not source_file or value in (None, "", [], {}):
            continue
        field_key = str(row.get("field_key", ""))
        field_name = field_names.get(field_key, field_key)
        word_table_index = (word_table_indices or {}).get(field_key)
        numeric = _numeric_value(value)
        if numeric is not None:
            item = {
                    "review_kind": "excel_fallback",
                    "field_key": field_key,
                    "field_name": field_name,
                    "excel_value": numeric,
                    "excel_file": source_file,
                    "excel_locator": _human_source_locator(
                        source_file,
                        str(selected.get("source_locator", "")),
                        field_name,
                    ),
                    "pdf_uploaded": bool(row.get("pdf_uploaded", False)),
                    "word_context_hint": "",
                    "word_period_hint": "",
                    "word_paragraph_hint": field_name,
                }
            if word_table_index is not None:
                item["word_table_index"] = word_table_index
            findings.append(item)
            continue
        first_value = _first_table_numeric_value(value)
        if first_value is not None:
            item = {
                    "review_kind": "excel_fallback",
                    "field_key": field_key,
                    "field_name": " / ".join(
                        item for item in (field_name, first_value["field_suffix"]) if item
                    ),
                    "excel_value": first_value["excel_value"],
                    "excel_file": source_file,
                    "excel_locator": _human_source_locator(
                        source_file,
                        str(selected.get("source_locator", "")),
                        field_name,
                    ),
                    "pdf_uploaded": bool(row.get("pdf_uploaded", False)),
                    "word_context_hint": first_value["word_context_hint"],
                    "word_period_hint": first_value["word_period_hint"],
                }
            if word_table_index is not None:
                item["word_table_index"] = word_table_index
            findings.append(item)
    return findings


def _reviewable_llm_notes(
    reviews: list[dict[str, str]],
    reconciliation_rows: list[dict[str, Any]],
    field_names: dict[str, str],
    word_table_indices: dict[str, int] | None = None,
) -> list[dict[str, str]]:
    """Attach LLM caution notes without letting the LLM recolour or alter data."""
    selected_by_field = {
        str(row.get("field_key", "")): row.get("selected", {})
        for row in reconciliation_rows
        if isinstance(row.get("selected"), dict)
    }
    notes: list[dict[str, str]] = []
    for review in reviews:
        if review.get("status") not in {"needs_review", "conflict", "missing"}:
            continue
        field_key = str(review.get("field_key", ""))
        selected = selected_by_field.get(field_key, {})
        if not isinstance(selected, dict):
            continue
        value = selected.get("value")
        numeric = _numeric_value(value)
        field_name = field_names.get(field_key, field_key)
        word_context_hint = ""
        word_period_hint = ""
        word_paragraph_hint = field_name
        if numeric is None:
            first_value = _first_table_numeric_value(value)
            if first_value is None:
                continue
            numeric = first_value["excel_value"]
            word_context_hint = first_value["word_context_hint"]
            word_period_hint = first_value["word_period_hint"]
            word_paragraph_hint = ""
            field_name = " / ".join(
                item for item in (field_name, first_value["field_suffix"]) if item
            )
        note = {
                "review_kind": "llm_review",
                "field_key": field_key,
                "field_name": field_name,
                # The generic Word annotator uses this as the selected text;
                # it is intentionally not an Excel assertion.
                "excel_value": numeric,
                "excel_file": str(selected.get("source_file", "")),
                "excel_locator": _human_source_locator(
                    str(selected.get("source_file", "")),
                    str(selected.get("source_locator", "")),
                    field_name,
                ),
                "word_context_hint": word_context_hint,
                "word_period_hint": word_period_hint,
                "word_paragraph_hint": word_paragraph_hint,
                "review_status": str(review.get("status", "")),
                "review_reason": str(review.get("reason", "")),
                "llm_comment": str(review.get("comment", "")),
            }
        word_table_index = (word_table_indices or {}).get(field_key)
        if word_table_index is not None:
            note["word_table_index"] = word_table_index
        notes.append(note)
    return notes


def _attach_llm_comments(
    annotations: list[dict[str, Any]],
    llm_notes: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Attach the model's complete human-facing comment to source findings."""
    notes_by_field = {
        str(item.get("field_key", "")): item
        for item in llm_notes
        if item.get("llm_comment")
    }
    enriched: list[dict[str, Any]] = []
    for item in annotations:
        copy = dict(item)
        note = notes_by_field.get(str(item.get("field_key", "")))
        if note:
            copy["llm_comment"] = str(note.get("llm_comment", ""))
            copy["review_status"] = str(note.get("review_status", ""))
            copy["review_reason"] = str(note.get("review_reason", ""))
        enriched.append(copy)
    return enriched


_LLM_REVIEW_SOURCE_KINDS = frozenset(
    {
        "pdf_ocr_xlsx",
        "ocr_xlsx",
        "audit_workbook",
        "asset_workbook",
        "income_workbook",
        "semantic_workbook",
    }
)


def _llm_review_rows(
    fields: dict[str, Any],
    evidence: dict[str, dict[str, Any]],
    reconciliation_rows: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Include every resolved PDF/Excel field in the LLM review batch once."""
    rows = list(reconciliation_rows)
    present = {str(item.get("field_key", "")) for item in rows}
    for field_key, value in fields.items():
        if field_key in present or value in (None, "", [], {}):
            continue
        source = evidence.get(field_key, {})
        if not isinstance(source, dict) or source.get("kind") not in _LLM_REVIEW_SOURCE_KINDS:
            continue
        candidate = {
            "value": value,
            "source_kind": source.get("kind", ""),
            "source_file": source.get("file", ""),
            "source_locator": source.get("locator", ""),
        }
        rows.append(
            {
                "field_key": field_key,
                "selected": candidate,
                "candidates": [candidate],
                "discrepancies": [],
            }
        )
        present.add(field_key)
    return rows


def _keep_unresolved_ocr_issues(
    issues: list[str],
    fields: dict[str, Any],
) -> list[str]:
    """Keep OCR failures only when the same field is still unresolved.

    PDF is optional in the four-node workflow.  A semantic Excel extractor can
    resolve a field even when the legacy OCR resolver reports its configured
    cells as missing; that stale diagnostic must not look like a current data
    failure in the run result.
    """
    kept: list[str] = []
    for issue in issues:
        if "：PDF OCR 表格单元格缺失：" not in issue:
            kept.append(issue)
            continue
        field_key = issue.split("：", 1)[0]
        value = fields.get(field_key)
        if value in (None, "", [], {}):
            kept.append(issue)
    return kept


def _trace_status_by_field(
    evidence: dict[str, dict[str, Any]],
    llm_reviews: list[dict[str, Any]],
    conflicts: list[dict[str, Any]],
    fallbacks: list[dict[str, Any]],
) -> dict[str, str]:
    """Reduce source/review evidence to the four Word display states."""
    review_by_field = {
        str(item.get("field_key", "")): str(item.get("status", ""))
        for item in llm_reviews
        if item.get("field_key")
    }
    fallback_fields = {str(item.get("field_key", "")) for item in fallbacks}
    result: dict[str, str] = {}
    for field_key, source in evidence.items():
        source = source if isinstance(source, dict) else {}
        if source.get("kind") == "missing":
            result[field_key] = "missing"
        elif review_by_field.get(field_key) in {
            "needs_review",
            "conflict",
            "missing",
        }:
            result[field_key] = "review"
        elif field_key in fallback_fields:
            result[field_key] = "fallback"
        else:
            result[field_key] = "verified"
    return result


def _trace_comment(
    *,
    field_key: str,
    field_name: str,
    status: str,
    source: dict[str, Any] | None,
    llm_review: dict[str, Any] | None,
    model_name: str,
) -> str:
    """Create a concise reviewer-facing provenance note with a clear title."""
    source = source if isinstance(source, dict) else {}
    llm_review = llm_review if isinstance(llm_review, dict) else {}
    source_kind = str(source.get("kind", ""))
    source_file = str(source.get("file", ""))
    source_locator = render_source_locator(
        source_file,
        source.get("locator", ""),
        field_name,
    )
    if status == "missing":
        return (
            f"【未找到数据】{field_name}已检索本次上传的审计PDF、Excel、补充材料及已启用的企业信息接口，"
            "仍未找到可可靠采用的数据，因此保留黄色XXX。请人工补充或核对原始材料。"
        )
    if source_kind in {"node_input", "manual", "manual_input"}:
        source_text = "来源：人工基础信息；该值由用户输入并确认采用。"
    elif source_kind == "qichacha_api":
        source_text = "来源：企查查API；接口返回结果已通过查询主体名称匹配后采用。"
    elif source_kind.startswith("bailian_glm"):
        model = model_name or "百炼大模型"
        source_text = (
            f"来源：百炼模型 {model}；模型基于已解析材料和企业信息证据生成，"
            "并经用户选择后写入报告。"
        )
    elif source_kind in {"computed", "derived", "system_calculation"}:
        source_text = "来源：系统计算；由报告中已选定且可追溯的基础字段计算得到。"
    else:
        if source_locator.startswith(("《", "审计 PDF")):
            source_text = f"来源：{source_locator}。"
        else:
            file_prefix = f"《{source_file}》" if source_file else "本次上传材料"
            source_text = f"来源：{file_prefix}{source_locator}。"
    review_status = str(llm_review.get("status", ""))
    review_reason = str(llm_review.get("reason", "")).strip()
    if status == "verified":
        title = "【来源已核验】"
        if review_status == "accept":
            review_text = "LLM已核对科目、期间、单位和来源定位，复核结论为通过。"
        elif source_kind in {"node_input", "manual", "manual_input"}:
            review_text = "本字段属于人工输入，不对其真实性作自动推断。"
        elif source_kind == "qichacha_api":
            review_text = "系统已完成API主体匹配校验。"
        elif source_kind.startswith("bailian_glm"):
            review_text = "该内容属于模型生成叙述，不作为审计数值依据。"
        else:
            review_text = "系统语义规则已完成唯一匹配；未发现需提示的来源冲突。"
    elif status == "fallback":
        title = "【来源待补充核验】"
        review_text = "当前已采用Excel或其他可用材料，但尚未完成审计PDF对照，请人工复核口径。"
    else:
        title = "【数据冲突，需人工复核】"
        review_text = (
            f"LLM/规则复核结论：{review_reason or '来源间存在差异或口径歧义，请核对原始材料。'}"
        )
    return f"{title}{field_name}。{source_text}{review_text}"


def _paragraph_trace_annotations(
    locations: list[dict[str, Any]],
    replacements: dict[str, str],
    evidence: dict[str, dict[str, Any]],
    field_names: dict[str, str],
    status_by_field: dict[str, str],
    review_by_field: dict[str, dict[str, Any]],
    model_name: str,
) -> list[dict[str, Any]]:
    annotations: list[dict[str, Any]] = []
    for location in locations:
        location_id = str(location.get("location_id", ""))
        if location_id not in replacements:
            continue
        field_key = str(location.get("field_key", ""))
        value = str(replacements.get(location_id, "") or "").strip()
        if not value:
            continue
        source = evidence.get(field_key, {})
        missing = bool(re.search(r"20XX|X{2,}", value, re.I))
        status = "missing" if missing else status_by_field.get(field_key, "verified")
        field_name = (
            field_names.get(field_key)
            or str(location.get("field_name") or field_key or "该字段")
        )
        comment = _trace_comment(
            field_key=field_key,
            field_name=field_name,
            status=status,
            source=source,
            llm_review=review_by_field.get(field_key),
            model_name=model_name,
        )
        marker = str(location.get("marker") or "")
        context = str(location.get("context") or "")
        context_hint = (
            ""
            if location.get("record_type") == "黄色标注内容块"
            else context.replace(marker, "") if marker else ""
        )
        lines = [line.strip() for line in value.splitlines() if line.strip()] or [value]
        for line in lines:
            annotations.append(
                {
                    "field_key": field_key,
                    "field_name": field_name,
                    "target": line,
                    "part": str(location.get("part") or "word/document.xml"),
                    "paragraph_index_hint": int(location.get("paragraph_index") or 0),
                    "target_occurrence": int(location.get("occurrence_index") or 1),
                    "context_hint": context_hint if len(lines) == 1 else "",
                    "status": status,
                    "comment": comment,
                }
            )
    return annotations


def _table_trace_annotations(
    table_replacements: dict[int, list[list[str]]],
    table_fields: dict[int, tuple[str, int, set[int]]],
    evidence: dict[str, dict[str, Any]],
    field_names: dict[str, str],
    status_by_field: dict[str, str],
    review_by_field: dict[str, dict[str, Any]],
    model_name: str,
) -> list[dict[str, Any]]:
    annotations: list[dict[str, Any]] = []
    for table_index, matrix in table_replacements.items():
        field_key, first_data_row, data_columns = table_fields.get(
            table_index,
            (f"word_table_{table_index}", 0, set()),
        )
        if table_index not in table_fields:
            continue
        field_name = field_names.get(field_key, field_key)
        source = evidence.get(field_key, {})
        base_status = status_by_field.get(field_key, "verified")
        for row_index, row in enumerate(matrix):
            if row_index < first_data_row:
                continue
            columns = data_columns or set(range(len(row)))
            for column_index, value in enumerate(row):
                if column_index not in columns:
                    continue
                target = str(value or "").strip()
                if not target:
                    continue
                status = "missing" if re.search(r"20XX|X{2,}", target, re.I) else base_status
                cell_name = field_name
                if row and column_index > 0 and str(row[0]).strip():
                    cell_name = f"{field_name} / {str(row[0]).strip()}"
                comment = _trace_comment(
                    field_key=field_key,
                    field_name=cell_name,
                    status=status,
                    source=source,
                    llm_review=review_by_field.get(field_key),
                    model_name=model_name,
                )
                annotations.append(
                    {
                        "field_key": field_key,
                        "field_name": cell_name,
                        "target": target,
                        "table_index": table_index,
                        "row_index": row_index,
                        "column_index": column_index,
                        "status": status,
                        "comment": comment,
                    }
                )
    return annotations


@dataclass(frozen=True)
class PipelineResult:
    report_path: Path
    ocr_workbook_path: Path | None
    manifest_path: Path
    issues: list[str]
    candidate_path: Path | None = None
    candidate_fields: dict[str, Any] = field(default_factory=dict)


OcrFieldResolver = Callable[
    [dict[str, list[dict[str, Any]]], dict[str, Any]],
    tuple[dict[str, Any], list[str]],
]


def _path(base: Path, value: str | Path) -> Path:
    path = Path(value)
    return path if path.is_absolute() else (base / path).resolve()


def _sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _legacy_evidence(path: Path) -> dict[str, dict[str, str]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    sheet = workbook["填充结果"]
    header = {str(cell.value): index for index, cell in enumerate(sheet[1])}
    result: dict[str, dict[str, str]] = {}
    for row in sheet.iter_rows(min_row=2, values_only=True):
        field_key = str(row[header["标准字段"]])
        result.setdefault(
            field_key,
            {
                "kind": str(row[header["来源类别"]] or "legacy"),
                "file": str(row[header["来源文件"]] or ""),
                "locator": str(row[header["来源位置"]] or ""),
            },
        )
    return result


def _default_ocr_field_resolver(
    normalized: dict[str, list[dict[str, Any]]], config: dict[str, Any]
) -> tuple[dict[str, Any], list[str]]:
    values: dict[str, Any] = {}
    history_keys = {
        "historical_balance_sheet_table",
        "historical_income_statement_table",
    }
    history_priorities: dict[str, int] = {}
    for record in normalized.get("financial_data", []):
        field_key = record.get("field_key")
        if field_key and record.get("value") not in (None, "", []):
            field_key = str(field_key)
            if field_key not in history_keys:
                values[field_key] = record["value"]
                continue
            priority = _ocr_history_record_priority(record)
            if field_key not in values or priority > history_priorities.get(field_key, -1):
                values[field_key] = record["value"]
                history_priorities[field_key] = priority
    configured, issues = resolve_configured_ocr_fields(normalized, config)
    for field_key, value in configured.items():
        # Configured semantic rules are still allowed to fill ordinary OCR
        # fields, but they must not overwrite a formal financial statement
        # selected from the OCR workbook's standard-financial-data records.
        if field_key in history_keys and field_key in values:
            continue
        values[field_key] = value
    return values, issues


def _ocr_history_record_priority(record: dict[str, Any]) -> int:
    """Rank OCR history tables by the semantic source represented by them.

    A single OCR workbook can contain both the formal audit statement and a
    prepared ``历资表``/``历利表`` copied into a valuation workbook.  The
    formal statement is the authoritative source for historical book values;
    the prepared table remains a fallback when no formal statement exists.
    This uses semantic sheet/table labels, never project-specific cells.
    """
    evidence_id = str(record.get("evidence_id", ""))
    if any(marker in evidence_id for marker in ("历资表", "历利表")):
        return 10
    if any(marker in evidence_id for marker in ("资产负债表", "利润表")):
        return 30
    return 0


def _ocr_has_formal_history(
    normalized: dict[str, list[dict[str, Any]]],
    field_key: str,
) -> bool:
    return any(
        str(record.get("field_key", "")) == field_key
        and record.get("value") not in (None, "", [])
        and _ocr_history_record_priority(record) >= 30
        for record in normalized.get("financial_data", [])
    )


def _provider_fields(payload: Any) -> dict[str, Any]:
    if not isinstance(payload, dict):
        return {}
    nested = payload.get("fields")
    return nested if isinstance(nested, dict) else payload


def _normalized_company_name(value: Any) -> str:
    """Normalize company names just enough for a provider identity check."""
    text = re.sub(r"[\s（）()\[\]【】]", "", str(value or ""))
    return text.replace("有限责任公司", "有限公司")


def _validated_qcc_payload(
    payload: Any,
    requested_name: str,
    role: str,
    issues: list[str],
) -> dict[str, Any]:
    """Reject a QCC response for a different company instead of filling it."""
    if not isinstance(payload, dict):
        return {}
    profile = payload.get("profile")
    returned_name = profile.get("name") if isinstance(profile, dict) else ""
    if returned_name and requested_name:
        requested = _normalized_company_name(requested_name)
        returned = _normalized_company_name(returned_name)
        if requested != returned:
            issues.append(
                f"企查查身份核验失败（{role}）：查询“{requested_name}”返回“{returned_name}”，相关字段已留空"
            )
            return {}
    return payload


def _asset_method_label(selected: Any) -> str:
    """Return the method label for the template's second result section."""
    text = str(selected or "")
    if "资产基础法" in text:
        return "资产基础法"
    if "市场法" in text:
        return "市场法"
    if "收益法" in text:
        return "收益法"
    return ""


def _filter_provider(
    payload: Any,
    allowed: set[str],
    provider_name: str,
    issues: list[str],
) -> dict[str, Any]:
    result = {}
    for field_key, value in _provider_fields(payload).items():
        if field_key not in allowed:
            issues.append(f"{provider_name} 返回越权字段，已丢弃：{field_key}")
        elif value not in (None, "", []):
            result[field_key] = value
    return result


def _paragraph_replacements(config: dict[str, Any], fields: dict[str, Any]) -> dict[tuple[str, int], str]:
    result = {}
    string_fields = defaultdict(
        lambda: "XXX",
        {key: str(value) for key, value in fields.items()},
    )
    for spec in config.get("paragraph_replacements", []):
        if "field_key" in spec:
            value = str(fields.get(spec["field_key"], ""))
        elif "template" in spec:
            blank_if = spec.get("blank_if_empty")
            if blank_if and not str(fields.get(blank_if, "") or "").strip():
                value = ""
            else:
                value = spec["template"].format_map(string_fields)
        else:
            value = str(spec.get("value", ""))
        # Do not leave empty manual-input brackets in an otherwise valid
        # report when a project intentionally omits an optional short name.
        value = value.replace("（简称：）", "").replace("（以下简称：）", "")
        result[(spec["part"], int(spec["paragraph_index"]))] = value
    return result


def _source_path(
    base: Path,
    config: dict[str, Any],
    overrides: dict[str, Path | None] | None,
    source_name: str,
) -> Path | None:
    if overrides is not None and source_name in overrides:
        override = overrides[source_name]
        return Path(override).resolve() if override is not None else None
    configured = config.get("sources", {}).get(source_name)
    return _path(base, configured) if configured else None


def _company_profile_table(profile: dict[str, Any], fallback_name: Any = "", fallback_capital: Any = "") -> list[list[str]]:
    """Render the backend-provided company profile into the template's 2-cell table."""
    profile = profile if isinstance(profile, dict) else {}
    present = lambda value: str(value) if value not in (None, "") else "XXX"
    def present_date(value: Any) -> str:
        text = present(value)
        match = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})", text)
        if not match:
            return text
        year, month, day = (int(part) for part in match.groups())
        return f"{year:04d}年{month:02d}月{day:02d}日"

    credit_code = present(profile.get("credit_code"))
    name = present(profile.get("name") or fallback_name)
    # The workbook/PDF value represents the valuation-date financial state;
    # QCC's business profile represents the current registration state.  When
    # both exist, the valuation-date material must win in an appraisal report.
    capital = present(fallback_capital or profile.get("registered_capital"))
    return [
        [f"统一社会信用代码：{credit_code}", f"企业名称：{name}"],
        [f"类型：{present(profile.get('company_type'))}", f"法定代表人：{present(profile.get('legal_representative'))}"],
        [f"注册资本：{capital}", f"成立日期：{present_date(profile.get('establish_date'))}"],
        [f"营业期限自：{present_date(profile.get('term_start'))}", f"营业期限至：{present_date(profile.get('term_end'))}"],
        [f"登记机关：{present(profile.get('registration_authority'))}", f"核准日期：{present_date(profile.get('approval_date'))}"],
        [f"登记状态：{present(profile.get('status'))}"],
        [f"注册地址：{present(profile.get('address'))}"],
        [f"许可项目：{present(profile.get('business_scope'))}"],
    ]


def _blank_cross_source_table(rows: list[dict[str, Any]]) -> list[list[str]]:
    """Build an unresolved long-term-assets table without reading old cells."""
    matrix = [["项目", "账面金额（元）", "数量", "现状、特点"]]
    for row in rows:
        matrix.append([
            str(row["label"]),
            "XXX",
            str(row.get("quantity", "")),
            str(row.get("condition", "")),
        ])
    return matrix


def _apply_ocr_overrides_to_table(
    matrix: list[list[str]], spec: dict[str, Any], overrides: dict[str, str]
) -> list[list[str]]:
    """Replace configured table amounts with semantically matched PDF OCR values."""
    by_label = {
        str(row.get("label", "")).strip(): str(row.get("ocr_field_key"))
        for row in spec.get("rows", [])
        if row.get("ocr_field_key")
    }
    for row in matrix:
        if not row:
            continue
        key = by_label.get(str(row[0]).strip())
        if key and overrides.get(key) not in (None, "") and len(row) > 1:
            row[1] = str(overrides[key])
    return matrix


def _qcc_table_rows(payload: dict[str, Any], key: str, width: int) -> list[list[str]]:
    rows = payload.get(key, []) if isinstance(payload, dict) else []
    matrix = []
    for row in rows:
        if not isinstance(row, dict):
            continue
        if key == "trademark_rows":
            # The reviewed report uses the second column as an optional image
            # cell and the third column as the textual trademark name.  Do not
            # put an image URL (or the fallback word “图样”) into the name
            # column.  A graphical mark is represented explicitly as “图形”.
            name = row.get("name", "") or ("图形" if row.get("image") else "")
            values = [row.get("application_date", ""), "", name, row.get("registration_number", ""), row.get("class", ""), row.get("status", ""), row.get("announcement_date", "")]
        else:
            values = [row.get(k, "") for k in ("index", "name", "registration_number", "first_publication_date", "approval_date")]
        matrix.append([str(value or "") for value in values[:width]])
    return matrix


def _money_yuan(value: Any, *, plain_number_unit: str = "万") -> str:
    """Normalize QCC/reference capital values to the template's yuan unit."""
    if value in (None, ""):
        return ""
    text = str(value).strip().replace(",", "").replace("人民币", "")
    unit = 1
    if "万" in text or plain_number_unit == "万":
        unit = 10000
    text = text.replace("万元", "").replace("万", "").replace("元", "").strip()
    try:
        from decimal import Decimal

        return f"{Decimal(text) * unit:,.2f}"
    except Exception:
        return str(value).strip()


def _equity_matrix_from_partners(rows: list[dict[str, Any]]) -> list[list[str]]:
    matrix = [["序号", "股东名称", "总出资（元）", "股权比例"]]
    total = 0.0
    for index, row in enumerate(rows, 1):
        name = str(row.get("name", ""))
        capital = _money_yuan(row.get("capital", ""))
        percent = str(row.get("percent", ""))
        if not (name or capital or percent):
            continue
        matrix.append([str(index), name, capital, percent])
        try:
            total += float(capital.replace(",", ""))
        except (ValueError, AttributeError):
            pass
    if len(matrix) > 1:
        # Keep one empty input row before the total row.  This is part of the
        # supplied report layout (and is retained in the reviewed reference),
        # so a one-shareholder result does not collapse the table vertically.
        if len(matrix) == 2:
            matrix.append(["", "", "", ""])
        matrix.append(["合计", "合计", f"{total:,.2f}" if total else "", "100%"])
    return matrix


def _valuation_date_ownership_matrix(
    audited_matrix: list[list[str]] | None,
    current_partner_rows: list[dict[str, Any]],
) -> list[list[str]]:
    """Prefer dated audit evidence over a current company-registry snapshot."""
    if audited_matrix and len(audited_matrix) > 1:
        return audited_matrix
    if current_partner_rows:
        return _equity_matrix_from_partners(current_partner_rows)
    return [
        ["序号", "股东名称", "总出资（元）", "股权比例"],
        ["XXX", "XXX", "XXX", "XXX"],
    ]


def _ownership_matrix_summary(matrix: list[list[str]] | None) -> str:
    """Render audited ownership rows as a compact valuation-date fact."""
    if not matrix or len(matrix) < 2:
        return ""
    parts = []
    for row in matrix[1:]:
        values = [str(value).strip() for value in row]
        if not values or values[0] in {"", "合计"} or len(values) < 4:
            continue
        name, capital, percent = values[1], values[2], values[3]
        if name:
            parts.append(f"{name}：出资额{capital}元，股权比例{percent}")
    return "；".join(parts)


def _ocr_ownership_matrix(
    normalized: dict[str, list[dict[str, Any]]], spec: dict[str, Any]
) -> tuple[list[list[str]], list[str]]:
    """Build the founding-ownership table from configured PDF-OCR cells.

    Audit reports often split a shareholder name over several OCR rows and
    may place the amount in a neighbouring row.  The project configuration
    therefore declares the exact cells to join instead of relying on a
    positional guess.  This keeps the rule reusable for another report
    layout while preserving the PDF/XLSX source boundary.
    """
    matched_cells = find_ocr_table(
        normalized,
        table_id=str(spec.get("table_id", "")),
        table_markers=list(spec.get("table_markers", [])),
        page_markers=list(spec.get("page_markers", [])),
    )
    cells = {
        (int(cell.get("row", 0)), int(cell.get("column", 0))): str(cell.get("text") or "").strip()
        for cell in matched_cells
    }

    def cell(ref: Any) -> str:
        if not isinstance(ref, (list, tuple)) or len(ref) != 2:
            return ""
        try:
            return cells.get((int(ref[0]), int(ref[1])), "")
        except (TypeError, ValueError):
            return ""

    matrix = [list(spec.get("header", ["序号", "股东名称", "总出资（元）", "股权比例"]))]
    if not spec.get("rows"):
        return _auto_ocr_ownership_matrix(matched_cells, matrix)
    names: list[str] = []
    total = 0.0
    for index, row in enumerate(spec.get("rows", []), 1):
        name = "".join(cell(ref) for ref in row.get("name_cells", []))
        capital = cell(row.get("capital_cell"))
        percent = str(row.get("percent", ""))
        if not (name or capital or percent):
            continue
        names.append(name)
        matrix.append([str(index), name, capital, percent])
        try:
            total += float(capital.replace(",", ""))
        except (TypeError, ValueError):
            pass
    if len(matrix) > 1 and spec.get("include_total", True):
        if len(matrix) == 2 and spec.get("preserve_blank_row", True):
            matrix.append(["", "", "", ""])
        matrix.append(["合计", "合计", f"{total:,.2f}" if total else "", "100%"])
    return matrix, names


def _auto_ocr_ownership_matrix(
    cells: list[dict[str, Any]], matrix: list[list[str]]
) -> tuple[list[list[str]], list[str]]:
    """Infer shareholder rows when a PDF changes page/table coordinates."""
    from demo.domain.financial_matching import parse_number

    by_row: dict[int, list[dict[str, Any]]] = {}
    for cell in cells:
        by_row.setdefault(int(cell.get("row", 0)), []).append(cell)
    groups: list[list[dict[str, Any]]] = []
    current: list[dict[str, Any]] = []
    for row_number in sorted(by_row):
        row = sorted(by_row[row_number], key=lambda item: int(item.get("column", 0)))
        first = "".join(str(cell.get("text") or "").strip() for cell in row if int(cell.get("column", 0)) == 1)
        if not current and (not first or first in {"项目", "出资方", "合计"}):
            continue
        if current or first:
            current.extend(row)
        if current and ("有限公司" in first or "公司" in first or "企业" in first):
            groups.append(current)
            current = []
    result_names: list[str] = []
    total = 0.0
    parsed: list[tuple[str, str]] = []
    for group in groups:
        name = "".join(
            str(cell.get("text") or "").strip()
            for cell in sorted(group, key=lambda item: (int(item.get("row", 0)), int(item.get("column", 0))))
            if int(cell.get("column", 0)) == 1 and parse_number(cell.get("text")) is None
        )
        amounts = [
            parse_number(cell.get("text"))
            for cell in group
            if parse_number(cell.get("text")) is not None
        ]
        if not name or not amounts:
            continue
        amount = float(amounts[0])
        parsed.append((name, f"{amount:,.2f}"))
        total += amount
    for index, (name, capital) in enumerate(parsed, 1):
        percent = f"{float(capital.replace(',', '')) / total:.0%}" if total else ""
        result_names.append(name)
        matrix.append([str(index), name, capital, percent])
    if parsed:
        matrix.append(["合计", "合计", f"{total:,.2f}", "100%"])
    return matrix, result_names


def _equity_matrix_from_reference(base: Path, config: dict[str, Any], table_index: int) -> list[list[str]]:
    source_path = _path(base, config["sources"]["reference_report"])
    source = read_table_matrix(source_path, int(table_index))
    matrix = [["序号", "股东名称", "总出资（元）", "股权比例"]]
    for row in source[1:]:
        values = [str(value or "").strip() for value in row]
        if len(values) < 4 or not any(values):
            continue
        matrix.append([values[0], values[1], _money_yuan(values[2]), values[3]])
    return matrix


def _profile_matrix_from_reference(base: Path, config: dict[str, Any], table_index: int) -> list[list[str]]:
    source_path = _path(base, config["sources"]["reference_report"])
    source = read_table_matrix(source_path, int(table_index))
    rows = [[str(value or "").strip() for value in row] for row in source if row]
    values: dict[str, str] = {}
    for row in rows:
        if not row:
            continue
        label = row[0]
        if label == "统一社会信用代码":
            values["credit_code"] = row[1] if len(row) > 1 else ""
            values["name"] = row[3] if len(row) > 3 else ""
        elif label in {"类型", "注册资本", "成立日期", "住所", "登记状态", "营业期限自", "营业期限至", "经营范围"}:
            values[label] = row[1] if len(row) > 1 else ""
            if label == "类型" and len(row) > 3:
                values["法定代表人"] = row[3]
            elif label == "注册资本" and len(row) > 3:
                values["成立日期"] = row[3]
            elif label == "营业期限自" and len(row) > 3:
                values["营业期限至"] = row[3]

    return [
        [f"统一社会信用代码：{values.get('credit_code', '')}", f"企业名称：{values.get('name', '')}"],
        [f"类型：{values.get('类型', '')}", f"法定代表人：{values.get('法定代表人', '')}"],
        [f"注册资本：{values.get('注册资本', '')}", f"成立日期：{values.get('成立日期', '')}"],
        [f"营业期限自：{values.get('营业期限自', '')}", f"营业期限至：{values.get('营业期限至', '')}"],
        ["登记机关：", "核准日期："],
        [f"登记状态：{values.get('登记状态', '')}"],
        [f"注册地址：{values.get('住所', '')}"],
        [f"许可项目：{values.get('经营范围', '')}"],
    ]


def run_pipeline(
    *,
    project_config: Path,
    pdf_path: Path | None,
    output_dir: Path,
    ocr_adapter: Any,
    llm_adapter: Any = None,
    word_comment_locator_adapter: Any = None,
    qichacha_adapter: Any = None,
    node_inputs: dict[str, Any] | None = None,
    ocr_field_resolver: OcrFieldResolver | None = None,
    template_path: Path | None = None,
    template_page_reader: Any = None,
    report_date: str | None = None,
    manual_inputs_override: dict[str, Any] | None = None,
    ocr_workbook_path: Path | None = None,
    source_overrides: dict[str, Path | None] | None = None,
    workflow_path: Path | None = None,
    prepare_only: bool = False,
    generate_all_narratives: bool = False,
    llm_values_override: dict[str, Any] | None = None,
    progress_callback: Callable[[str, str, str, int | None], None] | None = None,
) -> PipelineResult:
    if word_comment_locator_adapter is None:
        word_comment_locator_adapter = llm_adapter
    config_path = project_config.resolve()
    config = json.loads(config_path.read_text(encoding="utf-8"))
    workflow_file = (
        workflow_path.resolve()
        if workflow_path is not None
        else Path(__file__).with_name("workflow.yaml")
    )
    workflow_payload = json.loads(workflow_file.read_text(encoding="utf-8"))
    contract_result = validate_workflow_contract(workflow_payload, schemas)
    if not contract_result["valid"]:
        raise ValueError(
            "工作流契约校验失败：" + "；".join(contract_result["issues"])
        )
    workflow_definition = schemas.WorkflowDefinition.model_validate(workflow_payload)
    workflow_nodes = {node.name: node for node in workflow_definition.nodes}
    data_manifest_path = Path(__file__).with_name("data_manifest.yaml")
    data_manifest = json.loads(data_manifest_path.read_text(encoding="utf-8"))
    trace_versions = {
        name: str(version)
        for name, version in data_manifest.get("rule_versions", {}).items()
    }
    trace_versions["data_manifest"] = str(data_manifest.get("version", ""))
    recorder = WorkflowTraceRecorder(
        workflow_version=workflow_definition.version,
        contract_version=workflow_definition.contract_version,
        versions=trace_versions,
    )

    def emit_progress(
        node: str,
        step: str,
        message: str,
        percent: int | None = None,
    ) -> None:
        if progress_callback is not None:
            progress_callback(node, step, message, percent)

    def record_node(
        name: str,
        input_payload: dict[str, Any],
        output_payload: dict[str, Any],
        *,
        status: str = "completed",
        node_evidence: list[dict[str, Any]] | None = None,
        node_issues: list[str] | None = None,
    ) -> None:
        # The implementation still has a few legacy internal checkpoints for
        # diagnostics, but the public contract is now the four-node workflow.
        # Ignore those legacy names instead of making them executable nodes.
        if name not in workflow_nodes:
            return
        definition = workflow_nodes[name]
        recorder.record(
            node_name=name,
            input_model=getattr(schemas, definition.input_model),
            output_model=getattr(schemas, definition.output_model),
            input_payload=input_payload,
            output_payload=output_payload,
            status=status,
            evidence=node_evidence or [],
            issues=node_issues or [],
            human_checkpoint=definition.human_checkpoint,
        )

    def record_stage(
        name: str,
        input_payload: dict[str, Any],
        output_payload: dict[str, Any],
        *,
        status: str = "completed",
        node_evidence: list[dict[str, Any]] | None = None,
        node_issues: list[str] | None = None,
    ) -> None:
        """Record one of the four public nodes with its typed contract."""
        record_node(
            name,
            input_payload,
            output_payload,
            status=status,
            node_evidence=node_evidence,
            node_issues=node_issues,
        )

    base = config_path.parent
    template = template_path.resolve() if template_path else _path(base, config["template"])
    pdf = pdf_path.resolve() if pdf_path is not None else None
    output_dir = output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    mapping_path = _path(base, config["mapping"])
    mapping = json.loads(mapping_path.read_text(encoding="utf-8"))
    locations = validate_mapping(mapping)
    static_locations = mapping.get("static_locations", [])
    template_inventory = inventory_template(template)
    annotation_template = (
        _path(base, config["annotation_template"])
        if config.get("annotation_template")
        else template
    )
    if not annotation_template.is_file():
        annotation_template = template
    annotation_inventory = inventory_template(annotation_template)
    same_template_structure = (
        sum(item["record_type"] == "占位符" for item in template_inventory)
        == sum(item["record_type"] == "占位符" for item in annotation_inventory)
    )
    comment_template = (
        same_template_structure
        and any(item.get("comment_texts") for item in annotation_inventory)
    )
    if comment_template:
        locations = build_comment_aware_locations(annotation_inventory, locations)
        locations = align_locations_to_output_template(template_inventory, locations)
    else:
        annotation_template = template
        annotation_inventory = template_inventory
    field_names = {
        item["field_key"]: item["field_name"]
        for item in [*mapping.get("locations", []), *locations, *static_locations]
        if item.get("field_key")
    }
    for field_key, model_field in schemas.ManualBasicInputs.model_fields.items():
        if model_field.description and field_names.get(field_key) in (None, "", field_key):
            field_names[field_key] = str(model_field.description)
    # Internal keys must never leak into Word comments or user-facing review
    # messages. These table fields are configured outside the mapping list,
    # so give them the same business labels a reviewer sees in the template.
    field_names.update(
        {
            "asset_scope_summary_table": "资产负债范围表",
            "long_term_assets_table": "主要长期资产账面记录表",
            "major_long_term_assets": "主要长期资产账面记录",
            "book_net_assets": "审计后账面净资产",
            "commissioning_party_profile": "委托人工商信息",
            "target_company_profile": "被评估单位工商信息",
            "ownership_at_valuation_date": "评估基准日股权结构",
            "trademark_summary": "已申请注册的商标",
            "software_copyrights": "软件著作权",
        }
    )
    routes = load_yellow_routes(config["yellow_routes"])
    # Comments remain the source-of-truth.  The mapping agent is used only
    # for the small residue of locations that the deterministic comment parser
    # cannot classify; it receives a closed list of existing fields and may
    # never overwrite a confirmed mapping or invent a replacement value.
    mapping_agent_issues: list[str] = []
    if llm_adapter is not None and hasattr(llm_adapter, "map_template_locations"):
        allowed_mapping_fields = sorted(
            set(field_names)
            | {route.field_key for route in routes}
        )
        suggestions, mapping_agent_issues = llm_adapter.map_template_locations(
            locations,
            allowed_mapping_fields,
        )
        for location in locations:
            location_id = str(location.get("location_id", ""))
            suggested_field = suggestions.get(location_id)
            if not suggested_field or location.get("field_key"):
                continue
            location["field_key"] = suggested_field
            location["field_name"] = field_names.get(suggested_field, suggested_field)
            location["mapping_agent_assisted"] = True
    inventory_trace = [
        {
            "location_id": item["location_id"],
            "record_type": item["record_type"],
            "context": item["context"],
            "marker": item["marker"],
        }
        for item in template_inventory
    ]
    record_node(
        "inventory",
        {"template_path": str(template)},
        {"locations": inventory_trace},
        node_evidence=[
            {
                "source_kind": "word_template",
                "source_file": template.name,
                "source_locator": "全文占位符及黄色标注",
            }
        ],
    )
    yellow_location_ids = {
        item["location_id"]
        for item in template_inventory
        if item["record_type"] == "黄色标注内容块"
    }
    if not comment_template:
        validate_yellow_routes(routes, expected_location_ids=yellow_location_ids)

    template_hash = _sha256(template)
    issues: list[str] = list(mapping_agent_issues)
    emit_progress("ocr_llm_candidates", "detect_materials", "正在识别上传材料的文件类型", 18)
    emit_progress("ocr_llm_candidates", "parse_excel", "正在读取 Excel 工作表标题和表头", 20)
    with tempfile.TemporaryDirectory(prefix="appraisal-base-") as temporary:
        legacy = run_project(
            config_path,
            output_dir=Path(temporary),
            offline=True,
            report_date=report_date,
            manual_inputs_override=manual_inputs_override,
            source_overrides=source_overrides,
        )
        fields = json.loads((Path(temporary) / "normalized_fields.json").read_text(encoding="utf-8"))
        normalized_evidence = Path(temporary) / "normalized_evidence.json"
        evidence = (
            json.loads(normalized_evidence.read_text(encoding="utf-8"))
            if normalized_evidence.exists()
            else {}
        )
    emit_progress("ocr_llm_candidates", "parse_excel", "Excel 语义解析完成，开始整理候选字段", 24)

    # The valuation object is a controlled user input even though the current
    # template does not mark every occurrence in yellow.  Validate it once and
    # reuse the exact standard wording everywhere in the report.
    if fields.get("valuation_subject_type") not in (None, ""):
        fields["valuation_subject_type"] = validate_valuation_subject_type(fields["valuation_subject_type"])
    fields["asset_approach_method_label"] = _asset_method_label(
        fields.get("selected_valuation_method")
    )

    # A project may have a fixed-size summary table already present in the
    # template but not represented as a yellow location.  Load it explicitly
    # from its configured source so the template's defaults are never reused.
    scope_table = config.get("asset_scope_summary_table")
    if isinstance(scope_table, dict):
        source_name = str(scope_table["source"])
        source_path = _source_path(
            base,
            config,
            source_overrides,
            source_name,
        )
        semantic_scope = fields.get(scope_table["field_key"])
        semantic_scope_evidence = evidence.get(
            scope_table["field_key"],
            {},
        )
        if (
            str(semantic_scope_evidence.get("kind", "")).startswith(
                "semantic_excel"
            )
            and isinstance(semantic_scope, dict)
            and isinstance(semantic_scope.get("rows"), list)
        ):
            scope_rows = semantic_scope["rows"]
            scope_issues = []
        else:
            # A configured source locator is only metadata for the template;
            # it must never become a fallback reader for an arbitrary upload.
            # When semantic matching has no unique result, leave the cells
            # unresolved instead of reading a previous project's coordinates.
            scope_rows = blank_configured_table(scope_table, placeholder="XXX")
            scope_issues = ["未找到唯一的语义财务表，未读取配置坐标"]
        issues.extend(
            f"{scope_table['field_key']}：{message}"
            for message in scope_issues
        )
        fields[scope_table["field_key"]] = {
            "caption": scope_table.get("caption", ""),
            "rows": scope_rows,
        }
        if not str(
            evidence.get(scope_table["field_key"], {}).get("kind", "")
        ).startswith("semantic_excel"):
            evidence[scope_table["field_key"]] = {
                "kind": (
                    config.get("source_lineage", {})
                    .get(source_name, {})
                    .get("kind", source_name)
                    if source_path is not None and not scope_issues
                    else "missing"
                ),
                "file": source_path.name if source_path is not None else "",
                "locator": scope_table.get("source_locator", ""),
            }

    ocr_node_issues: list[str] = []
    if pdf is not None or ocr_workbook_path is not None:
        emit_progress(
            "ocr_llm_candidates",
            "ocr_pdf",
            "正在读取审计 PDF 的 OCR 结构化结果" if ocr_workbook_path else "正在执行审计 PDF OCR",
            28,
        )
    if ocr_workbook_path is not None:
        try:
            normalized = normalized_from_ocr_workbook(ocr_workbook_path.resolve())
        except Exception as exc:
            cache_issue = f"OCR 缓存读取失败：{exc}"
            issues.append(cache_issue)
            ocr_node_issues.append(cache_issue)
            if pdf is not None and ocr_adapter is not None:
                pages, ocr_issues = ocr_adapter.extract(pdf)
                issues.extend(ocr_issues)
                ocr_node_issues.extend(ocr_issues)
                normalized = normalize_ocr_pages(pages)
            else:
                normalized = normalize_ocr_pages([])
    elif pdf is not None:
        pages, ocr_issues = (
            ocr_adapter.extract(pdf)
            if ocr_adapter is not None
            else ([], ["OCR 未配置"])
        )
        issues.extend(ocr_issues)
        ocr_node_issues.extend(ocr_issues)
        normalized = normalize_ocr_pages(pages)
    else:
        normalized = normalize_ocr_pages([])
    emit_progress(
        "ocr_llm_candidates",
        "ocr_pdf",
        "审计 PDF/OCR 解析完成" if pdf is not None or ocr_workbook_path is not None else "未提供 PDF，跳过 OCR",
        34,
    )

    page_counts: dict[int, int] = {}
    for record in [*normalized.get("text_blocks", []), *normalized.get("table_cells", [])]:
        try:
            page_number = int(record.get("page_number", 0))
            page_count = int(record.get("page_count", 0))
        except (TypeError, ValueError):
            continue
        if page_number > 0:
            page_counts[page_number] = max(page_count, len(page_counts))
    ocr_page_summary = [
        {
            "page_number": page_number,
            "page_count": page_counts[page_number],
            "blocks": [],
            "tables": [],
        }
        for page_number in sorted(page_counts)
    ]
    record_node(
        "ocr_pdf",
        {"pdf_path": str(pdf) if pdf is not None else ""},
        {
            "document": {
                "source_file": pdf.name if pdf is not None else "",
                "pages": ocr_page_summary,
                "issues": ocr_node_issues,
            }
        },
        status=(
            "skipped"
            if pdf is None and ocr_workbook_path is None
            else "completed_with_issues"
            if ocr_node_issues
            else "completed"
        ),
        node_evidence=(
            [
                {
                    "source_kind": (
                        "ocr_cache"
                        if ocr_workbook_path
                        else "pdf_ocr"
                    ),
                    "source_file": (
                        ocr_workbook_path.name
                        if ocr_workbook_path
                        else pdf.name
                        if pdf is not None
                        else ""
                    ),
                    "source_locator": "页级 OCR 结构",
                    "text_block_count": len(
                        normalized.get("text_blocks", [])
                    ),
                    "table_cell_count": len(
                        normalized.get("table_cells", [])
                    ),
                }
            ]
            if pdf is not None or ocr_workbook_path is not None
            else []
        ),
        node_issues=ocr_node_issues,
    )

    # The founding-shareholder table is explicitly marked as a PDF audit
    # report lookup in the template.  Keep it separate from the QCC current
    # shareholder table and derive the sentence immediately above it from the
    # same configured OCR cells.
    historical_ownership_matrix: list[list[str]] | None = None
    founding_names: list[str] = []
    ownership_table_spec = config.get("historical_ownership_table")
    if isinstance(ownership_table_spec, dict):
        historical_ownership_matrix, founding_names = _ocr_ownership_matrix(normalized, ownership_table_spec)
        if founding_names:
            fields["founding_shareholder_1"] = founding_names[0]
            fields["founding_shareholder_2"] = founding_names[1] if len(founding_names) > 1 else ""
            evidence["founding_shareholder_1"] = {
                "kind": "pdf_ocr_xlsx",
                "file": (
                    ocr_workbook_path.name
                    if ocr_workbook_path
                    else pdf.name
                    if pdf is not None
                    else ""
                ),
                "locator": f"OCR_表格!{ownership_table_spec.get('table_id', '')}",
            }
            evidence["founding_shareholder_2"] = dict(evidence["founding_shareholder_1"])
    resolver = ocr_field_resolver or _default_ocr_field_resolver
    resolved_ocr, resolver_issues = resolver(normalized, config)
    issues.extend(_keep_unresolved_ocr_issues(resolver_issues, fields))

    ocr_allowed = fields_for_route(routes, RouteKind.PDF_OCR_XLSX)
    ocr_values = _filter_provider(resolved_ocr, ocr_allowed, "PDF OCR/XLSX 解析器", issues)
    ocr_aux_values = resolve_ocr_aux_fields(normalized, config)
    scope_is_semantic_excel = str(
        evidence.get(scope_table.get("field_key", ""), {}).get("kind", "")
    ).startswith("semantic_excel") if isinstance(scope_table, dict) else False
    if (
        isinstance(scope_table, dict)
        and scope_table.get("field_key") in fields
        and not scope_is_semantic_excel
    ):
        scope_rows = fields[scope_table["field_key"]].get("rows", [])
        fields[scope_table["field_key"]]["rows"] = _apply_ocr_overrides_to_table(
            scope_rows, scope_table, ocr_aux_values
        )
        if any(row.get("ocr_field_key") in ocr_aux_values for row in scope_table.get("rows", [])):
            evidence[scope_table["field_key"]] = {
                "kind": "pdf_ocr_xlsx",
                "file": "OCR结构化结果.xlsx",
                "locator": scope_table.get("source_locator", "") + "；语义 OCR 覆盖",
            }
    # For audit-PDF tables, construct the visible table from OCR evidence
    # rather than retaining a semantic Excel matrix.  Partial OCR evidence is
    # useful: matched rows are filled and the remainder keeps its yellow XXX.
    if isinstance(scope_table, dict):
        scope_rows_from_pdf = _apply_ocr_overrides_to_table(
            blank_configured_table(scope_table, placeholder="XXX"),
            scope_table,
            ocr_aux_values,
        )
        if any(row.get("ocr_field_key") in ocr_aux_values for row in scope_table.get("rows", [])):
            resolved_scope = {
                "caption": scope_table.get("caption", ""),
                "rows": scope_rows_from_pdf,
            }
            resolved_ocr[scope_table["field_key"]] = resolved_scope
            ocr_values[scope_table["field_key"]] = resolved_scope
    long_term_table_spec = config.get("long_term_assets_table")
    if isinstance(long_term_table_spec, dict):
        long_term_rows = [["项目", "账面金额（元）", "数量", "现状、特点"]]
        matched_long_term = False
        for row in long_term_table_spec.get("rows", []):
            key = str(row.get("ocr_field_key", ""))
            amount = ocr_aux_values.get(key, "XXX")
            matched_long_term = matched_long_term or amount != "XXX"
            long_term_rows.append([str(row.get("label", "")), str(amount), "XXX", "XXX"])
        if matched_long_term:
            resolved_ocr["long_term_assets_table"] = {"rows": long_term_rows}
    ocr_fallback_fields = set(config.get("ocr_fallback_fields", []))
    ocr_prefer_material_fields = set(config.get("ocr_prefer_material_fields", []))
    formal_ocr_history_fields = {
        field_key
        for field_key in (
            "historical_balance_sheet_table",
            "historical_income_statement_table",
        )
        if _ocr_has_formal_history(normalized, field_key)
    }
    for field_key in ocr_allowed:
        if (
            field_key in ocr_prefer_material_fields
            and field_key not in formal_ocr_history_fields
            and fields.get(field_key) not in (None, "", [])
        ):
            ocr_values[field_key] = fields[field_key]
        if field_key in ocr_fallback_fields and field_key not in ocr_values and fields.get(field_key) not in (None, "", []):
            ocr_values[field_key] = fields[field_key]
        if field_key in ocr_values:
            source = evidence.get(field_key, {})
            normalized["financial_data"].append(
                {
                    "field_key": field_key,
                    "field_name": field_key,
                    "period": "",
                    "value": ocr_values[field_key],
                    "unit": "",
                    "evidence_id": source.get("locator", "ocr:xlsx"),
                }
            )

    # Latest confirmed mapping requires audit-derived fields to be resolved
    # from the audit PDF/OCR first.  Uploaded valuation workbooks remain in
    # the candidate pool only for cross-checking and will produce an explicit
    # issue when their value differs.
    pdf_authoritative_fields = set(
        config.get(
            "pdf_authoritative_fields",
            [
                "historical_balance_sheet_table",
                "historical_income_statement_table",
                "tax_rates",
                "valuation_scope",
                "asset_scope_summary_table",
                "long_term_assets_table",
                "major_long_term_assets",
                "book_net_assets",
            ],
        )
    )
    emit_progress("ocr_llm_candidates", "reconcile_sources", "正在比对 PDF/OCR 与 Excel 的重复字段", 48)
    ocr_locators = pdf_field_locators(normalized, pdf_authoritative_fields, field_names)
    unresolved_pdf_page_keys = [
        key
        for key, locator in ocr_locators.items()
        if not re.search(r"第\d+页", str(locator or ""))
    ]
    if unresolved_pdf_page_keys and llm_adapter is not None and hasattr(llm_adapter, "locate_pdf_pages"):
        emit_progress(
            "ocr_llm_candidates",
            "locate_pdf_pages",
            f"正在用 LLM 补充 {len(unresolved_pdf_page_keys)} 个 PDF 页码定位",
            50,
        )
    ocr_locators, page_locator_issues = _resolve_unresolved_pdf_page_locators(
        field_names,
        ocr_locators,
        normalized,
        llm_adapter,
    )
    issues.extend(page_locator_issues)
    if unresolved_pdf_page_keys and llm_adapter is not None and hasattr(llm_adapter, "locate_pdf_pages"):
        resolved_count = sum(
            1 for key in unresolved_pdf_page_keys if re.search(r"第\d+页", str(ocr_locators.get(key, "")))
        )
        emit_progress(
            "ocr_llm_candidates",
            "locate_pdf_pages",
            f"PDF 页码定位完成：已补充 {resolved_count}/{len(unresolved_pdf_page_keys)} 项",
            51,
        )
    fields, evidence, reconciliation_issues, reconciliation_rows = _reconcile_pdf_authoritative_fields(
        fields=fields,
        evidence=evidence,
        resolved_ocr=resolved_ocr,
        field_keys=pdf_authoritative_fields,
        pdf_name=pdf.name if pdf is not None else "",
        ocr_source_name=ocr_workbook_path.name if ocr_workbook_path is not None else "",
        ocr_locators=ocr_locators,
        source_overrides=source_overrides,
    )
    issues.extend(reconciliation_issues)
    emit_progress("ocr_llm_candidates", "reconcile_sources", "来源整合完成，差异已记录", 52)
    word_table_indices = _word_table_indices(config)
    word_anchor_hints = _word_field_anchor_hints([*locations, *static_locations])
    reviewable_source_conflicts = _attach_word_anchor_hints(
        _reviewable_source_conflicts(
            reconciliation_rows,
            field_names,
            word_table_indices,
        ),
        word_anchor_hints,
    )
    reviewable_source_fallbacks = _attach_word_anchor_hints(
        _reviewable_source_fallbacks(
            reconciliation_rows,
            field_names,
            word_table_indices,
        ),
        word_anchor_hints,
    )
    llm_evidence_reviews: list[dict[str, str]] = []
    llm_evidence_review_issues: list[str] = []
    llm_review_rows = _llm_review_rows(fields, evidence, reconciliation_rows)
    if llm_adapter is not None and hasattr(llm_adapter, "review_extracted_evidence"):
        emit_progress(
            "ocr_llm_candidates",
            "review_evidence",
            f"LLM 正在逐项复核 {len(llm_review_rows)} 个表/字段证据",
            58,
        )
        try:
            def review_progress(done: int, total: int, label: str) -> None:
                if total <= 0:
                    return
                pct = 58 + int(3 * min(done, total) / total)
                current = min(max(done, 1), total)
                emit_progress(
                    "ocr_llm_candidates",
                    "review_evidence",
                    f"LLM 正在复核第 {current}/{total} 项：{label}（对照 PDF 页码与 Excel 工作表）",
                    pct,
                )

            reviewer = llm_adapter.review_extracted_evidence
            if "progress_callback" in inspect.signature(reviewer).parameters:
                llm_evidence_reviews, llm_evidence_review_issues = reviewer(
                    llm_review_rows,
                    field_names,
                    progress_callback=review_progress,
                )
            else:
                llm_evidence_reviews, llm_evidence_review_issues = reviewer(
                    llm_review_rows,
                    field_names,
                )
        except Exception as exc:
            llm_evidence_review_issues = [f"LLM 取数复核调用失败：{exc}"]
    issues.extend(llm_evidence_review_issues)
    for item in llm_evidence_reviews:
        if item.get("status") in {"needs_review", "conflict", "missing"}:
            issues.append(
                f"LLM取数复核：{item.get('field_key', '')}={item.get('status', '')}；"
                f"{item.get('reason', '')}"
            )
    reviewable_llm_notes = _attach_word_anchor_hints(
        _reviewable_llm_notes(
            llm_evidence_reviews,
            llm_review_rows,
            field_names,
            word_table_indices,
        ),
        word_anchor_hints,
    )
    emit_progress(
        "ocr_llm_candidates",
        "review_evidence",
        (
            f"LLM 已完成 {len(llm_evidence_reviews)} 个表/字段复核并生成审核批注"
            if llm_adapter is not None
            else "未启用 LLM，跳过批注生成"
        ),
        61,
    )
    for field_key in pdf_authoritative_fields:
        ocr_values[field_key] = fields.get(field_key, "")
    write_json(
        output_dir / "来源整合与差异.json",
        {
            "fields": reconciliation_rows,
            "llm_review_fields": llm_review_rows,
            "word_review_annotations": reviewable_source_conflicts,
            "word_source_fallback_notes": reviewable_source_fallbacks,
            "llm_reviews": llm_evidence_reviews,
            "word_llm_review_notes": reviewable_llm_notes,
        },
    )
    llm_evidence_review_path = write_json(
        output_dir / "LLM取数复核.json",
        {
            "version": "evidence_review.v1",
            "reviews": llm_evidence_reviews,
            "issues": llm_evidence_review_issues,
        },
    )
    ocr_workbook = (
        export_ocr_workbook(
            output_dir / "OCR结构化结果.xlsx",
            normalized,
        )
        if pdf is not None or ocr_workbook_path is not None
        else None
    )
    record_node(
        "export_ocr_workbook",
        {
            "normalized_ocr": {
                "text_blocks": [
                    {"count": len(normalized.get("text_blocks", []))}
                ],
                "table_cells": [
                    {"count": len(normalized.get("table_cells", []))}
                ],
                "financial_data": [
                    {"count": len(normalized.get("financial_data", []))}
                ],
                "issues": [{"count": len(normalized.get("issues", []))}],
            },
            "output_path": str(ocr_workbook) if ocr_workbook else "",
        },
        {"workbook_path": str(ocr_workbook) if ocr_workbook else ""},
        status="completed" if ocr_workbook else "skipped",
        node_evidence=(
            [
                {
                    "source_kind": "generated_artifact",
                    "source_file": ocr_workbook.name,
                    "source_locator": "OCR_文本、OCR_表格、标准财务数据、识别问题",
                }
            ]
            if ocr_workbook
            else []
        ),
    )

    configured_source_paths = {}
    for name in config.get("sources", {}):
        path = _source_path(base, config, source_overrides, name)
        if path is not None:
            configured_source_paths[name] = str(path)
    base_candidates = trace_candidates(fields, evidence, field_names)
    record_node(
        "extract_sources",
        {"sources": configured_source_paths},
        {
            "candidates": base_candidates,
            "issues": list(legacy.issues),
        },
        status="completed_with_issues" if legacy.issues else "completed",
        node_evidence=[
            {
                "source_kind": "source_materials",
                "source_file": Path(path).name,
                "source_locator": name,
            }
            for name, path in configured_source_paths.items()
        ],
        node_issues=list(legacy.issues),
    )

    manual_path = _path(base, config["manual_inputs"])
    configured_inputs = (
        {}
        if manual_inputs_override is not None
        else json.loads(manual_path.read_text(encoding="utf-8"))
        if manual_path.exists()
        else {}
    )
    if manual_inputs_override:
        configured_inputs.update({key: value for key, value in manual_inputs_override.items() if value not in (None, "")})
    node_allowed = fields_for_route(routes, RouteKind.NODE_INPUT)
    node_values = {
        field_key: value
        for field_key, value in configured_inputs.items()
        if field_key in node_allowed and value not in (None, "", [])
    }
    node_values.update(_filter_provider(node_inputs or {}, node_allowed, "节点输入", issues))
    if node_values.get("selected_valuation_method") not in (None, ""):
        node_values["selected_valuation_method"] = normalize_valuation_methods(
            node_values["selected_valuation_method"]
        )

    qcc_allowed = fields_for_route(routes, RouteKind.QICHACHA_API)
    # The 0817 input contract lets the reviewer choose the authoritative
    # provider for four business modules.  Keep the choice field-scoped:
    # selecting an uploaded file must never result in a silent QCC fallback.
    qcc_groups = {
        "registry_info_strategy": {
            "commissioning_party_profile",
            "target_company_profile",
        },
        "ownership_history_strategy": {
            "ownership_history",
            "ownership_at_valuation_date",
        },
        "unrecorded_intangibles_strategy": {
            "unrecorded_intangibles",
            "software_copyrights",
            "trademark_summary",
        },
    }
    for strategy_key, grouped_fields in qcc_groups.items():
        if configured_inputs.get(strategy_key, "file") != "qichacha":
            qcc_allowed -= grouped_fields
    qcc_company_profile_enabled = (
        configured_inputs.get("company_profile_strategy", "file") == "qichacha"
    )
    fetch_commissioning_qcc = "commissioning_party_profile" in qcc_allowed
    fetch_target_qcc = bool(qcc_allowed) or qcc_company_profile_enabled
    qcc_values: dict[str, Any] = {}
    qcc_profiles: dict[str, dict[str, Any]] = {}
    qcc_payloads: dict[str, dict[str, Any]] = {}
    software_no_result = False
    qcc_provider_issues: list[str] = []
    commissioning_name = str(fields.get("commissioning_party_name", ""))
    target_name = str(fields.get("target_company_name", ""))
    qcc_snapshot_path = output_dir / "qichacha_result.json"
    if not prepare_only and qcc_snapshot_path.is_file() and fetch_target_qcc:
        try:
            snapshot = json.loads(qcc_snapshot_path.read_text(encoding="utf-8"))
            saved_payloads = snapshot.get("payloads", {})
            expected_roles = {
                role: name
                for role, name in {
                    "commissioning": commissioning_name,
                    "target": target_name,
                }.items()
                if name
            }
            validated_payloads = {}
            for role, name in expected_roles.items():
                payload = _validated_qcc_payload(
                    saved_payloads.get(role, {}),
                    name,
                    "委托人" if role == "commissioning" else "被评估单位",
                    issues,
                )
                if payload:
                    validated_payloads[role] = payload
            qcc_payloads.update(validated_payloads)
            qcc_profiles.update(
                {
                    role: payload.get("profile", {})
                    for role, payload in validated_payloads.items()
                }
            )
            if "commissioning" in validated_payloads and fetch_commissioning_qcc:
                qcc_values.update(
                    _filter_provider(
                        validated_payloads["commissioning"],
                        {"commissioning_party_profile"},
                        "企查查快照（委托人）",
                        issues,
                    )
                )
            if "target" in validated_payloads:
                qcc_values.update(
                    _filter_provider(
                        validated_payloads["target"],
                        qcc_allowed - {"commissioning_party_profile"},
                        "企查查快照（被评估单位）",
                        issues,
                    )
                )
            software_no_result = bool(snapshot.get("software_no_result"))
            qcc_provider_issues.extend(
                str(item) for item in snapshot.get("issues", [])
            )
        except (OSError, ValueError, TypeError) as exc:
            issues.append(f"企查查快照读取失败，改用实时接口：{exc}")
    if qichacha_adapter is not None:
        emit_progress("ocr_llm_candidates", "query_qichacha", "企查查 API 正在搜索企业信息", 62)
        if (
            fetch_commissioning_qcc
            and commissioning_name
            and "commissioning" not in qcc_payloads
        ):
            payload, provider_issues = qichacha_adapter.fetch(commissioning_name)
            issues.extend(provider_issues)
            qcc_provider_issues.extend(provider_issues)
            software_no_result = software_no_result or any("接口 233 返回 201" in issue for issue in provider_issues)
            payload = _validated_qcc_payload(payload, commissioning_name, "委托人", issues)
            qcc_payloads["commissioning"] = payload
            qcc_profiles["commissioning"] = qcc_payloads["commissioning"].get("profile", {})
            qcc_values.update(_filter_provider(payload, {"commissioning_party_profile"}, "企查查 API（委托人）", issues))
        if fetch_target_qcc and target_name and "target" not in qcc_payloads:
            payload, provider_issues = qichacha_adapter.fetch(target_name)
            issues.extend(provider_issues)
            qcc_provider_issues.extend(provider_issues)
            software_no_result = software_no_result or any("接口 233 返回 201" in issue for issue in provider_issues)
            payload = _validated_qcc_payload(payload, target_name, "被评估单位", issues)
            qcc_payloads["target"] = payload
            qcc_profiles["target"] = qcc_payloads["target"].get("profile", {})
            qcc_values.update(_filter_provider(payload, qcc_allowed - {"commissioning_party_profile"}, "企查查 API（被评估单位）", issues))
            discover = getattr(qichacha_adapter, "discover_listed_comparables", None)
            scope = str(qcc_profiles["target"].get("business_scope", "")).strip()
            if callable(discover) and scope:
                # QCC validates query length; derive compact search terms from
                # the API-returned business scope rather than sending the
                # whole legal sentence. No LLM may invent a peer name here.
                from demo.adapters.company_api import comparable_search_terms

                search_terms = comparable_search_terms(scope)
                comparable_evidence, comparable_issues = discover(search_terms)
                issues.extend(comparable_issues)
                qcc_provider_issues.extend(comparable_issues)
                qcc_payloads["target"].setdefault("evidence", []).extend(comparable_evidence)
        emit_progress("ocr_llm_candidates", "query_qichacha", "企查查 API 查询完成，正在整理返回证据", 64)
    else:
        emit_progress("ocr_llm_candidates", "query_qichacha", "未启用企查查 API，跳过企业信息查询", 64)
    if qcc_profiles.get("commissioning"):
        evidence.setdefault(
            "commissioning_party_profile",
            {
                "kind": "qichacha_api",
                "file": "企查查 API（735）",
                "locator": "委托人工商信息",
            },
        )
    if qcc_profiles.get("target"):
        evidence.setdefault(
            "target_company_profile",
            {
                "kind": "qichacha_api",
                "file": "企查查 API（735）",
                "locator": "被评估单位工商信息",
            },
        )
    target_profile = qcc_profiles.get("target", {})
    if (
        fields.get("registered_capital") in (None, "", [])
        and target_profile.get("registered_capital") not in (None, "")
    ):
        fields["registered_capital"] = target_profile["registered_capital"]
        evidence["registered_capital"] = {
            "kind": "qichacha_api",
            "file": "企查查 API（735）",
            "locator": "企业工商详情.注册资本",
        }

    # The detailed IP records belong in the dedicated tables.  Keep the
    # yellow paragraph as a short, stable cross-reference instead of dumping
    # dozens of patent/trademark rows into one body paragraph.
    target_ip = qcc_payloads.get("target", {})
    patent_count = len(target_ip.get("patent_rows", [])) if isinstance(target_ip, dict) else 0
    trademark_count = len(target_ip.get("trademark_rows", [])) if isinstance(target_ip, dict) else 0
    if (patent_count or trademark_count) and "unrecorded_intangibles" in qcc_allowed:
        details = []
        if patent_count:
            details.append(f"专利{patent_count}项")
        if trademark_count:
            details.append(f"商标{trademark_count}项")
        qcc_values["unrecorded_intangibles"] = "已查询到" + "、".join(details) + "，详细信息见下表。"
        fields["trademark_summary"] = "商标明细见下表。" if trademark_count else fields.get("trademark_summary", "")
        evidence["trademark_summary"] = {
            "kind": "qichacha_api",
            "file": "企查查 API（231）",
            "locator": "全国商标查询",
        }
    software_count = len(target_ip.get("software_rows", [])) if isinstance(target_ip, dict) else 0
    software_query_ok = bool(target_ip.get("software_query_ok")) if isinstance(target_ip, dict) else False
    if not software_count and (software_no_result or software_query_ok) and "software_copyrights" in qcc_allowed:
        qcc_values["software_copyrights"] = "未查询到软件著作权登记记录。"
        evidence["software_copyrights"] = {
            "kind": "qichacha_api",
            "file": "企查查 API（233）",
            "locator": "软件著作权查询（有效请求但无结果）",
        }

    all_llm_allowed = fields_for_route(routes, RouteKind.BAILIAN_GLM)
    selected_modules = normalize_narrative_modules(fields.get("narrative_modules"))
    # Candidate generation always covers every fixed LLM slot.  The old
    # pre-generation checkbox is retained only for direct CLI compatibility;
    # the web workflow passes ``generate_all_narratives=True`` and applies the
    # user's selection in the later fill invocation.
    llm_allowed = (
        all_llm_allowed
        if generate_all_narratives
        else select_narrative_fields(all_llm_allowed, selected_modules)
    )
    llm_values: dict[str, Any] = {}
    llm_source_kind = "user_selected_llm" if llm_values_override is not None else "bailian_glm"
    if llm_values_override is not None:
        # Selection is a closed-world operation: only keys represented by a
        # yellow LLM slot may be written back to the template.
        llm_values = select_llm_candidates(
            llm_values_override,
            list(llm_values_override.keys()),
        )
        llm_values = {
            key: value for key, value in llm_values.items() if key in all_llm_allowed
        }
    elif llm_adapter is not None:
        structured_evidence = []
        for field_key, value in fields.items():
            source = evidence.get(field_key, {})
            if (
                value in (None, "", [], {})
                or source.get("kind") in {"missing", "bailian_glm"}
            ):
                continue
            if isinstance(value, (dict, list)):
                rendered = json.dumps(value, ensure_ascii=False, default=str)
            else:
                rendered = str(value)
            structured_evidence.append(
                {
                    "evidence_id": f"field:{field_key}",
                    "text": (
                        f"{field_names.get(field_key, field_key)}："
                        f"{rendered[:4000]}"
                    ),
                }
            )
        reference_report = _source_path(
            base,
            config,
            source_overrides,
            "reference_report",
        )
        if (
            reference_report is not None
            and reference_report.suffix.lower() == ".docx"
        ):
            try:
                structured_evidence.extend(
                    read_narrative_evidence(
                        reference_report,
                        source_name="reference_report",
                    )
                )
            except (OSError, ValueError, KeyError) as exc:
                issues.append(f"参考 Word 叙述证据读取失败：{exc}")
        # The 0817 input contract lets the user choose uploaded material
        # instead of Qichacha for registry, ownership, intangible, and
        # profile information.  Word materials become evidence for the LLM
        # mapper/narrative generator; they are never treated as an implicit
        # substitute for an audit-PDF financial field.
        for source_name in (
            "registry_material",
            "ownership_history_material",
            "unrecorded_intangibles_material",
            "company_profile_material",
        ):
            material = _source_path(base, config, source_overrides, source_name)
            if material is None or material.suffix.lower() != ".docx":
                continue
            try:
                structured_evidence.extend(
                    read_narrative_evidence(material, source_name=source_name)
                )
            except (OSError, ValueError, KeyError) as exc:
                issues.append(f"{source_name}读取失败：{exc}")
        if target_profile and qcc_company_profile_enabled:
            # QCC 735 is a current registry snapshot.  Only stable identity
            # facts are supplied to the valuation-date narrative; dynamic
            # capital, ownership, company type and personnel must come from
            # dated uploaded material when the report describes the base date.
            stable_profile = {
                key: target_profile.get(key)
                for key in (
                    "name",
                    "credit_code",
                    "establish_date",
                    "address",
                    "business_scope",
                )
                if target_profile.get(key) not in (None, "")
            }
            structured_evidence.append(
                {
                    "evidence_id": "api:qichacha:target:735:profile",
                    "text": "被评估单位工商信息："
                    + json.dumps(stable_profile, ensure_ascii=False, default=str),
                }
            )
        audited_ownership_summary = _ownership_matrix_summary(historical_ownership_matrix)
        if audited_ownership_summary:
            structured_evidence.append(
                {
                    "evidence_id": "field:ownership_at_valuation_date_audited",
                    "text": (
                        "评估基准日审计材料股权结构（优先于企查查当前工商快照）："
                        + audited_ownership_summary
                    ),
                }
            )
        # Optional QCC business APIs contribute facts to the LLM evidence
        # pool, never directly to a Word paragraph. Their codes are enabled
        # explicitly through QICHACHA_EXTRA_API_CODES, so a normal run does
        # not incur new paid calls unexpectedly.
        for role, qcc_payload in (
            qcc_payloads.items() if qcc_company_profile_enabled else []
        ):
            for item in qcc_payload.get("evidence", []) if isinstance(qcc_payload, dict) else []:
                if not isinstance(item, dict) or not item.get("evidence_id") or not item.get("text"):
                    continue
                raw_evidence_id = str(item["evidence_id"])
                if role == "target" and raw_evidence_id.endswith(":735:profile"):
                    # The sanitized stable profile above is the sole 735 input
                    # to LLM narratives. Keep the full live profile for the
                    # deterministic registry table, not for base-date prose.
                    continue
                qcc_prefix = "api:qichacha:"
                evidence_suffix = (
                    raw_evidence_id[len(qcc_prefix):]
                    if raw_evidence_id.startswith(qcc_prefix)
                    else raw_evidence_id
                )
                qualified_evidence_id = (
                    raw_evidence_id
                    if evidence_suffix.startswith(("commissioning:", "target:"))
                    else f"{qcc_prefix}{role}:{evidence_suffix}"
                )
                structured_evidence.append(
                    {
                        "evidence_id": qualified_evidence_id,
                        "text": f"{role}企查查 ApiCode {item.get('api_code', '')}：{item['text']}",
                    }
                )
        # Node 2 runs before the reviewer has selected modules.  The web form
        # therefore sends an empty list at this stage; that means “produce all
        # candidates”, not “do not produce narratives”.
        llm_request_modules = (
            list(SELECTABLE_LLM_TEMPLATE_FIELDS)
            if generate_all_narratives
            else selected_modules
        )
        llm_evidence = {
            "selected_modules": llm_request_modules,
            "evidence": [
                {"evidence_id": item["evidence_id"], "text": item["text"]}
                for item in [*normalized["text_blocks"], *normalized["table_cells"]]
            ] + structured_evidence,
        }
        if prepare_only:
            # Keep the exact, already-filtered evidence used for the first
            # candidate generation so a later single-module regeneration can
            # use the same source set without rerunning OCR and Excel parsing.
            write_json(output_dir / "llm候选证据.json", llm_evidence)
        candidate_count = len(llm_request_modules) + (1 if "company_profile_section" in llm_allowed else 0)
        emit_progress(
            "ocr_llm_candidates",
            "generate_candidates",
            f"LLM 正在分别生成 {candidate_count} 个候选模块（每个模块独立调用）",
            65,
        )
        payload, provider_issues = llm_adapter.generate(llm_evidence)
        issues.extend(provider_issues)
        llm_values = _filter_provider(payload, llm_allowed, "百炼 GLM", issues)
        emit_progress(
            "ocr_llm_candidates",
            "generate_candidates",
            f"已返回 {len(llm_values)} 个 LLM 候选模块，等待人工选择",
            68,
        )
    # Narrative fields are the only fields allowed to use a project-level local
    # fallback.  This keeps financial and legal facts fail-closed while allowing
    # a report to be generated when an external LLM returns an unusable or
    # partial payload.  Merge per field rather than only when the whole payload
    # is empty: GLM may return ``main_products`` while omitting one of the other
    # numbered overview slots, and a blank numbered slot is a formatting defect.
    if llm_adapter is not None and source_overrides is None:
        fallback = config.get("llm_fallback_fields", {})
        if isinstance(fallback, dict):
            fallback_values = {
                key: str(value)
                for key, value in fallback.items()
                if key in llm_allowed and value not in (None, "", [])
            }
            missing_fallback = {
                key: value for key, value in fallback_values.items()
                if key not in llm_values or llm_values.get(key) in (None, "", [])
            }
            if missing_fallback:
                llm_values.update(missing_fallback)
                llm_source_kind = "bailian_glm_evidence_fallback"
                issues.append("百炼 GLM 未返回全部授权叙述字段，缺失字段已使用项目配置的 PDF 证据化回填。")

    providers = {
        RouteKind.PDF_OCR_XLSX: (ocr_values, "pdf_ocr_xlsx"),
        RouteKind.QICHACHA_API: (qcc_values, "qichacha_api"),
        RouteKind.BAILIAN_GLM: (llm_values, llm_source_kind),
        RouteKind.NODE_INPUT: (node_values, "node_input"),
    }
    for route in routes:
        provider_values, source_kind = providers[route.route_kind]
        value = provider_values.get(route.field_key, "")
        fields[route.field_key] = value
        prior_source = evidence.get(route.field_key, {})
        if (
            value in (None, "", [])
            and prior_source.get("kind") == "unfinished_appraisal"
        ):
            issues.append(
                f"{route.field_key}：评估工作簿尚未完成，已留空"
            )
            continue
        preserve_material_locator = route.route_kind == RouteKind.PDF_OCR_XLSX and value
        evidence_fallback = source_kind == "bailian_glm_evidence_fallback" and value
        pdf_name = pdf.name if pdf is not None else ""
        if preserve_material_locator:
            evidence_file = prior_source.get("file", "") or pdf_name
        elif evidence_fallback:
            evidence_file = pdf_name
        elif source_kind == "qichacha_api" and value:
            evidence_file = "企查查 API（735/231/514/233）"
        elif source_kind == "node_input" and value:
            evidence_file = "人工基础信息"
        elif source_kind.startswith("bailian_glm") and value:
            evidence_file = "百炼大模型"
        else:
            evidence_file = ""
        evidence[route.field_key] = {
            "kind": source_kind if value not in (None, "", []) else "missing",
            "file": evidence_file,
            "locator": (
                prior_source.get("locator", "") or route.location_id
                if preserve_material_locator
                else "PDF OCR 证据化叙述回填"
                if evidence_fallback
                else field_names.get(route.field_key, route.field_key)
                if source_kind in {"node_input", "qichacha_api"} or source_kind.startswith("bailian_glm")
                else route.location_id if value not in (None, "", []) else ""
            ),
        }
        if value in (None, "", []):
            issues.append(f"{route.field_key}：指定来源无可用值，已留空")

    audited_ownership_summary = _ownership_matrix_summary(historical_ownership_matrix)
    if audited_ownership_summary:
        fields["ownership_at_valuation_date"] = audited_ownership_summary
        evidence["ownership_at_valuation_date"] = {
            "kind": "pdf_ocr_xlsx",
            "file": (
                ocr_workbook_path.name
                if ocr_workbook_path
                else pdf.name
                if pdf is not None
                else ""
            ),
            "locator": f"OCR_表格!{ownership_table_spec.get('table_id', '')}",
        }

    # The template has one body placeholder below “3、被评估单位概述”. Its
    # comment lists the six optional node-2 modules; it does not provide six
    # separate Word destinations. Compose only the accepted candidates into
    # that one paragraph, preserving the template's typography and position.
    profile_body = compose_company_profile_narrative(
        str(fields.get("company_profile_section", "")),
        fields,
    )
    if profile_body:
        fields["company_profile_text"] = profile_body
        evidence["company_profile_text"] = {
            "kind": "bailian_glm_profile_composite",
            "file": pdf.name if pdf is not None else "",
            "locator": "company_profile_section + 用户选中的概述模块",
        }

    # The six overview slots are numbered sub-items in the template.  LLM
    # responses are allowed to include or omit the marker, but the generated
    # report must contain exactly one marker at each slot.
    for field_key, prefix in config.get("narrative_prefixes", {}).items():
        value = str(fields.get(field_key, "") or "").strip()
        if not value:
            continue
        value = re.sub(rf"^(?:{re.escape(prefix)})+", "", value).strip()
        fields[field_key] = prefix + value

    non_narrative_values = {
        field_key: value
        for field_key, value in fields.items()
        if field_key not in llm_allowed
    }
    non_narrative_candidates = trace_candidates(
        non_narrative_values,
        evidence,
        field_names,
    )
    non_narrative_resolved = trace_resolved_fields(
        non_narrative_values,
        evidence,
        field_names,
    )
    record_node(
        "resolve_fields",
        {
            "candidates": non_narrative_candidates,
            "mappings": trace_mappings(locations),
        },
        {"fields": non_narrative_resolved},
        status="completed_with_issues" if issues else "completed",
        node_evidence=[
            trace_evidence(evidence_item)
            for evidence_item in evidence.values()
            if isinstance(evidence_item, dict)
        ],
        node_issues=list(issues),
    )
    record_node(
        "select_narrative_modules",
        {"selected_modules": selected_modules},
        {"selected_modules": selected_modules},
        node_evidence=[
            {
                "source_kind": "node_input",
                "source_file": "人工基础信息",
                "source_locator": "主体概况模块",
            }
        ],
    )
    narrative_resolved = trace_resolved_fields(
        fields,
        evidence,
        field_names,
        keys=llm_allowed,
    )
    narrative_prompt_version = getattr(
        llm_adapter,
        "prompt_version",
        "yellow_narratives.v1",
    )
    recorder.versions["narrative_model"] = str(
        getattr(llm_adapter, "model", "") or ""
    )
    recorder.versions["narrative_prompt"] = str(narrative_prompt_version)
    record_node(
        "generate_narrative",
        {"fields": non_narrative_resolved},
        {
            "fields": narrative_resolved,
            "prompt_version": str(narrative_prompt_version),
        },
        status=(
            "skipped"
            if llm_adapter is None
            else "completed_with_issues"
            if any("LLM" in issue or "GLM" in issue for issue in issues)
            else "completed"
        ),
        node_evidence=[
            trace_evidence(evidence.get(field_key))
            for field_key in llm_allowed
            if fields.get(field_key) not in (None, "", [], {})
        ],
        node_issues=[
            issue
            for issue in issues
            if "LLM" in issue or "GLM" in issue
        ],
    )

    if prepare_only:
        if qcc_payloads:
            write_json(
                output_dir / "qichacha_result.json",
                {
                    "payloads": qcc_payloads,
                    "software_no_result": software_no_result,
                    "issues": qcc_provider_issues,
                },
            )
        candidate_locations: dict[str, list[str]] = defaultdict(list)
        for location in locations:
            field_key = str(location.get("field_key") or "")
            if field_key in all_llm_allowed:
                candidate_locations[field_key].append(str(location["location_id"]))
        profile_body_locations = [
            str(location["location_id"])
            for location in locations
            if str(location.get("field_key") or "") == "company_profile_text"
        ]
        candidate_fields = {
            key: value
            for key, value in llm_values.items()
            if key in all_llm_allowed and value not in (None, "", [], {})
        }
        automatic_fields = (
            {"company_profile_section": str(candidate_fields["company_profile_section"])}
            if candidate_fields.get("company_profile_section") not in (None, "", [], {})
            else {}
        )
        candidate_payload = {
            "workflow_stage": "ocr_llm_candidates",
            "selection_required": True,
            "prompt_version": str(narrative_prompt_version),
            "model": str(getattr(llm_adapter, "model", "")),
            "automatic_fields": automatic_fields,
            "candidates": [
                {
                    "field_key": key,
                    "field_name": NARRATIVE_MODULE_LABELS[key],
                    "value": str(candidate_fields.get(key, "")),
                    # The six optional choices share the one profile-body
                    # placeholder in the final template.
                    "location_ids": candidate_locations.get(key, []) or profile_body_locations,
                    "available": key in candidate_fields,
                    "selected": False,
                }
                for key in SELECTABLE_LLM_TEMPLATE_FIELDS
            ],
        }
        candidate_path = write_json(output_dir / "llm候选内容.json", candidate_payload)
        trace_manual_inputs = {
            key: value
            for key, value in (manual_inputs_override or {}).items()
            if key in schemas.ManualBasicInputs.model_fields
        }
        record_stage(
            "start_input",
            {
                "manual_inputs": trace_manual_inputs,
                "materials": {
                    key: str(value)
                    for key, value in (source_overrides or {}).items()
                    if value is not None
                },
                "template_path": str(template),
            },
            {
                "accepted": True,
                "material_roles": [
                    key for key, value in (source_overrides or {}).items() if value is not None
                ],
                "template_path": str(template),
                "issues": [],
            },
        )
        record_stage(
            "ocr_llm_candidates",
            {
                "materials": {
                    key: str(value)
                    for key, value in (source_overrides or {}).items()
                    if value is not None
                },
                "pdf_present": pdf is not None,
                "llm_enabled": llm_adapter is not None,
            },
            {
                "ocr_performed": bool(pdf is not None and ocr_workbook_path is None),
                "ocr_workbook_path": str(ocr_workbook) if ocr_workbook else None,
                "candidates": candidate_payload["candidates"],
                "selection_required": True,
            },
            status="completed_with_issues" if issues else "completed",
            node_evidence=[
                {
                    "source_kind": "bailian_glm",
                    "source_file": "百炼模型",
                    "source_locator": "Word固定LLM位置",
                }
            ] if llm_adapter is not None else [],
            node_issues=list(issues),
        )
        record_stage(
            "fill_word",
            {
                "template_path": str(template),
                "selected_llm_fields": {},
                "deterministic_sources": {},
            },
            {"report_path": "", "replacement_count": 0, "unresolved_count": 0},
            status="skipped",
        )
        record_stage(
            "output",
            {"report_path": ""},
            {"report_path": ""},
            status="skipped",
        )
        trace_path = recorder.export(output_dir / "workflow_trace.json")
        manifest_path = write_json(
            output_dir / "run_manifest.json",
            {
                "project_id": config["project_id"],
                "created_at": datetime.now().isoformat(timespec="seconds"),
                "template": str(template),
                "template_sha256": template_hash,
                "pdf": str(pdf) if pdf else "",
                "pdf_sha256": _sha256(pdf) if pdf else "",
                "workflow_version": workflow_definition.version,
                "workflow_contract_version": workflow_definition.contract_version,
                "status": "awaiting_selection",
                "outputs": [
                    *([str(ocr_workbook)] if ocr_workbook else []),
                    str(candidate_path),
                    str(trace_path),
                ],
            },
        )
        return PipelineResult(
            output_dir / "资产评估报告_待复核.docx",
            ocr_workbook,
            manifest_path,
            issues,
            candidate_path,
            candidate_fields,
        )

    required_monetary_fields = list(config.get("required_monetary_fields", []))
    monetary_gate = require_financial_fields(fields, required_monetary_fields)
    monetary_policy = apply_missing_field_policy(
        fields,
        evidence,
        required_monetary_fields,
        "金额及财务结果字段",
    )
    fields = monetary_policy["fields"]
    evidence = monetary_policy["evidence"]
    financial_issues = list(monetary_policy["issues"])
    if monetary_gate["conflicts"]:
        financial_issues.append(
            "高优先级：金额及财务结果存在冲突，已保留待人工复核："
            + json.dumps(
                monetary_gate["conflicts"],
                ensure_ascii=False,
                default=str,
            )
        )
    issues.extend(financial_issues)
    financial_validation = {
        "valid": monetary_policy["valid"] and not monetary_gate["conflicts"],
        "missing_fields": monetary_policy["missing_fields"],
        "conflicts": monetary_gate["conflicts"],
    }

    replacements = build_replacements(locations, fields)
    table_replacements = {}
    table_column_ratios: dict[int, list[float]] = {}
    table_specs = list(config.get("financial_tables", []))
    scope_table = config.get("asset_scope_summary_table")
    if isinstance(scope_table, dict):
        table_specs.insert(0, scope_table)
    for spec in table_specs:
        value = fields.get(spec["field_key"])
        if isinstance(value, dict) and isinstance(value.get("rows"), list):
            table_replacements[int(spec["target_table_index"])] = value["rows"]
        else:
            table_replacements[int(spec["target_table_index"])] = blank_configured_table(
                spec,
                placeholder="XXX",
            )
        if str(spec.get("field_key", "")).startswith("historical_"):
            matrix = table_replacements[int(spec["target_table_index"])]
            column_count = len(matrix[0]) if matrix and matrix[0] else 0
            if column_count > 1:
                value_ratio = 0.68 / (column_count - 1)
                table_column_ratios[int(spec["target_table_index"])] = [
                    0.32,
                    *([value_ratio] * (column_count - 1)),
                ]
    if historical_ownership_matrix and len(historical_ownership_matrix) > 1:
        table_replacements[int(ownership_table_spec["target_table_index"])] = historical_ownership_matrix
    elif isinstance(ownership_table_spec, dict):
        table_replacements[int(ownership_table_spec["target_table_index"])] = [
            list(
                ownership_table_spec.get(
                    "header",
                    ["序号", "股东名称", "总出资（元）", "股权比例"],
                )
            ),
            ["XXX", "XXX", "XXX", "XXX"],
        ]
    # These tables are present in the communication template but were not
    # represented by yellow paragraphs.  Fill them explicitly so template
    # defaults (such as the listed-company capital) cannot leak into output.
    table_replacements[0] = _company_profile_table(
        qcc_profiles.get("commissioning", {}),
        fields.get("commissioning_party_name", ""),
    )
    table_replacements[1] = _company_profile_table(
        qcc_profiles.get("target", {}),
        fields.get("target_company_name", ""),
        fields.get("registered_capital", ""),
    )
    partner_rows = qcc_payloads.get("target", {}).get("partner_rows", [])
    valuation_date_table_index = int(
        ownership_table_spec.get("valuation_date_target_table_index", 3)
        if isinstance(ownership_table_spec, dict)
        else 3
    )
    table_replacements[valuation_date_table_index] = (
        _valuation_date_ownership_matrix(
            historical_ownership_matrix,
            partner_rows,
        )
    )
    long_term_table = config.get("long_term_assets_table")
    if isinstance(long_term_table, dict):
        semantic_long_term = fields.get("long_term_assets_table")
        if isinstance(semantic_long_term, dict) and isinstance(
            semantic_long_term.get("rows"), list
        ):
            long_term_rows = semantic_long_term["rows"]
        else:
            long_term_rows = _blank_cross_source_table(long_term_table.get("rows", []))
        table_replacements[int(long_term_table["target_table_index"])] = long_term_rows
    trademark_rows = _qcc_table_rows(qcc_payloads.get("target", {}), "trademark_rows", 7)
    software_rows = _qcc_table_rows(qcc_payloads.get("target", {}), "software_rows", 5)
    if trademark_rows:
        table_replacements[8] = [["申请日期", "商标", "商标名称", "注册号", "国际分类", "商标状态", "注册公告日期"], *trademark_rows]
    else:
        table_replacements[8] = [
            ["申请日期", "商标", "商标名称", "注册号", "国际分类", "商标状态", "注册公告日期"],
            ["XXX", "XXX", "XXX", "XXX", "XXX", "XXX", "XXX"],
        ]
    if "software_copyrights" in qcc_allowed:
        table_replacements[9] = [["序号", "软件名称", "登记号", "首次发表日期", "登记批准日期"], ["", "", "", "", ""]]
    if software_rows:
        table_replacements[9] = [["序号", "软件名称", "登记号", "首次发表日期", "登记批准日期"], *software_rows]
    elif software_no_result and "software_copyrights" in qcc_allowed:
        table_replacements[9] = [["序号", "软件名称", "登记号", "首次发表日期", "登记批准日期"], ["", "未查询到软件著作权登记记录。", "", "", ""]]
    elif "software_copyrights" in qcc_allowed:
        table_replacements[9] = [
            ["序号", "软件名称", "登记号", "首次发表日期", "登记批准日期"],
            ["XXX", "XXX", "XXX", "XXX", "XXX"],
        ]

    report = output_dir / "资产评估报告_待复核.docx"
    emit_progress("fill_word", "map_word", "正在按模板批注原文匹配 Word 写入位置", 76)
    fill_template(
        template,
        report,
        replacements,
        table_replacements=table_replacements,
        table_column_ratios=table_column_ratios,
        paragraph_replacements=_paragraph_replacements(config, fields),
        replacement_modes={route.location_id: route.replacement_mode for route in routes},
        progress_callback=lambda step, message: emit_progress("fill_word", step, message, None),
    )
    replace_transaction_type_literals(report, fields.get("transaction_type"))
    replace_image_markers(report)
    replace_report_number_year(report, fields.get("report_number_year"))
    emit_progress(
        "fill_word",
        "write_comments",
        "正在按表格、科目和期间定位批注，并由 LLM 确认目标位置",
        88,
    )
    emit_progress("fill_word", "check_placeholders", "正在检查未找到证据的 XXX 并保留黄色高亮", 91)
    unresolved_findings = highlight_unresolved_placeholders(report)
    source_conflict_annotations = annotate_source_conflicts(
        report,
        [
            *_attach_llm_comments(
                [*reviewable_source_conflicts, *reviewable_source_fallbacks],
                reviewable_llm_notes,
            ),
            *reviewable_llm_notes,
        ],
        location_selector=word_comment_locator_adapter,
    )
    review_by_field = {
        str(item.get("field_key", "")): item
        for item in llm_evidence_reviews
        if item.get("field_key")
    }
    status_by_field = _trace_status_by_field(
        evidence,
        llm_evidence_reviews,
        reviewable_source_conflicts,
        reviewable_source_fallbacks,
    )
    table_fields: dict[int, tuple[str, int, set[int]]] = {}
    for spec in table_specs:
        target_index = int(spec["target_table_index"])
        field_key = str(spec["field_key"])
        # The asset-scope table has no header row; financial history tables do.
        first_data_row = 0 if field_key == "asset_scope_summary_table" else 1
        data_columns = {1} if field_key == "asset_scope_summary_table" else set()
        table_fields[target_index] = (field_key, first_data_row, data_columns)
    if isinstance(ownership_table_spec, dict):
        table_fields[int(ownership_table_spec["target_table_index"])] = (
            "ownership_at_valuation_date",
            1,
            set(),
        )
        table_fields[valuation_date_table_index] = (
            "ownership_at_valuation_date",
            1,
            set(),
        )
    table_fields[0] = ("commissioning_party_profile", 0, set())
    table_fields[1] = ("target_company_profile", 0, set())
    if isinstance(long_term_table, dict):
        table_fields[int(long_term_table["target_table_index"])] = (
            "long_term_assets_table",
            1,
            {1, 2, 3},
        )
    table_fields[8] = ("trademark_summary", 1, set())
    table_fields[9] = ("software_copyrights", 1, set())
    model_name = str(getattr(llm_adapter, "model", "") or "")
    trace_annotations = [
        *_paragraph_trace_annotations(
            locations,
            replacements,
            evidence,
            field_names,
            status_by_field,
            review_by_field,
            model_name,
        ),
        *_table_trace_annotations(
            table_replacements,
            table_fields,
            evidence,
            field_names,
            status_by_field,
            review_by_field,
            model_name,
        ),
    ]
    # Cover static/unmapped placeholders as well: every visible yellow XXX gets
    # a titled missing-data comment even when the template has no field mapping.
    trace_annotations.extend(
        {
            "field_key": "unresolved_word_location",
            "field_name": "该未填位置",
            "target": str(item.get("current_text", "XXX")),
            "part": str(item.get("part") or "word/document.xml"),
            "paragraph_index_hint": int(item.get("paragraph_index") or 0),
            "target_occurrence": int(item.get("occurrence_index") or 1),
            "context_hint": str(item.get("context") or "").replace(
                str(item.get("current_text", "XXX")),
                "",
            ),
            **(
                {
                    "table_index": int(item["table_index"]) - 1,
                    "row_index": int(item["row_index"]) - 1,
                    "column_index": int(item["column_index"]) - 1,
                }
                if item.get("table_index") not in (None, "")
                else {}
            ),
            "status": "missing",
            "comment": (
                "【未找到数据】该位置已检索本次上传的审计PDF、Excel、补充材料及已启用的企业信息接口，"
                "仍未找到可可靠采用的数据，因此保留黄色XXX。请人工补充或核对原始材料。"
            ),
        }
        for item in unresolved_findings
    )
    if llm_adapter is not None and hasattr(llm_adapter, "write_traceability_comments"):
        emit_progress(
            "fill_word",
            "write_traceability_comments",
            "LLM 正在根据已确认的文件、页码、工作表和复核状态撰写来源批注",
            93,
        )
        try:
            generated_comments, trace_comment_issues = (
                llm_adapter.write_traceability_comments(trace_annotations)
            )
            issues.extend(trace_comment_issues)
            for annotation_index, body in generated_comments.items():
                if not 0 <= annotation_index < len(trace_annotations):
                    continue
                annotation = trace_annotations[annotation_index]
                title = TRACE_TITLES.get(str(annotation.get("status", "")), "")
                annotation["comment"] = f"{title}{body}"
        except Exception as exc:
            issues.append(f"LLM 来源批注撰写失败，已使用规则正文：{exc}")
    traceability_annotations = annotate_traceable_content(report, trace_annotations)
    emit_progress("output", "save_word", "评估报告 Word 已保存，正在做最终文件检查", 96)
    generation_issues = issues_from_word_findings(
        unresolved_findings,
        [*locations, *static_locations],
        fields,
        evidence,
    )
    existing_issue_keys = {
        (item.get("location_id"), item.get("category"))
        for item in generation_issues
    }
    generation_issues.extend(
        item
        for item in issues_from_special_evidence(
            [*locations, *static_locations],
            evidence,
        )
        if (item.get("location_id"), item.get("category"))
        not in existing_issue_keys
    )
    if _sha256(template) != template_hash:
        raise RuntimeError("模板被意外修改")
    record_node(
        "legacy_fill_word",
        {
            "template_path": str(template),
            "output_path": str(report),
            "fields": trace_resolved_fields(fields, evidence, field_names),
        },
        {
            "report_path": str(report),
            "replacement_count": len(replacements),
        },
        node_evidence=[
            {
                "source_kind": "word_template",
                "source_file": template.name,
                "source_locator": "映射位置与表格",
            }
        ],
        status=(
            "completed"
            if financial_validation["valid"]
            else "completed_with_issues"
        ),
        node_issues=financial_issues,
    )
    template_pages: dict[str, int | str] = {}
    generated_pages: dict[str, int | str] = {}
    if template_page_reader is not None:
        try:
            report_page_texts, report_page_issues = (
                template_page_reader.extract(report)
            )
            issues.extend(report_page_issues)
            generated_pages = map_location_pages(
                unresolved_findings,
                report_page_texts,
                document_paragraph_texts(report),
            )
        except Exception as exc:
            issues.append(f"生成报告页码解析失败：{exc}")
        try:
            template_page_texts, template_page_issues = (
                template_page_reader.extract(template)
            )
            issues.extend(template_page_issues)
            template_pages = map_location_pages(
                [*locations, *static_locations],
                template_page_texts,
                document_paragraph_texts(template),
            )
        except Exception as exc:
            issues.append(f"模板页码解析失败：{exc}")
    generation_issues = apply_page_locations(
        generation_issues,
        generated_pages,
        template_pages,
    )
    generation_issues = organize_generation_issues(
        generation_issues
    )
    for issue in generation_issues:
        page = issue.get("page_number") or "页码不可用"
        issues.append(
            f"Word第{page}页 {issue.get('location_description', '')}："
            f"{issue.get('problem', '')}"
        )
    normalized_fields_path = write_json(
        output_dir / "normalized_fields.json",
        fields,
    )
    normalized_evidence_path = write_json(
        output_dir / "normalized_evidence.json",
        evidence,
    )
    planned_manifest_path = output_dir / "run_manifest.json"
    candidate_locations: dict[str, list[str]] = defaultdict(list)
    for location in locations:
        field_key = str(location.get("field_key") or "")
        if field_key in all_llm_allowed:
            candidate_locations[field_key].append(str(location["location_id"]))
    selected_candidate_fields = {
        key: str(value)
        for key, value in fields.items()
        if key in all_llm_allowed and value not in (None, "", [], {})
    }
    record_stage(
        "start_input",
        {
            "manual_inputs": {
                key: value
                for key, value in (manual_inputs_override or {}).items()
                if key in schemas.ManualBasicInputs.model_fields
            },
            "materials": {
                key: str(value)
                for key, value in (source_overrides or {}).items()
                if value is not None
            },
            "template_path": str(template),
        },
        {
            "accepted": True,
            "material_roles": [key for key, value in (source_overrides or {}).items() if value is not None],
            "template_path": str(template),
            "issues": [],
        },
    )
    record_stage(
        "ocr_llm_candidates",
        {
            "materials": {
                key: str(value)
                for key, value in (source_overrides or {}).items()
                if value is not None
            },
            "pdf_present": pdf is not None,
            "llm_enabled": llm_adapter is not None or llm_values_override is not None,
        },
        {
            "ocr_performed": bool(pdf is not None and ocr_workbook_path is None),
            "ocr_workbook_path": str(ocr_workbook) if ocr_workbook else None,
            "candidates": [
                {
                    "field_key": key,
                    "field_name": field_names.get(key, key),
                    "value": value,
                    "location_ids": candidate_locations.get(key, []),
                    "selected": True,
                }
                for key, value in selected_candidate_fields.items()
            ],
            "selection_required": False,
        },
        status="completed_with_issues" if issues else "completed",
        node_issues=[issue for issue in issues if "LLM" in issue or "GLM" in issue],
    )
    record_stage(
        "fill_word",
        {
            "template_path": str(template),
            "selected_llm_fields": selected_candidate_fields,
            "deterministic_sources": {
                key: str(value.get("file", ""))
                for key, value in evidence.items()
                if isinstance(value, dict) and value.get("file")
            },
        },
        {
            "report_path": str(report),
            "replacement_count": len(replacements),
            "unresolved_count": len(generation_issues),
        },
        status="completed_with_issues" if financial_validation["valid"] is False or generation_issues else "completed",
        node_issues=financial_issues,
    )
    record_stage(
        "output",
        {"report_path": str(report)},
        {"report_path": str(report)},
        status="completed_with_issues" if generation_issues else "completed",
        node_issues=list(issues),
    )
    trace_path = recorder.export(output_dir / "workflow_trace.json")
    comment_annotations = [
        {
            "location_id": item["location_id"],
            "context": item.get("context", ""),
            "comment_ids": item.get("comment_ids", []),
            "comment_texts": item.get("comment_texts", []),
            "field_key": item.get("field_key", ""),
            "comment_source_kind": item.get("comment_source_kind", ""),
            "comment_source_instruction": item.get("comment_source_instruction", ""),
            "force_unresolved": bool(item.get("force_unresolved")),
        }
        # ``locations`` contains the authoritative comment-to-field/source
        # mapping, including comment-only anchors; the raw inventory does not.
        for item in locations
        if item.get("comment_texts")
    ]
    comment_path = write_json(output_dir / "template_comments.json", comment_annotations)
    all_comments = extract_word_comments(annotation_template)
    manifest = {
        "project_id": config["project_id"],
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "template": str(template),
        "annotation_template": str(annotation_template),
        "template_sha256": template_hash,
        "pdf": str(pdf) if pdf is not None else "",
        "pdf_sha256": _sha256(pdf) if pdf is not None else "",
        "mapping_version": "1.0.0",
        "yellow_route_version": "yellow_routes.v1",
        "financial_rule_version": "financial_aliases.v1",
        "prompt_version": getattr(llm_adapter, "prompt_version", "yellow_narratives.v1"),
        "llm_models": {
            "narrative": str(getattr(llm_adapter, "model", "")),
            "evidence_review": str(getattr(llm_adapter, "review_model", "")),
        },
        "ocr_cache_reused": bool(ocr_workbook_path),
        "ocr_cache_source": str(ocr_workbook_path) if ocr_workbook_path else "",
        "workflow_version": workflow_definition.version,
        "workflow_contract_version": workflow_definition.contract_version,
        "template_annotations": {
            "kind": "word_comments" if comment_annotations else "yellow_highlight_legacy",
            "comment_count": len(all_comments),
            "annotated_location_count": len(comment_annotations),
        },
        "financial_validation": financial_validation,
        "generation_validation": {
            "valid": not generation_issues,
            "unresolved_count": len(generation_issues),
        },
        "outputs": [
            *([str(ocr_workbook)] if ocr_workbook else []),
            str(report),
            str(normalized_fields_path),
            str(normalized_evidence_path),
            str(llm_evidence_review_path),
            str(trace_path),
            str(comment_path),
        ],
    }
    manifest_path = write_json(planned_manifest_path, manifest)
    return PipelineResult(report, ocr_workbook, manifest_path, issues)
