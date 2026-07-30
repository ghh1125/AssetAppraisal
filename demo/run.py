from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
from collections import defaultdict
from zipfile import BadZipFile
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Any, Mapping

from .adapters.audit import export_audit, write_json
from .adapters.generation_issues import export_generation_issues
from .adapters.excel import (
    read_cells,
    try_read_cells,
    try_read_configured_table,
)
from .adapters.materials import resolve_material_field
from .adapters.ocr_factory import create_ocr_adapter
from .adapters.semantic_excel import extract_workbook_facts
from .adapters.word import (
    fill_template,
    highlight_unresolved_placeholders,
    replace_report_number_year,
)
from .domain.mapping import validate_mapping
from .domain.calculations import derive_system_fields
from .domain.field_validation import (
    apply_missing_field_policy,
    normalize_narrative_modules,
    normalize_report_serial,
    normalize_valuation_methods,
    report_number_year,
    validate_final_valuation_method,
    validate_transaction_type,
    validate_valuation_subject_type,
)
from .domain.replacement import build_replacements
from .domain.financial_matching import blank_configured_table
from .domain.historical_table_merge import merge_historical_tables
from .domain.source_precedence import prefer_semantic_result
from .domain.generation_issues import (
    apply_page_locations,
    issues_from_special_evidence,
    issues_from_word_findings,
    organize_generation_issues,
)


@dataclass
class RunResult:
    report_path: Path
    audit_path: Path
    manifest_path: Path
    issues: list[str]


def _asset_method_label(value: Any) -> str:
    text = str(value or "")
    if "资产基础法" in text:
        return "资产基础法"
    if "市场法" in text:
        return "市场法"
    if "收益法" in text:
        return "收益法"
    return ""


def _load_local_env() -> None:
    """Load the project-local .env without overriding an existing environment.

    This keeps local Demo execution convenient while leaving production
    secret injection to c2m or the process supervisor. Values are never
    printed or written to any run artifact.
    """
    candidates = (Path.cwd() / ".env", Path(__file__).resolve().parents[1] / ".env")
    env_path = next((path for path in candidates if path.exists()), None)
    if env_path is None:
        return
    for raw_line in env_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key, value = key.strip(), value.strip()
        if not key:
            continue
        if len(value) >= 2 and value[0] == value[-1] and value[0] in {"'", '"'}:
            value = value[1:-1]
        os.environ.setdefault(key, value)


def _select_cli_ocr_adapter(
    provider: str | None,
    env: Mapping[str, str],
) -> Any:
    configured = dict(env)
    if provider:
        configured["APPRAISAL_OCR_PROVIDER"] = provider
    return create_ocr_adapter(configured)


def _path(base: Path, value: str) -> Path:
    path = Path(value)
    return path if path.is_absolute() else (base / path).resolve()


def _source_evidence(
    source_name: str,
    sources: dict[str, Path],
    locator: str,
    source_lineage: dict[str, dict[str, str]],
) -> dict[str, str]:
    lineage = source_lineage.get(source_name, {})
    origin_source = lineage.get("origin_source")
    files = [sources[source_name].name]
    if origin_source and origin_source in sources:
        files.append(sources[origin_source].name)
    return {
        "kind": lineage.get("kind", source_name),
        "file": "；".join(files),
        "locator": locator,
    }


def _excel_value(
    record: dict,
    sources: dict[str, Path],
    source_lineage: dict[str, dict[str, str]],
):
    kind = record.get("source_kind", "")
    source_name = "income_workbook" if "收益法" in kind else "reporting_workbook" if "上报表" in kind else None
    if not source_name or source_name not in sources:
        return None, None
    locator = record.get("source_locator", "")
    matches = re.findall(r"([\w\u4e00-\u9fff（）()]+)!([A-Z]+\d+)", locator)
    for sheet, cell in matches:
        try:
            key = f"{sheet}!{cell}"
            value = read_cells(sources[source_name], [key])[key]
            if value not in (None, ""):
                return value, _source_evidence(source_name, sources, key, source_lineage)
        except (KeyError, ValueError):
            continue
    return None, None


def _read_long_term_assets_table(
    config: dict[str, Any],
    sources: dict[str, Path],
    issues: list[str],
) -> list[list[str]]:
    matrix = [["项目", "账面金额（元）", "数量", "现状、特点"]]
    for row in config.get("long_term_assets_table", {}).get("rows", []):
        locator = str(row["locator"])
        source_name = str(row["source"])
        values, read_issues = try_read_cells(sources.get(source_name), [locator])
        value = values.get(locator, "XXX")
        issues.extend(
            f"{row.get('label', source_name)}：{message}"
            for message in read_issues
        )
        if isinstance(value, (int, float)):
            value = f"{value:,.2f}"
        matrix.append([
            str(row.get("label", "")),
            str(value if value not in (None, "") else "XXX"),
            str(row.get("quantity", "")),
            str(row.get("condition", "")),
        ])
    return matrix


def _provider_values(payload: Any) -> dict[str, Any]:
    if isinstance(payload, dict):
        nested = payload.get("fields")
        return nested if isinstance(nested, dict) else payload
    if isinstance(payload, list):
        merged: dict[str, Any] = {}
        for item in payload:
            merged.update(_provider_values(item))
        return merged
    return {}


def _merge_provider(
    payload: Any,
    *,
    kind: str,
    known_fields: set[str],
    fields: dict[str, Any],
    evidence: dict[str, dict],
) -> None:
    for key, value in _provider_values(payload).items():
        if key in known_fields and value not in (None, "", []) and key not in fields:
            fields[key] = value
            evidence[key] = {"kind": kind, "file": kind, "locator": key}


def run_project(
    config_path: Path,
    output_dir: Path | None = None,
    offline: bool = False,
    *,
    report_date: str | None = None,
    ocr_adapter: Any = None,
    company_api_adapter: Any = None,
    llm_adapter: Any = None,
    manual_inputs_override: dict[str, Any] | None = None,
    source_overrides: dict[str, Path | None] | None = None,
) -> RunResult:
    config_path = config_path.resolve()
    config = json.loads(config_path.read_text(encoding="utf-8"))
    base = config_path.parent
    template = _path(base, config["template"])
    mapping_path = _path(base, config["mapping"])
    manual_path = _path(base, config["manual_inputs"])
    sources = {name: _path(base, value) for name, value in config.get("sources", {}).items()}
    if source_overrides is not None:
        for name, path in source_overrides.items():
            if path is None:
                sources.pop(name, None)
            else:
                sources[name] = Path(path).resolve()
    source_lineage = config.get("source_lineage", {})
    mapping = json.loads(mapping_path.read_text(encoding="utf-8"))
    locations = validate_mapping(mapping)
    static_locations = mapping.get("static_locations", [])
    manual = (
        {}
        if manual_inputs_override is not None
        else json.loads(manual_path.read_text(encoding="utf-8"))
        if manual_path.exists()
        else {}
    )
    if manual_inputs_override:
        manual.update({key: value for key, value in manual_inputs_override.items() if value not in (None, "")})
    if manual.get("report_serial") not in (None, ""):
        raw_report_serial = manual["report_serial"]
        manual["report_serial"] = normalize_report_serial(raw_report_serial)
        parsed_report_year = report_number_year(raw_report_serial)
        if parsed_report_year:
            manual["report_number_year"] = parsed_report_year
    if manual.get("valuation_subject_type") not in (None, ""):
        manual["valuation_subject_type"] = validate_valuation_subject_type(manual["valuation_subject_type"])
    if manual.get("selected_valuation_method") not in (None, ""):
        manual["selected_valuation_method"] = normalize_valuation_methods(manual["selected_valuation_method"])
    if manual.get("final_valuation_method") not in (None, ""):
        manual["final_valuation_method"] = validate_final_valuation_method(manual["final_valuation_method"])
    if manual.get("transaction_type") not in (None, ""):
        manual["transaction_type"] = validate_transaction_type(manual["transaction_type"])
    manual["narrative_modules"] = normalize_narrative_modules(manual.get("narrative_modules"))
    fields: dict[str, object] = {}
    evidence: dict[str, dict] = {}
    issues: list[str] = []
    table_replacements: dict[int, list[list[str]]] = {}
    records_by_key = {record["field_key"]: record for record in locations}
    for record in locations:
        key = record["field_key"]
        if key in fields:
            continue
        if manual.get(key) not in (None, ""):
            fields[key] = manual[key]
            evidence[key] = {"kind": "manual", "file": manual_path.name, "locator": key}
            continue
        value, source = _excel_value(record, sources, source_lineage)
        if value not in (None, ""):
            fields[key] = value
            evidence[key] = source or {}

    for key, value in manual.items():
        if value not in (None, "") and key not in fields:
            fields[key] = value
            evidence[key] = {"kind": "manual", "file": manual_path.name, "locator": key}
    fields["asset_approach_method_label"] = _asset_method_label(fields.get("selected_valuation_method"))

    for spec in config.get("financial_tables", []):
        source_name = spec["source"]
        matrix, read_issues = try_read_configured_table(
            sources.get(source_name),
            spec,
        )
        if matrix is None:
            matrix = blank_configured_table(spec, placeholder="XXX")
        issues.extend(
            f"{spec['field_key']}：{message}" for message in read_issues
        )
        key = spec["field_key"]
        fields[key] = {"caption": spec["caption"], "rows": matrix}
        evidence[key] = (
            _source_evidence(
                source_name,
                sources,
                spec["source_locator"],
                source_lineage,
            )
            if source_name in sources and not read_issues
            else {
                "kind": "missing",
                "file": "",
                "locator": spec["source_locator"],
            }
        )
        table_replacements[int(spec["target_table_index"])] = matrix

    # The balance-sheet overview is a real table under its lead-in paragraph,
    # not a prose field.  Fill it in every run mode so the CLI cannot leave
    # the template's default numbers behind.
    scope_table = config.get("asset_scope_summary_table")
    if isinstance(scope_table, dict):
        source_name = str(scope_table["source"])
        matrix, read_issues = try_read_configured_table(
            sources.get(source_name),
            scope_table,
        )
        if matrix is None:
            matrix = blank_configured_table(
                scope_table,
                placeholder="XXX",
            )
        issues.extend(
            f"{scope_table['field_key']}：{message}"
            for message in read_issues
        )
        key = str(scope_table["field_key"])
        fields[key] = {"caption": scope_table.get("caption", ""), "rows": matrix}
        evidence[key] = (
            _source_evidence(
                source_name,
                sources,
                scope_table.get("source_locator", ""),
                source_lineage,
            )
            if source_name in sources and not read_issues
            else {
                "kind": "missing",
                "file": "",
                "locator": scope_table.get("source_locator", ""),
            }
        )
        table_replacements[int(scope_table["target_table_index"])] = matrix

    long_term_table = config.get("long_term_assets_table")
    if isinstance(long_term_table, dict):
        matrix = _read_long_term_assets_table(config, sources, issues)
        table_replacements[int(long_term_table["target_table_index"])] = matrix

    # The communication template's second IP table is software copyright,
    # although its legacy header says patent.  Normalize the header even when
    # the QCC provider is unavailable and data rows must remain blank.
    yellow_fields = {item.get("field_key") for item in config.get("yellow_routes", [])}
    if "software_copyrights" in yellow_fields:
        table_replacements[9] = [
            ["序号", "软件名称", "登记号", "首次发表日期", "登记批准日期"],
            ["", "", "", "", ""],
        ]

    for spec in config.get("financial_fields", []):
        source_name = spec["source"]
        locator = spec["locator"]
        values, read_issues = try_read_cells(
            sources.get(source_name),
            [locator],
        )
        issues.extend(
            f"{spec['field_key']}：{message}" for message in read_issues
        )
        raw = values.get(locator)
        if raw in (None, ""):
            issues.append(f"{spec['field_key']}：来源单元格 {locator} 为空")
            continue
        value = Decimal(str(raw)) * Decimal(str(spec.get("scale", 1)))
        fields[spec["field_key"]] = int(value) if value == value.to_integral() else float(value)
        evidence[spec["field_key"]] = _source_evidence(
            source_name, sources, locator, source_lineage
        )

    for spec in config.get("material_fields", []):
        try:
            value, source = resolve_material_field(
                spec,
                sources,
                source_lineage,
            )
        except (KeyError, OSError, ValueError, BadZipFile) as exc:
            value = ""
            source = {
                "kind": "missing",
                "file": "",
                "locator": "",
            }
            issues.append(f"{spec['field_key']}：材料无法读取：{exc}")
        fields[spec["field_key"]] = value
        evidence[spec["field_key"]] = source

    # Project workbooks evolve and frequently rename sheets or move cells.
    # Apply the deterministic semantic reader after fixed project locators so
    # an exact row/column-header match can replace an accidental value read
    # from a legacy coordinate (for example a zero that is no longer the
    # equity-value cell).
    semantic_primary_fields = {
        "book_net_assets",
        "asset_approach_value",
        "income_approach_value",
        "market_approach_value",
        "asset_scope_summary_table",
        "long_term_assets_table",
        "major_long_term_assets",
        "historical_balance_sheet_table",
        "historical_income_statement_table",
    }
    semantic_history_roles: dict[str, str] = {}
    for source_name in (
        "reporting_workbook",
        "audited_financials",
        "income_workbook",
    ):
        source_path = sources.get(source_name)
        if source_path is None or source_path.suffix.lower() not in {".xlsx", ".xlsm"}:
            continue
        try:
            semantic = extract_workbook_facts(source_path, source_name)
        except (KeyError, OSError, ValueError, BadZipFile) as exc:
            issues.append(f"{source_name}：语义定位失败：{exc}")
            continue
        issues.extend(semantic.get("issues", []))
        for rejected_key, rejected_source in semantic.get(
            "evidence",
            {},
        ).items():
            if (
                isinstance(rejected_source, dict)
                and rejected_source.get("kind") == "unfinished_appraisal"
            ):
                fields.pop(rejected_key, None)
                evidence[rejected_key] = rejected_source
        for field_key, value in semantic.get("fields", {}).items():
            if value in (None, "", []):
                continue
            existing = fields.get(field_key)
            existing_source = evidence.get(field_key, {})
            existing_is_valid = (
                existing not in (None, "", [])
                and existing_source.get("kind") != "missing"
            )
            existing_is_semantic = str(
                existing_source.get("kind", "")
            ).startswith("semantic_excel")
            semantic_source = semantic.get("evidence", {}).get(
                field_key,
                {
                    "kind": "semantic_excel",
                    "file": source_path.name,
                    "locator": field_key,
                },
            )
            if field_key in {
                "historical_balance_sheet_table",
                "historical_income_statement_table",
            }:
                existing_history_role = semantic_history_roles.get(field_key)
                if source_name == "audited_financials":
                    fields[field_key] = value
                    evidence[field_key] = semantic_source
                    semantic_history_roles[field_key] = source_name
                    continue
                if existing_history_role == "audited_financials":
                    continue
            if (
                field_key
                in {
                    "historical_balance_sheet_table",
                    "historical_income_statement_table",
                }
                and existing_is_valid
                and existing_is_semantic
                and isinstance(existing, dict)
                and isinstance(value, dict)
            ):
                merged = merge_historical_tables(existing, value)
                if merged != existing:
                    fields[field_key] = merged
                    if merged == value:
                        evidence[field_key] = semantic_source
                        semantic_history_roles[field_key] = source_name
                    else:
                        evidence[field_key] = {
                            "kind": "semantic_excel_merged",
                            "file": "；".join(
                                dict.fromkeys(
                                    filter(
                                        None,
                                        (
                                            str(existing_source.get("file", "")),
                                            str(semantic_source.get("file", "")),
                                        ),
                                    )
                                )
                            ),
                            "locator": "；".join(
                                dict.fromkeys(
                                    filter(
                                        None,
                                        (
                                            str(existing_source.get("locator", "")),
                                            str(semantic_source.get("locator", "")),
                                        ),
                                    )
                                )
                            ),
                        }
                continue
            if existing_is_semantic:
                continue
            if field_key not in semantic_primary_fields and existing_is_valid:
                continue
            selected = prefer_semantic_result(
                fixed_value=existing,
                fixed_evidence=existing_source,
                semantic_value=value,
                semantic_evidence=semantic_source,
            )
            fields[field_key] = selected["value"]
            evidence[field_key] = selected["evidence"]
            if field_key in {
                "historical_balance_sheet_table",
                "historical_income_statement_table",
            }:
                semantic_history_roles[field_key] = source_name

    unfinished_asset_source = evidence.get("asset_approach_value", {})
    result_section_source = evidence.get(
        "asset_approach_result_section",
        {},
    )
    if (
        unfinished_asset_source.get("kind") == "unfinished_appraisal"
        and (
            fields.get("asset_approach_result_section")
            in (None, "", [], {})
            or result_section_source.get("kind") == "missing"
        )
    ):
        evidence["asset_approach_result_section"] = dict(
            unfinished_asset_source
        )

    book_value = fields.get("book_net_assets")
    appraised_value = fields.get("asset_approach_value")
    result_section_is_missing = (
        fields.get("asset_approach_result_section") in (None, "", [])
        or evidence.get("asset_approach_result_section", {}).get("kind") == "missing"
    )
    if (
        result_section_is_missing
        and isinstance(book_value, (int, float))
        and isinstance(appraised_value, (int, float))
    ):
        increase = appraised_value - book_value
        increase_rate = increase / book_value if book_value else 0
        fields["asset_approach_result_section"] = (
            "（二）资产基础法评估结果："
            f"净资产账面价值{book_value:,.2f}万元，"
            f"评估价值{appraised_value:,.2f}万元，"
            f"增值{increase:,.2f}万元，增值率{increase_rate:.2%}。"
        )
        evidence["asset_approach_result_section"] = {
            "kind": "semantic_excel_derived",
            "file": "；".join(
                sorted(
                    {
                        str(evidence.get("book_net_assets", {}).get("file", "")),
                        str(evidence.get("asset_approach_value", {}).get("file", "")),
                    }
                    - {""}
                )
            ),
            "locator": "账面净资产、评估净资产派生",
        }

    if not offline:
        known_fields = set(records_by_key)
        if ocr_adapter is not None and "audit_pdf" in sources:
            payload, provider_issues = ocr_adapter.extract(sources["audit_pdf"])
            issues.extend(provider_issues)
            _merge_provider(payload, kind="ocr", known_fields=known_fields, fields=fields, evidence=evidence)
        if company_api_adapter is not None and fields.get("target_company_name"):
            payload, provider_issues = company_api_adapter.fetch(str(fields["target_company_name"]))
            issues.extend(provider_issues)
            _merge_provider(payload, kind="company_api", known_fields=known_fields, fields=fields, evidence=evidence)
        if llm_adapter is not None:
            payload, provider_issues = llm_adapter.generate(dict(fields))
            issues.extend(provider_issues)
            _merge_provider(payload, kind="llm", known_fields=known_fields, fields=fields, evidence=evidence)

    # The body below the yellow company-profile heading is a normal Word
    # placeholder, not a second LLM route.  Keep it synchronized with the
    # authorized seven-field narrative contract.
    if fields.get("company_profile_section") and not fields.get("company_profile_text"):
        fields["company_profile_text"] = str(fields["company_profile_section"]).rstrip("。；; ")
        evidence["company_profile_text"] = {
            "kind": "llm_profile_alias",
            "file": evidence.get("company_profile_section", {}).get("file", ""),
            "locator": "company_profile_section",
        }

    before_derived = set(fields)
    final_value_field = {
        "收益法": "income_approach_value",
        "市场法": "market_approach_value",
        "资产基础法": "asset_approach_value",
    }.get(
        str(fields.get("final_valuation_method", "")).strip(),
        config.get("final_value_field"),
    )
    fields = derive_system_fields(
        fields,
        report_date or datetime.now().date().isoformat(),
        final_value_field=final_value_field,
    )
    for key in set(fields) - before_derived:
        evidence[key] = {"kind": "system", "file": "", "locator": key}
    required_financial_fields = list(config.get("required_financial_fields", []))
    financial_validation = apply_missing_field_policy(
        fields,
        evidence,
        required_financial_fields,
        "财务材料字段",
    )
    fields = financial_validation["fields"]
    evidence = financial_validation["evidence"]
    issues.extend(financial_validation["issues"])
    for key, record in records_by_key.items():
        if fields.get(key) in (None, "", []):
            fields[key] = ""
            if evidence.get(key, {}).get("kind") != "unfinished_appraisal":
                evidence[key] = {
                    "kind": "missing",
                    "file": "",
                    "locator": "",
                }
            if key not in required_financial_fields:
                issues.append(f"{key}：无可用值，已按规则留空")
    replacements = build_replacements(locations, fields)
    paragraph_replacements: dict[tuple[str, int], str] = {}
    for spec in config.get("paragraph_replacements", []):
        if "field_key" in spec:
            value = str(fields.get(spec["field_key"], ""))
        elif "template" in spec:
            if spec.get("blank_if_empty") and not str(fields.get(spec["blank_if_empty"], "") or "").strip():
                value = ""
            else:
                values = defaultdict(
                    lambda: "XXX",
                    {key: str(value) for key, value in fields.items()},
                )
                value = spec["template"].format_map(values)
        else:
            value = str(spec.get("value", ""))
        paragraph_replacements[(spec["part"], int(spec["paragraph_index"]))] = value
    run_dir = output_dir.resolve() if output_dir else (base / "../../runs" / config["project_id"] / datetime.now().strftime("%Y%m%d_%H%M%S")).resolve()
    run_dir.mkdir(parents=True, exist_ok=True)
    report = run_dir / "资产评估报告_待复核.docx"
    audit = run_dir / "字段审计清单.xlsx"
    before_hash = hashlib.sha256(template.read_bytes()).hexdigest()
    fill_template(
        template,
        report,
        replacements,
        table_replacements=table_replacements,
        paragraph_replacements=paragraph_replacements,
        replacement_modes={
            item["location_id"]: item.get("replacement_mode", "replace_paragraph")
            for item in config.get("yellow_routes", [])
        },
    )
    replace_report_number_year(report, fields.get("report_number_year"))
    unresolved_findings = highlight_unresolved_placeholders(report)
    word_issues = issues_from_word_findings(
        unresolved_findings,
        [*locations, *static_locations],
        fields,
        evidence,
    )
    existing_issue_keys = {
        (item.get("location_id"), item.get("category"))
        for item in word_issues
    }
    word_issues.extend(
        item
        for item in issues_from_special_evidence(
            [*locations, *static_locations],
            evidence,
        )
        if (item.get("location_id"), item.get("category"))
        not in existing_issue_keys
    )
    generation_issues = apply_page_locations(
        word_issues,
        {},
        {},
    )
    generation_issues = organize_generation_issues(
        generation_issues
    )
    issue_workbook = export_generation_issues(
        run_dir / "生成问题清单.xlsx",
        generation_issues,
    )
    issue_json = write_json(
        run_dir / "生成问题清单.json",
        generation_issues,
    )
    for issue in generation_issues:
        issues.append(
            f"Word页码不可用 {issue.get('location_description', '')}："
            f"{issue.get('problem', '')}"
        )
    if hashlib.sha256(template.read_bytes()).hexdigest() != before_hash:
        raise RuntimeError("模板被意外修改")
    export_audit(audit, [*locations, *static_locations], fields, evidence)
    normalized_evidence = write_json(
        run_dir / "normalized_evidence.json",
        evidence,
    )
    manifest = {
        "project_id": config["project_id"], "template": str(template), "template_sha256": before_hash,
        "mapping_version": "1.0.0", "offline": offline, "replacement_count": len(replacements),
        "financial_validation": {
            "valid": financial_validation["valid"],
            "missing_fields": financial_validation["missing_fields"],
            "conflicts": [],
        },
        "generation_validation": {
            "valid": not generation_issues,
            "unresolved_count": len(generation_issues),
        },
        "outputs": [
            str(report),
            str(audit),
            str(issue_workbook),
            str(issue_json),
            str(normalized_evidence),
        ],
    }
    manifest_path = write_json(run_dir / "run_manifest.json", manifest)
    write_json(run_dir / "issues.json", issues)
    write_json(run_dir / "normalized_fields.json", fields)
    return RunResult(report, audit, manifest_path, issues)


def main(argv: list[str] | None = None) -> int:
    _load_local_env()
    parser = argparse.ArgumentParser(description="资产评估报告业务 Demo")
    parser.add_argument("project")
    parser.add_argument("--output-dir", type=Path)
    parser.add_argument("--offline", action="store_true")
    parser.add_argument("--pdf", type=Path, help="启用端到端 OCR 流程时指定的 PDF")
    parser.add_argument("--template", type=Path, help="可选 Word 模板；始终只读")
    parser.add_argument(
        "--ocr-provider",
        choices=["aliyun", "paddle", "none"],
        help="OCR 提供方；默认读取 APPRAISAL_OCR_PROVIDER，未配置时使用 aliyun",
    )
    parser.add_argument(
        "--ocr-engine",
        dest="ocr_provider",
        choices=["aliyun", "paddle", "none"],
        help=argparse.SUPPRESS,
    )
    parser.add_argument("--use-glm", action="store_true", help="使用百炼模型生成叙述字段并执行三类审核")
    parser.add_argument("--use-qichacha", action="store_true", help="使用配置的企查查兼容 API")
    parser.add_argument("--node-inputs-json", type=Path, help="两个节点输入字段的 JSON 文件")
    parser.add_argument("--commissioning-party-name", help="用户输入：委托方全称")
    parser.add_argument("--commissioning-party-short-name", help="用户输入：委托方简称")
    parser.add_argument("--report-serial", help="用户输入：评估报告编号流水号")
    parser.add_argument("--target-company-name", help="用户输入：被评估单位全称（企查查核验，可选）")
    parser.add_argument("--valuation-purpose-inputs", help="用户输入：评估目的")
    parser.add_argument("--selected-valuation-method", help="用户输入：选用评估方法")
    parser.add_argument("--valuation-subject-type", help="用户输入：评估对象/价值类型")
    parser.add_argument("--transaction-type", help="用户输入：交易类型")
    parser.add_argument("--final-valuation-method", help="用户输入：最终采用的评估方法")
    parser.add_argument("--target-company-short-name", help="用户输入：被评估单位简称")
    parser.add_argument("--report-date")
    args = parser.parse_args(argv)
    pipeline_requested = bool(
        args.pdf
        or args.template
        or args.use_glm
        or args.use_qichacha
        or args.node_inputs_json
        or any(
            value
            for value in (
                args.commissioning_party_name,
                args.commissioning_party_short_name,
                args.report_serial,
                args.target_company_name,
                args.valuation_purpose_inputs,
                args.selected_valuation_method,
                args.valuation_subject_type,
                args.transaction_type,
                args.final_valuation_method,
                args.target_company_short_name,
            )
        )
    )
    if pipeline_requested:
        if args.output_dir is None:
            parser.error("端到端流程必须指定 --output-dir")
        from .adapters.template_pages import LibreOfficeTemplatePageReader
        from .pipeline import run_pipeline

        ocr_adapter = None
        if args.pdf is not None:
            ocr_adapter = _select_cli_ocr_adapter(args.ocr_provider, os.environ)
        llm_adapter = None
        qichacha_adapter = None
        http_client = None
        if args.use_glm or args.use_qichacha:
            try:
                import httpx
            except ImportError as exc:
                parser.error(f"使用外部服务需安装 services 依赖：{exc}")
            http_client = httpx.Client(timeout=120)
        if args.use_glm:
            api_key = os.environ.get("DASHSCOPE_API_KEY", "")
            if not api_key:
                parser.error("--use-glm 需要环境变量 DASHSCOPE_API_KEY")
            from .adapters.llm_factory import build_bailian_adapters

            project_config = json.loads(Path(args.project).read_text(encoding="utf-8"))
            adapters = build_bailian_adapters(
                client=http_client,
                api_key=api_key,
                root=Path(__file__).parent,
                config=project_config,
                env=os.environ,
                base_url=os.environ.get(
                    "APPRAISAL_LLM_BASE_URL",
                    "https://dashscope.aliyuncs.com/compatible-mode/v1",
                ),
            )
            llm_adapter = adapters["narrative"]
        if args.use_qichacha:
            from .adapters.company_api import QichachaApiAdapter

            app_key = os.environ.get("QICHACHA_APP_KEY", "")
            secret_key = os.environ.get("QICHACHA_SECRET_KEY", "")
            if not app_key or not secret_key:
                parser.error("--use-qichacha 需要环境变量 QICHACHA_APP_KEY 和 QICHACHA_SECRET_KEY")
            endpoints = {
                code: os.environ[name]
                for code, name in {
                    "735": "QICHACHA_ENDPOINT_735",
                    "231": "QICHACHA_ENDPOINT_231",
                    "514": "QICHACHA_ENDPOINT_514",
                    "233": "QICHACHA_ENDPOINT_233",
                }.items()
                if os.environ.get(name)
            }
            qichacha_adapter = QichachaApiAdapter(
                http_client,
                app_key,
                secret_key,
                base_url=os.environ.get("QICHACHA_API_BASE_URL", "https://api.qichacha.com"),
                endpoints=endpoints,
            )
        node_inputs = (
            json.loads(args.node_inputs_json.read_text(encoding="utf-8"))
            if args.node_inputs_json
            else {}
        )
        cli_inputs = {
            "commissioning_party_name": args.commissioning_party_name,
            "commissioning_party_short_name": args.commissioning_party_short_name,
            "report_serial": args.report_serial,
            "target_company_name": args.target_company_name,
            "valuation_purpose_inputs": args.valuation_purpose_inputs,
            "selected_valuation_method": args.selected_valuation_method,
            "valuation_subject_type": args.valuation_subject_type,
            "transaction_type": args.transaction_type,
            "final_valuation_method": args.final_valuation_method,
            "target_company_short_name": args.target_company_short_name,
        }
        cli_inputs = {key: value for key, value in cli_inputs.items() if value not in (None, "")}
        node_inputs.update({
            key: value
            for key, value in cli_inputs.items()
            if key in {"selected_valuation_method", "valuation_purpose_inputs"}
        })
        result = run_pipeline(
            project_config=Path(args.project),
            pdf_path=args.pdf,
            output_dir=args.output_dir,
            ocr_adapter=ocr_adapter,
            llm_adapter=llm_adapter,
            qichacha_adapter=qichacha_adapter,
            node_inputs=node_inputs,
            manual_inputs_override=cli_inputs,
            template_path=args.template,
            template_page_reader=LibreOfficeTemplatePageReader(),
            report_date=args.report_date,
            generate_all_narratives=True,
        )
        if result.ocr_workbook_path is not None:
            print(f"OCR Excel：{result.ocr_workbook_path}")
        else:
            print("OCR Excel：未提供 PDF，已跳过")
        print(f"报告：{result.report_path}")
        print(f"审计：{result.audit_path}")
        print(f"留空或复核事项：{len(result.issues)}")
        return 0
    result = run_project(Path(args.project), args.output_dir, args.offline)
    print(f"报告：{result.report_path}")
    print(f"审计：{result.audit_path}")
    print(f"留空字段：{len(result.issues)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
