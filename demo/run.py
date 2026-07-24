from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from typing import Any

from .adapters.audit import export_audit, write_json
from .adapters.excel import read_cells, read_configured_table
from .adapters.materials import resolve_material_field
from .adapters.word import fill_template, replace_report_number_year, unresolved_placeholders
from .domain.mapping import validate_mapping
from .domain.calculations import derive_system_fields
from .domain.field_validation import normalize_report_serial, report_number_year, validate_valuation_subject_type
from .domain.registry import human_fill
from .domain.replacement import build_replacements


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
    return "资产基础法"


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


def _read_long_term_assets_table(config: dict[str, Any], sources: dict[str, Path]) -> list[list[str]]:
    matrix = [["项目", "账面金额（元）", "数量", "现状、特点"]]
    for row in config.get("long_term_assets_table", {}).get("rows", []):
        locator = str(row["locator"])
        value = read_cells(sources[row["source"]], [locator])[locator]
        if isinstance(value, (int, float)):
            value = f"{value:,.2f}"
        matrix.append([
            str(row.get("label", "")),
            str(value if value not in (None, "") else ""),
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
    source_overrides: dict[str, Path] | None = None,
) -> RunResult:
    config_path = config_path.resolve()
    config = json.loads(config_path.read_text(encoding="utf-8"))
    base = config_path.parent
    template = _path(base, config["template"])
    mapping_path = _path(base, config["mapping"])
    manual_path = _path(base, config["manual_inputs"])
    sources = {name: _path(base, value) for name, value in config.get("sources", {}).items()}
    if source_overrides:
        sources.update({name: Path(path).resolve() for name, path in source_overrides.items()})
    source_lineage = config.get("source_lineage", {})
    mapping = json.loads(mapping_path.read_text(encoding="utf-8"))
    locations = validate_mapping(mapping)
    static_locations = mapping.get("static_locations", [])
    manual = json.loads(manual_path.read_text(encoding="utf-8")) if manual_path.exists() else {}
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
        matrix = read_configured_table(sources[source_name], spec)
        key = spec["field_key"]
        fields[key] = {"caption": spec["caption"], "rows": matrix}
        evidence[key] = _source_evidence(
            source_name, sources, spec["source_locator"], source_lineage
        )
        table_replacements[int(spec["target_table_index"])] = matrix

    # The balance-sheet overview is a real table under its lead-in paragraph,
    # not a prose field.  Fill it in every run mode so the CLI cannot leave
    # the template's default numbers behind.
    scope_table = config.get("asset_scope_summary_table")
    if isinstance(scope_table, dict):
        source_name = str(scope_table["source"])
        matrix = read_configured_table(sources[source_name], scope_table)
        key = str(scope_table["field_key"])
        fields[key] = {"caption": scope_table.get("caption", ""), "rows": matrix}
        evidence[key] = _source_evidence(
            source_name, sources, scope_table.get("source_locator", ""), source_lineage
        )
        table_replacements[int(scope_table["target_table_index"])] = matrix

    long_term_table = config.get("long_term_assets_table")
    if isinstance(long_term_table, dict):
        matrix = _read_long_term_assets_table(config, sources)
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
        raw = read_cells(sources[source_name], [locator])[locator]
        if raw in (None, ""):
            issues.append(f"{spec['field_key']}：来源单元格 {locator} 为空")
            continue
        value = Decimal(str(raw)) * Decimal(str(spec.get("scale", 1)))
        fields[spec["field_key"]] = int(value) if value == value.to_integral() else float(value)
        evidence[spec["field_key"]] = _source_evidence(
            source_name, sources, locator, source_lineage
        )

    for spec in config.get("material_fields", []):
        value, source = resolve_material_field(spec, sources, source_lineage)
        fields[spec["field_key"]] = value
        evidence[spec["field_key"]] = source

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
    fields = derive_system_fields(
        fields,
        report_date or datetime.now().date().isoformat(),
        final_value_field=config.get("final_value_field"),
    )
    for key in set(fields) - before_derived:
        evidence[key] = {"kind": "system", "file": "", "locator": key}
    missing_financial = [
        key for key in config.get("required_financial_fields", [])
        if fields.get(key) in (None, "", [])
    ]
    if missing_financial:
        raise ValueError("财务材料字段未能提取：" + "、".join(missing_financial))
    for key, record in records_by_key.items():
        if fields.get(key) in (None, "", []):
            fields[key] = ""
            evidence[key] = {"kind": "blank", "file": "", "locator": ""}
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
                value = spec["template"].format_map({key: str(value) for key, value in fields.items()})
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
    remaining_placeholders = unresolved_placeholders(report)
    if remaining_placeholders:
        raise ValueError("Word 模板仍有未替换占位符：" + "、".join(remaining_placeholders))
    if hashlib.sha256(template.read_bytes()).hexdigest() != before_hash:
        raise RuntimeError("模板被意外修改")
    export_audit(audit, [*locations, *static_locations], fields, evidence)
    manifest = {
        "project_id": config["project_id"], "template": str(template), "template_sha256": before_hash,
        "mapping_version": "1.0.0", "offline": offline, "replacement_count": len(replacements),
        "outputs": [str(report), str(audit)],
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
    parser.add_argument("--ocr-engine", choices=["paddle"], default="paddle")
    parser.add_argument("--use-glm", action="store_true", help="使用百炼 glm-5.2 生成七个叙述字段")
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
    if args.pdf:
        if args.output_dir is None:
            parser.error("端到端 OCR 流程必须指定 --output-dir")
        from .adapters.paddle_ocr import PaddleStructureOcrAdapter, create_local_pipeline
        from .adapters.template_pages import LibreOfficeTemplatePageReader
        from .pipeline import run_pipeline

        ocr_adapter = PaddleStructureOcrAdapter(create_local_pipeline())
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
            from .adapters.bailian_glm import BailianYellowNarrativeAdapter

            api_key = os.environ.get("DASHSCOPE_API_KEY", "")
            if not api_key:
                parser.error("--use-glm 需要环境变量 DASHSCOPE_API_KEY")
            prompt = (Path(__file__).parent / "prompts/yellow_narratives.v1.txt").read_text(encoding="utf-8")
            llm_adapter = BailianYellowNarrativeAdapter(
                http_client,
                api_key,
                prompt,
                base_url=os.environ.get(
                    "APPRAISAL_LLM_BASE_URL",
                    "https://dashscope.aliyuncs.com/compatible-mode/v1",
                ),
                model=os.environ.get("APPRAISAL_LLM_MODEL", "glm-5.2"),
            )
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
        )
        print(f"OCR Excel：{result.ocr_workbook_path}")
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
