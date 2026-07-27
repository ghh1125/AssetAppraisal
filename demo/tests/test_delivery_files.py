import json
from pathlib import Path

from demo.adapters.bailian_glm import ALLOWED_FIELDS, OUTPUT_SCHEMA


def test_manifest_and_changelog_cover_handoff_requirements():
    manifest = json.loads(Path("demo/data_manifest.yaml").read_text(encoding="utf-8"))
    assert len(manifest["sources"]) >= 4
    for source in manifest["sources"]:
        assert all(source.get(key) for key in ("name", "purpose", "version", "update_method", "missing_behavior"))
    changelog = Path("demo/CHANGELOG.md").read_text(encoding="utf-8")
    assert "147" in changelog and "company_narrative.v1" in changelog
    assert "yellow_routes.v1" in changelog
    assert "financial_aliases.v1" in changelog
    assert "yellow_narratives.v2" in changelog

    providers = manifest["providers"]
    assert {item["name"] for item in providers} >= {"qichacha", "bailian_glm"}
    qcc = next(item for item in providers if item["name"] == "qichacha")
    assert set(qcc["api_codes"]) == {"735", "231", "514", "233"}
    glm = next(item for item in providers if item["name"] == "bailian_glm")
    assert glm["model"] == "qwen3.7-flash"


def test_prompt_schema_is_the_runtime_seven_field_contract():
    schema = json.loads(Path("demo/prompts/yellow_narratives_output.v2.json").read_text(encoding="utf-8"))
    assert schema["version"] == "yellow_narratives_output.v2"
    assert set(schema["properties"]["fields"]["properties"]) == set(ALLOWED_FIELDS)
    assert set(schema["properties"]["fields"]["required"]) == set(ALLOWED_FIELDS)
    assert schema["properties"]["fields"]["additionalProperties"] is False
    assert schema == OUTPUT_SCHEMA


def test_prompt_compatibility_schema_uses_the_same_field_names():
    schema = json.loads(Path("demo/prompts/narrative_output.v1.json").read_text(encoding="utf-8"))
    assert set(schema["properties"]["fields"]["properties"]) == set(ALLOWED_FIELDS)


def test_demo_has_required_handoff_structure():
    required = ["README.md", "workflow.yaml", "schemas.py", "domain", "prompts", "fixtures", "expected", "tests", "data_manifest.yaml", "CHANGELOG.md"]
    assert all((Path("demo") / name).exists() for name in required)


def test_readme_documents_all_outputs_route_counts_and_failure_policy():
    readme = Path("demo/README.md").read_text(encoding="utf-8")
    for output in (
        "OCR结构化结果.xlsx",
        "资产评估报告_待复核.docx",
        "字段审计清单.xlsx",
        "normalized_fields.json",
        "issues.json",
        "run_manifest.json",
    ):
        assert output in readme
    for route_count in ("百炼叙述：7", "企查查 API：5", "PDF OCR/XLSX：6", "节点输入：2"):
        assert route_count in readme
    assert "Python 3.11" in readme
    assert "DASHSCOPE_API_KEY" in readme
    assert "指定来源无值时留空" in readme
