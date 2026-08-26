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
    assert "yellow_narratives.v3" in changelog

    providers = manifest["providers"]
    assert {item["name"] for item in providers} >= {"qichacha", "bailian_glm"}
    qcc = next(item for item in providers if item["name"] == "qichacha")
    assert set(qcc["api_codes"]) == {"735", "231", "514", "233"}
    glm = next(item for item in providers if item["name"] == "bailian_glm")
    assert glm["model"] == "deepseek-v4-pro-0813"


def test_prompt_schema_is_the_runtime_seven_field_contract():
    schema = json.loads(Path("demo/prompts/yellow_narratives_output.v3.json").read_text(encoding="utf-8"))
    assert schema["version"] == "yellow_narratives_output.v3"
    assert set(schema["properties"]["fields"]["properties"]) == set(ALLOWED_FIELDS)
    assert set(schema["properties"]["fields"]["required"]) == set(ALLOWED_FIELDS)
    assert schema["properties"]["fields"]["additionalProperties"] is False
    assert schema == OUTPUT_SCHEMA


def test_evidence_review_prompt_has_a_closed_status_contract():
    schema = json.loads(Path("demo/prompts/evidence_review_output.v1.json").read_text(encoding="utf-8"))
    assert schema["version"] == "evidence_review_output.v1"
    assert set(schema["properties"]["reviews"]["items"]["properties"]["status"]["enum"]) == {
        "accept", "needs_review", "conflict", "missing"
    }


def test_traceability_comment_prompt_has_a_closed_comment_contract():
    schema = json.loads(
        Path("demo/prompts/traceability_comments_output.v1.json").read_text(encoding="utf-8")
    )
    assert schema["version"] == "traceability_comments_output.v1"
    assert set(schema["properties"]["comments"]["items"]["required"]) == {
        "comment_id", "comment"
    }
    prompt = Path("demo/prompts/traceability_comments.v1.txt").read_text(encoding="utf-8")
    assert "不得修改、计算、猜测" in prompt
    assert "required_title" in prompt


def test_prompt_compatibility_schema_uses_the_same_field_names():
    schema = json.loads(Path("demo/prompts/narrative_output.v1.json").read_text(encoding="utf-8"))
    assert set(schema["properties"]["fields"]["properties"]) == set(ALLOWED_FIELDS)


def test_demo_has_required_handoff_structure():
    required = ["README.md", "workflow.yaml", "schemas.py", "domain", "prompts", "fixtures", "expected", "tests", "data_manifest.yaml", "CHANGELOG.md"]
    assert all((Path("demo") / name).exists() for name in required)


def test_readme_documents_all_outputs_route_counts_and_failure_policy():
    readme = Path("demo/README.md").read_text(encoding="utf-8")
    assert "资产评估报告_待复核.docx" in readme
    for removed_output in ("字段审计清单.xlsx", "生成问题清单.xlsx", "issues.json"):
        assert removed_output not in readme
    for route_label in ("百炼叙述", "企查查 API", "PDF OCR/XLSX", "节点输入", "unresolved_manual"):
        assert route_label in readme
    assert "Python 3.11" in readme
    assert "DASHSCOPE_API_KEY" in readme
    assert "指定来源无值时保留并标黄" in readme
