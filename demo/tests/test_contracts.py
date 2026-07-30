import json
from pathlib import Path

from demo import schemas


def test_workflow_nodes_have_documented_models():
    workflow = json.loads(Path("demo/workflow.yaml").read_text(encoding="utf-8"))
    assert [node["name"] for node in workflow["nodes"]] == [
        "start_input",
        "ocr_llm_candidates",
        "fill_word",
        "output",
    ]
    for node in workflow["nodes"]:
        for model_name in (node["input_model"], node["output_model"]):
            model = getattr(schemas, model_name)
            for field in model.model_fields.values():
                assert field.description and any("\u4e00" <= c <= "\u9fff" for c in field.description)
                assert field.examples


def test_workflow_has_human_review_checkpoint():
    workflow = json.loads(Path("demo/workflow.yaml").read_text(encoding="utf-8"))
    assert workflow["nodes"][1]["human_checkpoint"] == "用户逐项选择要写入Word的LLM候选内容"


def test_schema_exposes_explicit_business_required_metadata():
    schema = schemas.InventoryInput.model_json_schema()
    assert schema["properties"]["template_path"]["x-是否必填"] == "是"
    optional_schema = schemas.ResolvedField.model_json_schema()
    assert optional_schema["properties"]["evidence"]["x-是否必填"] == "否"


def test_manifest_registers_current_contract_and_trace_versions():
    manifest = json.loads(
        Path("demo/data_manifest.yaml").read_text(encoding="utf-8")
    )
    versions = manifest["rule_versions"]

    assert versions["workflow_contract"] == "workflow_contract.v2"
    assert versions["workflow_trace"] == "workflow_trace.v1"
    assert versions["narrative_policy"] == "narrative_policy.v1"
    assert versions["financial_validation"] == "financial_validation.v2"
    assert "review_output_schema" not in versions
    assert versions["optional_sources"] == "optional_sources.v1"
    assert versions["workbook_semantics"] == "workbook_semantics.v3"
    assert versions["generation_issues"] == "generation_issues.v2"
    assert versions["placeholder_policy"] == "placeholder_policy.v1"
