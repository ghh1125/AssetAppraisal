import json
from pathlib import Path

from demo import schemas


def test_workflow_nodes_have_documented_models():
    workflow = json.loads(Path("demo/workflow.yaml").read_text(encoding="utf-8"))
    assert [node["name"] for node in workflow["nodes"]] == [
        "inventory",
        "ocr_pdf",
        "export_ocr_workbook",
        "extract_sources",
        "resolve_fields",
        "select_narrative_modules",
        "generate_narrative",
        "fill_word",
        "llm_format_review",
        "llm_data_validation",
        "llm_semantic_review",
        "review_aggregate",
        "export_audit",
    ]
    for node in workflow["nodes"]:
        for model_name in (node["input_model"], node["output_model"]):
            model = getattr(schemas, model_name)
            for field in model.model_fields.values():
                assert field.description and any("\u4e00" <= c <= "\u9fff" for c in field.description)
                assert field.examples


def test_workflow_has_human_review_checkpoint():
    workflow = json.loads(Path("demo/workflow.yaml").read_text(encoding="utf-8"))
    assert workflow["nodes"][-1]["human_checkpoint"] == "评估师审核生成 Word、字段审计清单和问题清单"


def test_schema_exposes_explicit_business_required_metadata():
    schema = schemas.InventoryInput.model_json_schema()
    assert schema["properties"]["template_path"]["x-是否必填"] == "是"
    optional_schema = schemas.ResolvedField.model_json_schema()
    assert optional_schema["properties"]["evidence"]["x-是否必填"] == "否"
