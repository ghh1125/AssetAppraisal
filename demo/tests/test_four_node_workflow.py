from __future__ import annotations

import json
from pathlib import Path

from demo.domain.narrative_policy import select_llm_candidates
from demo.domain.workflow_contracts import validate_workflow_contract
from demo import schemas
from demo.pipeline import run_pipeline
from demo.tests.test_pipeline import FixtureLlmAdapter, FixtureOcrAdapter, fixture_ocr_fields


ROOT = Path(__file__).resolve().parents[1]


def test_workflow_is_split_into_four_explicit_nodes() -> None:
    payload = json.loads((ROOT / "workflow.yaml").read_text(encoding="utf-8"))
    result = validate_workflow_contract(payload, schemas)

    assert result["valid"], result["issues"]
    assert [node["name"] for node in payload["nodes"]] == [
        "start_input",
        "ocr_llm_candidates",
        "fill_word",
        "output",
    ]
    assert all(node.get("description") for node in payload["nodes"])
    assert payload["nodes"][1]["human_checkpoint"]


def test_select_llm_candidates_only_keeps_allowed_template_slots() -> None:
    candidates = {
        "company_profile_section": "公司概况候选",
        "industry_overview": "行业候选",
        "main_products": "产品候选",
        "unknown_field": "不能写入模板的内容",
    }

    selected = select_llm_candidates(candidates, ["industry_overview"])

    assert selected == {"industry_overview": "行业候选"}


def test_candidate_node_pauses_before_word_is_written(tmp_path: Path) -> None:
    result = run_pipeline(
        project_config=ROOT / "projects/tongfu.yaml",
        pdf_path=ROOT.parent / "资产评估工作流/通富2025.6.30合并及母公司审计报告.pdf",
        output_dir=tmp_path,
        ocr_adapter=FixtureOcrAdapter(),
        llm_adapter=FixtureLlmAdapter(),
        ocr_field_resolver=fixture_ocr_fields,
        prepare_only=True,
        generate_all_narratives=True,
    )

    assert result.candidate_path is not None and result.candidate_path.exists()
    assert not result.report_path.exists()
    trace = json.loads((tmp_path / "workflow_trace.json").read_text(encoding="utf-8"))
    assert [node["node_name"] for node in trace["nodes"]] == [
        "start_input", "ocr_llm_candidates", "fill_word", "output"
    ]
    assert trace["nodes"][1]["output_data"]["selection_required"] is True
    assert trace["nodes"][2]["status"] == "skipped"


def test_selected_candidates_are_the_only_llm_slots_written(tmp_path: Path) -> None:
    prepared = run_pipeline(
        project_config=ROOT / "projects/tongfu.yaml",
        pdf_path=ROOT.parent / "资产评估工作流/通富2025.6.30合并及母公司审计报告.pdf",
        output_dir=tmp_path / "prepare",
        ocr_adapter=FixtureOcrAdapter(),
        llm_adapter=FixtureLlmAdapter(),
        ocr_field_resolver=fixture_ocr_fields,
        prepare_only=True,
        generate_all_narratives=True,
    )
    chosen_key, chosen_value = next(iter(prepared.candidate_fields.items()))
    final = run_pipeline(
        project_config=ROOT / "projects/tongfu.yaml",
        pdf_path=ROOT.parent / "资产评估工作流/通富2025.6.30合并及母公司审计报告.pdf",
        output_dir=tmp_path / "final",
        ocr_adapter=None,
        ocr_workbook_path=prepared.ocr_workbook_path,
        llm_values_override={chosen_key: chosen_value},
        generate_all_narratives=True,
    )

    assert final.report_path.exists()
    fields = json.loads((tmp_path / "final/normalized_fields.json").read_text(encoding="utf-8"))
    assert fields[chosen_key] == chosen_value
    for key in prepared.candidate_fields:
        if key != chosen_key:
            assert fields.get(key, "") == ""
