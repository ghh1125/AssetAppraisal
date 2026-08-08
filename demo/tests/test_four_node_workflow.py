from __future__ import annotations

import json
from pathlib import Path

from demo.domain.narrative_policy import LLM_TEMPLATE_FIELDS, select_llm_candidates
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


def test_output_node_contract_contains_only_the_word_report() -> None:
    payload = json.loads((ROOT / "workflow.yaml").read_text(encoding="utf-8"))
    output_node = payload["nodes"][-1]

    assert "审计清单" not in output_node["description"]
    assert "问题清单" not in output_node["description"]
    assert set(schemas.OutputInput.model_fields) == {"report_path"}
    assert set(schemas.OutputOutput.model_fields) == {"report_path"}


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


def test_generate_all_candidate_node_requests_every_fixed_llm_slot(tmp_path: Path) -> None:
    class StrictGenerateAllAdapter:
        prompt_version = "yellow_narratives.test"

        def generate(self, evidence):
            assert set(evidence["selected_modules"]) == (
                set(LLM_TEMPLATE_FIELDS) - {"company_profile_section"}
            )
            return {field_key: f"{field_key}候选" for field_key in LLM_TEMPLATE_FIELDS}, []

    result = run_pipeline(
        project_config=ROOT / "projects/tongfu.yaml",
        pdf_path=None,
        output_dir=tmp_path,
        ocr_adapter=None,
        llm_adapter=StrictGenerateAllAdapter(),
        prepare_only=True,
        generate_all_narratives=True,
        manual_inputs_override={
            "commissioning_party_name": "委托方有限公司",
            "target_company_name": "被评估单位有限公司",
        },
    )

    assert set(result.candidate_fields) == set(LLM_TEMPLATE_FIELDS)


def test_candidate_file_always_exposes_the_six_selectable_report_modules(tmp_path: Path) -> None:
    """The UI must show every report module, even when an LLM lacks evidence.

    Company profile is a deterministic Word field and is deliberately not a
    checkbox.  The six user-facing choices must keep their fixed order so the
    reviewer understands exactly which Word sections can be accepted.
    """

    class SparseAdapter:
        prompt_version = "yellow_narratives.test"

        def generate(self, evidence):
            return {
                "company_profile_section": "公司概况自动填充",
                "main_products": "主要产品候选",
            }, []

    result = run_pipeline(
        project_config=ROOT / "projects/tongfu.yaml",
        pdf_path=None,
        output_dir=tmp_path,
        ocr_adapter=None,
        llm_adapter=SparseAdapter(),
        prepare_only=True,
        generate_all_narratives=True,
        manual_inputs_override={
            "commissioning_party_name": "委托方有限公司",
            "target_company_name": "被评估单位有限公司",
        },
    )

    payload = json.loads(result.candidate_path.read_text(encoding="utf-8"))
    assert [item["field_key"] for item in payload["candidates"]] == [
        "industry_overview",
        "business_and_segments",
        "main_products",
        "customers_suppliers",
        "profit_model_swot",
        "comparable_list",
    ]
    assert payload["candidates"][2]["value"] == "主要产品候选"
    assert all(isinstance(item["available"], bool) for item in payload["candidates"])
    assert payload["automatic_fields"] == {
        "company_profile_section": "公司概况自动填充"
    }


def test_qichacha_evidence_ids_are_unique_between_company_roles(tmp_path: Path) -> None:
    class SameIdQichachaAdapter:
        def fetch(self, company_name):
            return {
                "profile": {"name": company_name},
                "fields": {},
                "evidence": [
                    {
                        "evidence_id": "api:qichacha:735:company_profile_section",
                        "api_code": "735",
                        "text": f"企业名称：{company_name}",
                    }
                ],
            }, []

    class EvidenceAuditLlmAdapter:
        prompt_version = "yellow_narratives.test"

        def generate(self, evidence):
            ids = [
                item["evidence_id"]
                for item in evidence["evidence"]
                if item["evidence_id"].startswith("api:qichacha:")
            ]
            assert len(ids) == len(set(ids))
            assert "api:qichacha:commissioning:735:company_profile_section" in ids
            assert "api:qichacha:target:735:company_profile_section" in ids
            return {}, []

    run_pipeline(
        project_config=ROOT / "projects/tongfu.yaml",
        pdf_path=None,
        output_dir=tmp_path,
        ocr_adapter=None,
        llm_adapter=EvidenceAuditLlmAdapter(),
        qichacha_adapter=SameIdQichachaAdapter(),
        prepare_only=True,
        generate_all_narratives=True,
        manual_inputs_override={
            "commissioning_party_name": "委托方有限公司",
            "target_company_name": "被评估单位有限公司",
        },
    )


def test_candidate_node_persists_qichacha_snapshot_for_word_fill(tmp_path: Path) -> None:
    class SnapshotQichachaAdapter:
        def fetch(self, company_name):
            return {
                "profile": {"name": company_name, "credit_code": f"code-{company_name}"},
                "fields": {},
                "evidence": [],
            }, []

    run_pipeline(
        project_config=ROOT / "projects/tongfu.yaml",
        pdf_path=None,
        output_dir=tmp_path,
        ocr_adapter=None,
        qichacha_adapter=SnapshotQichachaAdapter(),
        prepare_only=True,
        generate_all_narratives=True,
        source_overrides={
            "audit_pdf": None,
            "reporting_workbook": None,
            "income_workbook": None,
            "reference_report": None,
        },
        manual_inputs_override={
            "commissioning_party_name": "委托方有限公司",
            "target_company_name": "被评估单位有限公司",
        },
    )

    snapshot_path = tmp_path / "qichacha_result.json"
    assert snapshot_path.is_file()
    snapshot = json.loads(snapshot_path.read_text(encoding="utf-8"))
    assert snapshot["payloads"]["commissioning"]["profile"]["name"] == "委托方有限公司"
    assert snapshot["payloads"]["target"]["profile"]["name"] == "被评估单位有限公司"


def test_word_fill_reuses_qichacha_snapshot_without_new_api_calls(tmp_path: Path) -> None:
    class CountingQichachaAdapter:
        def __init__(self):
            self.calls = []

        def fetch(self, company_name):
            self.calls.append(company_name)
            return {
                "profile": {
                    "name": company_name,
                    "credit_code": f"code-{company_name}",
                    "registered_capital": "1,000万元",
                },
                "fields": {},
                "evidence": [],
            }, []

    adapter = CountingQichachaAdapter()
    run_kwargs = {
        "project_config": ROOT / "projects/tongfu.yaml",
        "pdf_path": None,
        "output_dir": tmp_path,
        "ocr_adapter": None,
        "qichacha_adapter": adapter,
        "generate_all_narratives": True,
        "source_overrides": {
            "audit_pdf": None,
            "reporting_workbook": None,
            "income_workbook": None,
            "reference_report": None,
        },
        "manual_inputs_override": {
            "commissioning_party_name": "委托方有限公司",
            "target_company_name": "被评估单位有限公司",
        },
    }
    run_pipeline(**run_kwargs, prepare_only=True)
    assert adapter.calls == ["委托方有限公司", "被评估单位有限公司"]

    final = run_pipeline(
        **run_kwargs,
        llm_values_override={},
    )

    assert final.report_path.is_file()
    assert adapter.calls == ["委托方有限公司", "被评估单位有限公司"]
    fields = json.loads((tmp_path / "normalized_fields.json").read_text(encoding="utf-8"))
    assert fields["registered_capital"] == "1,000.00万元"


def test_word_fill_retries_only_qichacha_role_missing_from_snapshot(tmp_path: Path) -> None:
    class PartialQichachaAdapter:
        def fetch(self, company_name):
            if company_name == "委托方有限公司":
                return {"profile": {"name": company_name}, "fields": {}}, []
            return {"profile": {"name": "错误主体有限公司"}, "fields": {}}, []

    class MissingRoleQichachaAdapter:
        def __init__(self):
            self.calls = []

        def fetch(self, company_name):
            self.calls.append(company_name)
            return {
                "profile": {
                    "name": company_name,
                    "registered_capital": "2,000万元",
                },
                "fields": {},
            }, []

    common = {
        "project_config": ROOT / "projects/tongfu.yaml",
        "pdf_path": None,
        "output_dir": tmp_path,
        "ocr_adapter": None,
        "generate_all_narratives": True,
        "source_overrides": {
            "audit_pdf": None,
            "reporting_workbook": None,
            "income_workbook": None,
            "reference_report": None,
        },
        "manual_inputs_override": {
            "commissioning_party_name": "委托方有限公司",
            "target_company_name": "被评估单位有限公司",
        },
    }
    run_pipeline(
        **common,
        qichacha_adapter=PartialQichachaAdapter(),
        prepare_only=True,
    )
    fallback = MissingRoleQichachaAdapter()

    final = run_pipeline(
        **common,
        qichacha_adapter=fallback,
        llm_values_override={},
    )

    assert final.report_path.is_file()
    assert fallback.calls == ["被评估单位有限公司"]


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
