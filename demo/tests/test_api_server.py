import json

from fastapi.testclient import TestClient

import demo.api_server as api_server
import demo.pipeline as pipeline_module


def test_node_progress_has_user_readable_substeps_and_updates_active_step():
    states = api_server._initial_node_states()
    assert all(node["steps"] for node in states)
    api_server.JOBS.clear()
    api_server._set_job("progress-test", nodes=states)
    api_server._set_step(
        "progress-test",
        "ocr_llm_candidates",
        "query_qichacha",
        "running",
        "企查查 API 正在搜索企业信息",
    )
    node = next(item for item in api_server.JOBS["progress-test"]["nodes"] if item["key"] == "ocr_llm_candidates")
    step = next(item for item in node["steps"] if item["key"] == "query_qichacha")
    assert step["status"] == "running"
    assert step["message"] == "企查查 API 正在搜索企业信息"
    assert node["active_step"] == "query_qichacha"
    api_server._set_step(
        "progress-test",
        "ocr_llm_candidates",
        "query_qichacha",
        "completed",
        "企查查 API 查询完成，正在整理返回证据",
    )
    assert step["status"] == "completed"
    assert node["active_step"] == ""


def test_progress_completion_messages_close_the_active_step():
    assert api_server._progress_step_status("企查查 API 查询完成，正在整理返回证据") == "completed"
    assert api_server._progress_step_status("企查查 API 正在搜索企业信息") == "running"


def test_second_node_substeps_follow_the_actual_parse_review_and_candidate_order():
    states = api_server._initial_node_states()
    node = next(item for item in states if item["key"] == "ocr_llm_candidates")
    assert [step["key"] for step in node["steps"]] == [
        "detect_materials",
        "ocr_pdf",
        "parse_excel",
        "reconcile_sources",
        "review_evidence",
        "query_qichacha",
        "generate_candidates",
        "wait_selection",
    ]


def confirmed_inputs(**overrides):
    payload = {
        "commissioning_party_name": "委托方有限公司",
        "commissioning_party_short_name": "委托方",
        "transaction_type": "收购",
        "target_company_name": "示例有限公司",
        "target_company_short_name": "示例",
        "valuation_subject_type": "股东全部权益价值",
        "selected_valuation_method": ["收益法"],
        "final_valuation_method": "收益法",
        "valuation_base_date": "2025-06-30",
        "registry_info_strategy": "file",
        "ownership_history_strategy": "file",
        "unrecorded_intangibles_strategy": "file",
        "company_profile_strategy": "file",
    }
    payload.update(overrides)
    return payload


def test_api_requires_audit_material(monkeypatch, tmp_path):
    monkeypatch.setattr(api_server, "RUNS_ROOT", tmp_path)
    monkeypatch.setattr(
        api_server,
        "_execute_run",
        lambda *args, **kwargs: None,
    )

    response = TestClient(api_server.app).post(
        "/api/v1/asset-appraisal/runs",
        data={
            "inputs": json.dumps(
                confirmed_inputs()
            )
        },
    )

    assert response.status_code == 422
    assert "审计报告材料" in response.json()["detail"]


def test_api_accepts_confirmed_inputs_without_report_serial(monkeypatch, tmp_path):
    """0817 input contract removes report serial and requires a base date."""
    monkeypatch.setattr(api_server, "RUNS_ROOT", tmp_path)
    monkeypatch.setattr(api_server, "_execute_run", lambda *args, **kwargs: None)

    response = TestClient(api_server.app).post(
        "/api/v1/asset-appraisal/runs",
        data={
            "inputs": json.dumps(
                confirmed_inputs()
            )
        },
        files=[
            ("audit_materials", ("审计报告.pdf", b"pdf", "application/pdf")),
            ("audit_materials", ("审计附表.xlsx", b"xlsx", "application/octet-stream")),
        ],
    )

    assert response.status_code == 202
    run_id = response.json()["run_id"]
    assert (tmp_path / run_id / "input" / "audit" / "审计报告.pdf").is_file()
    assert (tmp_path / run_id / "input" / "audit" / "审计附表.xlsx").is_file()


def test_api_accepts_xlsm_income_workbook(monkeypatch, tmp_path):
    monkeypatch.setattr(api_server, "RUNS_ROOT", tmp_path)
    monkeypatch.setattr(
        api_server,
        "_execute_run",
        lambda *args, **kwargs: None,
    )

    response = TestClient(api_server.app).post(
        "/api/v1/asset-appraisal/runs",
        data={"inputs": json.dumps(confirmed_inputs())},
        files=[
            ("audit_materials", ("审计报告.pdf", b"pdf", "application/pdf")),
            ("income_workbook", (
                "收益法.xlsm",
                b"placeholder",
                "application/vnd.ms-excel.sheet.macroEnabled.12",
            )),
        ],
    )

    assert response.status_code == 202


def test_typed_workbook_slots_keep_roles_when_filenames_are_arbitrary(monkeypatch, tmp_path):
    monkeypatch.setattr(api_server, "RUNS_ROOT", tmp_path)
    monkeypatch.setattr(api_server, "_execute_run", lambda *args, **kwargs: None)
    response = TestClient(api_server.app).post(
        "/api/v1/asset-appraisal/runs",
        data={"inputs": json.dumps(confirmed_inputs())},
        files=[
            ("audit_materials", ("审计报告.pdf", b"pdf", "application/pdf")),
            ("reporting_workbook", ("客户自定义名称一.xlsx", b"one", "application/octet-stream")),
            ("income_workbook", ("客户自定义名称二.xlsx", b"two", "application/octet-stream")),
        ],
    )
    assert response.status_code == 202
    run_id = response.json()["run_id"]
    assert (tmp_path / run_id / "input" / "客户自定义名称一.xlsx").is_file()
    assert (tmp_path / run_id / "input" / "客户自定义名称二.xlsx").is_file()


def test_api_disables_the_hidden_legacy_audited_workbook(monkeypatch, tmp_path):
    """Web uploads must be the only financial sources used by a run."""
    monkeypatch.setattr(api_server, "RUNS_ROOT", tmp_path)
    captured = {}
    monkeypatch.setattr(
        api_server,
        "_execute_run",
        lambda _run_id, _pdf, source_overrides, *_args: captured.update(
            source_overrides=source_overrides
        ),
    )
    response = TestClient(api_server.app).post(
        "/api/v1/asset-appraisal/runs",
        data={"inputs": json.dumps(confirmed_inputs())},
        files=[
            ("audit_materials", ("审计报告.pdf", b"pdf", "application/pdf")),
            ("reporting_workbook", ("资产清查.xlsx", b"one", "application/octet-stream")),
            ("income_workbook", ("收益法.xlsx", b"two", "application/octet-stream")),
        ],
    )

    assert response.status_code == 202
    assert captured["source_overrides"]["audited_financials"] is None


def test_api_accepts_generic_multi_file_materials(monkeypatch, tmp_path):
    monkeypatch.setattr(api_server, "RUNS_ROOT", tmp_path)
    monkeypatch.setattr(api_server, "_execute_run", lambda *args, **kwargs: None)
    response = TestClient(api_server.app).post(
        "/api/v1/asset-appraisal/runs",
        data={"inputs": json.dumps(confirmed_inputs())},
        files=[
            ("materials", ("审计报告.pdf", b"pdf", "application/pdf")),
            ("materials", ("资产清查.xlsx", b"xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")),
            ("materials", ("补充说明.docx", b"docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")),
        ],
    )
    assert response.status_code == 202
    run_id = response.json()["run_id"]
    assert (tmp_path / run_id / "input" / "audit" / "审计报告.pdf").is_file()
    assert (tmp_path / run_id / "input" / "资产清查.xlsx").is_file()
    assert (tmp_path / run_id / "input" / "补充说明.docx").is_file()


def test_api_rejects_completely_empty_run(monkeypatch, tmp_path):
    monkeypatch.setattr(api_server, "RUNS_ROOT", tmp_path)

    response = TestClient(api_server.app).post(
        "/api/v1/asset-appraisal/runs",
        data={"inputs": "{}"},
    )

    assert response.status_code == 422


def test_web_selects_configured_ocr_only_for_uncached_pdf(monkeypatch, tmp_path):
    marker = object()
    calls = []
    pdf_path = tmp_path / "audit.pdf"
    pdf_path.write_bytes(b"%PDF")
    monkeypatch.setattr(
        api_server,
        "create_ocr_adapter",
        lambda env: calls.append(dict(env)) or marker,
    )

    selected = api_server._select_ocr_adapter(
        pdf_path,
        None,
        {"APPRAISAL_OCR_PROVIDER": "aliyun"},
    )

    assert selected is marker
    assert calls == [{"APPRAISAL_OCR_PROVIDER": "aliyun"}]


def test_web_does_not_create_ocr_adapter_on_cache_hit(monkeypatch, tmp_path):
    pdf_path = tmp_path / "audit.pdf"
    cache_path = tmp_path / "OCR结构化结果.xlsx"
    monkeypatch.setattr(
        api_server,
        "create_ocr_adapter",
        lambda _env: (_ for _ in ()).throw(
            AssertionError("OCR adapter must not be created")
        ),
    )

    assert api_server._select_ocr_adapter(pdf_path, cache_path, {}) is None
    assert api_server._select_ocr_adapter(None, None, {}) is None


def test_api_selection_submits_only_generated_candidates(monkeypatch, tmp_path):
    monkeypatch.setattr(api_server, "RUNS_ROOT", tmp_path)
    api_server.JOBS.clear()
    api_server.JOBS["run-1"] = {
        "run_id": "run-1",
        "status": "awaiting_selection",
        "candidates": [
            {"field_key": "industry_overview", "value": "行业候选"},
        ],
        "selection_context": {"inputs": {}},
    }
    captured = {}
    monkeypatch.setattr(
        api_server,
        "_execute_fill",
        lambda run_id, selected: captured.update(run_id=run_id, selected=selected),
    )

    response = TestClient(api_server.app).post(
        "/api/v1/asset-appraisal/runs/run-1/select",
        data={
            "selected_fields": json.dumps({
                "industry_overview": "客户端不应覆盖候选",
                "unknown_field": "越权内容",
            })
        },
    )

    assert response.status_code == 202
    assert captured == {"run_id": "run-1", "selected": {"industry_overview": "行业候选"}}


def test_fill_node_keeps_qichacha_client_available_for_missing_snapshot_data(monkeypatch, tmp_path):
    monkeypatch.setattr(api_server, "RUNS_ROOT", tmp_path)
    api_server.JOBS.clear()
    api_server.JOBS["run-1"] = {
        "run_id": "run-1",
        "status": "awaiting_selection",
        "selection_context": {
            "pdf_path": "",
            "source_overrides": {},
            "inputs": {},
            "use_qichacha": True,
        },
    }
    calls = []
    monkeypatch.setattr(api_server, "_load_local_env", lambda: None)
    monkeypatch.setattr(
        api_server,
        "_build_external_adapters",
        lambda use_glm, use_qichacha: calls.append((use_glm, use_qichacha))
        or (None, None, None),
    )
    monkeypatch.setattr(
        pipeline_module,
        "run_pipeline",
        lambda **kwargs: object(),
    )

    api_server._execute_fill("run-1", {})

    assert calls == [(False, True)]


def test_public_artifacts_expose_only_the_report_word(tmp_path):
    (tmp_path / "资产评估报告_待复核.docx").write_bytes(b"docx")
    for name in (
        "OCR结构化结果.xlsx",
        "字段审计清单.xlsx",
        "生成问题清单.xlsx",
        "生成问题清单.json",
        "workflow_trace.json",
    ):
        (tmp_path / name).write_bytes(b"internal")

    assert api_server._artifact_list(tmp_path) == [
        {"name": "资产评估报告_待复核.docx", "label": "评估报告 Word"}
    ]
