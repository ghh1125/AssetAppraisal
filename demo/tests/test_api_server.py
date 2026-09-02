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
    assert api_server._fill_progress_percent(15) >= 72
    assert api_server._fill_progress_percent(68) > api_server._fill_progress_percent(15)
    assert api_server._fill_progress_percent(100) == 96


def test_public_error_hides_provider_details_but_keeps_next_action():
    message = api_server._public_error(RuntimeError("阿里云 OCR 任务失败（id-123，403）：secret detail"))
    assert "secret" not in message
    assert "扫描件" in message


def test_finishing_candidate_node_closes_steps_replayed_during_fill():
    api_server.JOBS.clear()
    api_server._set_job("fill-progress-test", nodes=api_server._initial_node_states())
    api_server._set_step(
        "fill-progress-test",
        "ocr_llm_candidates",
        "query_qichacha",
        "running",
        "企查查 API 查询完成，正在整理返回证据",
    )
    api_server._set_node(
        "fill-progress-test",
        "ocr_llm_candidates",
        "completed",
        "候选内容已确认，材料复核完成",
    )
    node = next(item for item in api_server.JOBS["fill-progress-test"]["nodes"] if item["key"] == "ocr_llm_candidates")
    assert node["status"] == "completed"
    assert all(step["status"] != "running" for step in node["steps"])


def test_update_candidate_payload_changes_only_requested_module(tmp_path):
    path = tmp_path / "llm候选内容.json"
    path.write_text(
        json.dumps(
            {
                "candidates": [
                    {"field_key": "main_products", "value": "旧产品"},
                    {"field_key": "industry_overview", "value": "行业"},
                ]
            },
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    payload = api_server._update_candidate_payload(path, "main_products", "新产品")

    assert payload["candidates"][0]["value"] == "新产品"
    assert payload["candidates"][1]["value"] == "行业"


def test_regenerate_candidate_endpoint_queues_single_module(monkeypatch, tmp_path):
    api_server.JOBS.clear()
    api_server.RUNS_ROOT = tmp_path
    run_id = "candidate-regenerate"
    api_server.JOBS[run_id] = {
        "run_id": run_id,
        "status": "awaiting_selection",
        "candidates": [{"field_key": "main_products", "value": "旧产品"}],
        "selection_context": {"use_glm": True},
    }
    captured = {}

    def fake_regenerate(run_id_arg, field_key, feedback):
        captured.update(run_id=run_id_arg, field_key=field_key, feedback=feedback)

    monkeypatch.setattr(api_server, "_execute_regenerate_candidate", fake_regenerate)
    response = TestClient(api_server.app).post(
        f"/api/v1/asset-appraisal/runs/{run_id}/candidates/main_products/regenerate",
        data={"feedback": "补充产品应用场景"},
    )

    assert response.status_code == 202
    assert captured == {
        "run_id": run_id,
        "field_key": "main_products",
        "feedback": "补充产品应用场景",
    }


def test_edit_candidate_endpoint_persists_user_text(monkeypatch, tmp_path):
    api_server.JOBS.clear()
    api_server.RUNS_ROOT = tmp_path
    run_id = "candidate-edit"
    run_dir = tmp_path / run_id
    run_dir.mkdir()
    (run_dir / "llm候选内容.json").write_text(
        json.dumps({"candidates": [{"field_key": "main_products", "value": "旧产品"}]}, ensure_ascii=False),
        encoding="utf-8",
    )
    api_server.JOBS[run_id] = {
        "run_id": run_id,
        "status": "awaiting_selection",
        "candidates": [{"field_key": "main_products", "value": "旧产品"}],
    }

    response = TestClient(api_server.app).post(
        f"/api/v1/asset-appraisal/runs/{run_id}/candidates/main_products/edit",
        data={"value": "人工修改后的产品描述"},
    )

    assert response.status_code == 200
    assert response.json()["candidates"][0]["value"] == "人工修改后的产品描述"


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


def test_api_accepts_scanned_image_as_audit_material(monkeypatch, tmp_path):
    monkeypatch.setattr(api_server, "RUNS_ROOT", tmp_path)
    captured = {}
    monkeypatch.setattr(
        api_server,
        "_execute_run",
        lambda _run_id, source_path, *_args: captured.update(source_path=source_path),
    )

    response = TestClient(api_server.app).post(
        "/api/v1/asset-appraisal/runs",
        data={"inputs": json.dumps(confirmed_inputs())},
        files=[("audit_materials", ("营业执照扫描件.png", b"png", "image/png"))],
    )

    assert response.status_code == 202
    assert captured["source_path"].suffix.lower() == ".png"


def test_api_accepts_direct_workbook_without_archive_or_audit_material(monkeypatch, tmp_path):
    monkeypatch.setattr(api_server, "RUNS_ROOT", tmp_path)
    captured = {}
    monkeypatch.setattr(
        api_server,
        "_execute_run",
        lambda _run_id, pdf_path, source_overrides, *_args: captured.update(
            pdf_path=pdf_path,
            source_overrides=source_overrides,
        ),
    )

    response = TestClient(api_server.app).post(
        "/api/v1/asset-appraisal/runs",
        data={"inputs": json.dumps(confirmed_inputs())},
        files=[("income_workbook", ("收益法.xlsx", b"reviewed", "application/octet-stream"))],
    )

    assert response.status_code == 202
    assert captured["pdf_path"] is None
    assert captured["source_overrides"]["income_workbook"].read_bytes() == b"reviewed"


def test_completed_intake_supplies_both_workbooks_and_user_upload_overrides_one(monkeypatch, tmp_path):
    monkeypatch.setattr(api_server, "RUNS_ROOT", tmp_path)
    api_server.WORKBOOK_INTAKES.clear()
    intake_id = "intake-001"
    intake_dir = tmp_path / "_workbook_intakes" / intake_id
    intake_dir.mkdir(parents=True)
    asset = intake_dir / "资产法.xlsx"
    income = intake_dir / "收益法.xlsx"
    audit = intake_dir / "审计报告.pdf"
    asset.write_bytes(b"generated-asset")
    income.write_bytes(b"generated-income")
    audit.write_bytes(b"%PDF")
    api_server._set_workbook_intake(
        intake_id,
        status="completed",
        target_company_name="示例有限公司",
        reporting_workbook=str(asset),
        income_workbook=str(income),
        selected_source_file=str(audit),
        supporting_sources={},
        artifacts=[],
    )
    captured = {}
    monkeypatch.setattr(
        api_server,
        "_execute_run",
        lambda _run_id, pdf_path, source_overrides, *_args: captured.update(
            pdf_path=pdf_path,
            source_overrides=source_overrides,
        ),
    )

    response = TestClient(api_server.app).post(
        "/api/v1/asset-appraisal/runs",
        data={
            "inputs": json.dumps(confirmed_inputs()),
            "workbook_intake_id": intake_id,
        },
        files=[("income_workbook", ("人工修改收益法.xlsx", b"reviewed-income", "application/octet-stream"))],
    )

    assert response.status_code == 202
    assert captured["source_overrides"]["reporting_workbook"].read_bytes() == b"generated-asset"
    assert captured["source_overrides"]["income_workbook"].read_bytes() == b"reviewed-income"
    job = response.json()
    assert job["workbook_sources"] == {
        "reporting_workbook": "generated_intake",
        "income_workbook": "user_upload",
    }


def test_workbook_intake_upload_creates_optional_preprocessing_job(monkeypatch, tmp_path):
    monkeypatch.setattr(api_server, "RUNS_ROOT", tmp_path)
    api_server.WORKBOOK_INTAKES.clear()
    captured = {}
    monkeypatch.setattr(
        api_server,
        "_execute_workbook_intake",
        lambda intake_id, archive_path, target: captured.update(
            intake_id=intake_id,
            archive_path=archive_path,
            target=target,
        ),
    )

    response = TestClient(api_server.app).post(
        "/api/v1/asset-appraisal/workbook-intakes",
        data={"target_company_name": "示例有限公司"},
        files=[("archive", ("材料包.zip", b"zip-content", "application/zip"))],
    )

    assert response.status_code == 202
    assert response.json()["status"] == "queued"
    assert len(response.json()["steps"]) >= 10
    assert captured["target"] == "示例有限公司"
    assert captured["archive_path"].read_bytes() == b"zip-content"


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


def test_fill_node_passes_llm_to_evidence_review_without_regenerating_candidates(monkeypatch, tmp_path):
    monkeypatch.setattr(api_server, "RUNS_ROOT", tmp_path)
    api_server.JOBS.clear()
    api_server.JOBS["run-1"] = {
        "run_id": "run-1",
        "status": "awaiting_selection",
        "selection_context": {
            "pdf_path": "",
            "source_overrides": {},
            "inputs": {},
            "use_glm": True,
            "use_qichacha": False,
        },
    }
    llm = object()
    captured = {}
    monkeypatch.setattr(api_server, "_load_local_env", lambda: None)
    monkeypatch.setattr(
        api_server,
        "_build_external_adapters",
        lambda use_glm, use_qichacha: (llm, None, None),
    )
    monkeypatch.setattr(
        pipeline_module,
        "run_pipeline",
        lambda **kwargs: captured.update(kwargs) or object(),
    )

    api_server._execute_fill("run-1", {"main_products": "已确认候选"})

    assert captured["llm_adapter"] is llm
    assert captured["word_comment_locator_adapter"] is llm
    assert captured["llm_values_override"] == {"main_products": "已确认候选"}


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
