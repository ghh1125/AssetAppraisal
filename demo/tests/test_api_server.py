import json

from fastapi.testclient import TestClient

import demo.api_server as api_server


def test_api_accepts_manual_only_run(monkeypatch, tmp_path):
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
                {"target_company_name": "示例有限公司"}
            )
        },
    )

    assert response.status_code == 202
    assert [node["key"] for node in response.json()["nodes"]] == [
        "start_input", "ocr_llm_candidates", "fill_word", "output"
    ]


def test_api_accepts_xlsm_income_workbook(monkeypatch, tmp_path):
    monkeypatch.setattr(api_server, "RUNS_ROOT", tmp_path)
    monkeypatch.setattr(
        api_server,
        "_execute_run",
        lambda *args, **kwargs: None,
    )

    response = TestClient(api_server.app).post(
        "/api/v1/asset-appraisal/runs",
        data={"inputs": "{}"},
        files={
            "income_workbook": (
                "收益法.xlsm",
                b"placeholder",
                "application/vnd.ms-excel.sheet.macroEnabled.12",
            )
        },
    )

    assert response.status_code == 202


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
