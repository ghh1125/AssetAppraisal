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
