from __future__ import annotations

import demo.run as run_module


def test_cli_provider_override_is_forwarded_to_factory(monkeypatch):
    marker = object()
    captured = {}

    def fake_factory(env):
        captured.update(env)
        return marker

    monkeypatch.setattr(run_module, "create_ocr_adapter", fake_factory)

    adapter = run_module._select_cli_ocr_adapter("paddle", {"KEEP": "value"})

    assert adapter is marker
    assert captured["APPRAISAL_OCR_PROVIDER"] == "paddle"
    assert captured["KEEP"] == "value"


def test_cli_without_override_preserves_environment_provider(monkeypatch):
    captured = {}
    monkeypatch.setattr(
        run_module,
        "create_ocr_adapter",
        lambda env: captured.update(env) or object(),
    )

    run_module._select_cli_ocr_adapter(
        None,
        {"APPRAISAL_OCR_PROVIDER": "aliyun"},
    )

    assert captured["APPRAISAL_OCR_PROVIDER"] == "aliyun"
