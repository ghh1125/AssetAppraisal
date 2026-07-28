from __future__ import annotations

from pathlib import Path

import pytest

from demo.adapters.aliyun_docmind_ocr import AliyunDocMindOcrAdapter
from demo.adapters.ocr_factory import create_ocr_adapter


class FakeAliyunClient:
    pass


def test_aliyun_is_default_and_does_not_load_paddle():
    captured = {}

    def aliyun_factory(**kwargs):
        captured.update(kwargs)
        return FakeAliyunClient()

    adapter = create_ocr_adapter(
        {
            "ALIBABA_CLOUD_ACCESS_KEY_ID": "id",
            "ALIBABA_CLOUD_ACCESS_KEY_SECRET": "secret",
        },
        aliyun_client_factory=aliyun_factory,
        paddle_factory=lambda: pytest.fail("Paddle must not load"),
    )

    assert isinstance(adapter, AliyunDocMindOcrAdapter)
    assert captured["access_key_id"] == "id"
    assert captured["access_key_secret"] == "secret"
    assert adapter.vlm is False


def test_missing_aliyun_credentials_returns_issue_adapter():
    pages, issues = create_ocr_adapter({}).extract(Path("audit.pdf"))

    assert pages == []
    assert issues == ["阿里云 OCR 凭证缺失：请配置 AccessKey ID 和 AccessKey Secret"]


def test_paddle_requires_explicit_provider():
    marker = object()

    adapter = create_ocr_adapter(
        {"APPRAISAL_OCR_PROVIDER": "paddle"},
        paddle_factory=lambda: marker,
    )

    assert adapter is marker


def test_none_provider_skips_ocr():
    assert create_ocr_adapter({"APPRAISAL_OCR_PROVIDER": "none"}) is None


def test_aliyun_vlm_and_timeout_are_configurable():
    adapter = create_ocr_adapter(
        {
            "APPRAISAL_OCR_PROVIDER": "aliyun",
            "ALIBABA_CLOUD_ACCESS_KEY_ID": "id",
            "ALIBABA_CLOUD_ACCESS_KEY_SECRET": "secret",
            "APPRAISAL_OCR_VLM": "true",
            "APPRAISAL_OCR_TIMEOUT_SECONDS": "120",
        },
        aliyun_client_factory=lambda **_: FakeAliyunClient(),
    )

    assert adapter.vlm is True
    assert adapter.timeout_seconds == 120


def test_unknown_provider_returns_review_issue():
    pages, issues = create_ocr_adapter(
        {"APPRAISAL_OCR_PROVIDER": "unexpected"}
    ).extract(Path("audit.pdf"))

    assert pages == []
    assert issues == ["未知 OCR 提供方：unexpected"]
