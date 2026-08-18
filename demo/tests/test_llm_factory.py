from pathlib import Path

from demo.adapters.llm_factory import build_bailian_adapters


def test_bailian_factory_creates_a_narrative_adapter_with_a_separate_review_model():
    adapters = build_bailian_adapters(
        client=object(),
        api_key="key",
        root=Path("demo"),
        config={
            "llm": {
                "default_model": "qwen3.7-max-2026-05-17",
                "tasks": {"evidence_review": "qwen3.8-max"},
            }
        },
    )

    assert adapters["narrative"].model == "qwen3.7-max-2026-05-17"
    assert adapters["narrative"].review_model == "qwen3.8-max"
    assert adapters["narrative"].review_prompt
    assert adapters["narrative"].fallback_model == "qwen3.8-max"
    assert "reviews" not in adapters
