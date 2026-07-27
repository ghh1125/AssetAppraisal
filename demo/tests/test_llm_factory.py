from pathlib import Path

from demo.adapters.llm_factory import build_bailian_adapters


def test_bailian_factory_creates_one_configured_adapter_per_task():
    adapters = build_bailian_adapters(
        client=object(),
        api_key="key",
        root=Path("demo"),
        config={"llm": {"default_model": "qwen3.7-max-2026-05-17"}},
    )

    assert adapters["narrative"].model == "qwen3.7-max-2026-05-17"
    assert adapters["reviews"]["format"].model == "qwen3.7-max-2026-05-17"
    assert adapters["reviews"]["data"].model == "qwen3.7-max-2026-05-17"
    assert adapters["reviews"]["semantic"].model == "qwen3.7-max-2026-05-17"
