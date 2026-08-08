from __future__ import annotations

from pathlib import Path
from typing import Any, Mapping

from demo.domain.llm_config import resolve_llm_fallback_model, resolve_llm_models

from .bailian_glm import BailianYellowNarrativeAdapter


def build_bailian_adapters(
    *,
    client: Any,
    api_key: str,
    root: Path,
    config: Mapping[str, Any],
    env: Mapping[str, str] | None = None,
    base_url: str = "https://dashscope.aliyuncs.com/compatible-mode/v1",
) -> dict[str, Any]:
    prompt_dir = root / "prompts"
    models = resolve_llm_models(config, env or {})
    fallback_model = resolve_llm_fallback_model(config, env or {})
    narrative_prompt = (prompt_dir / "yellow_narratives.v2.txt").read_text(encoding="utf-8")
    return {
        "narrative": BailianYellowNarrativeAdapter(
            client,
            api_key,
            narrative_prompt,
            base_url=base_url,
            model=models["narrative"],
            fallback_model=fallback_model,
        ),
        "models": models,
    }
