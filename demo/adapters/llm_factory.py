from __future__ import annotations

from pathlib import Path
from typing import Any, Mapping

from demo.domain.llm_config import resolve_llm_models

from .bailian_glm import BailianYellowNarrativeAdapter
from .llm_review import BailianReviewAdapter


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
    narrative_prompt = (prompt_dir / "yellow_narratives.v2.txt").read_text(encoding="utf-8")
    reviews = {}
    for task, filename in (
        ("format", "review_format.v1.txt"),
        ("data", "review_data.v1.txt"),
        ("semantic", "review_semantic.v1.txt"),
    ):
        task_key = {"format": "format_review", "data": "data_validation", "semantic": "semantic_review"}[task]
        reviews[task] = BailianReviewAdapter(
            client,
            api_key,
            (prompt_dir / filename).read_text(encoding="utf-8"),
            task=task_key,
            base_url=base_url,
            model=models[task_key],
            prompt_version=f"{filename.removesuffix('.txt')}",
        )
    return {
        "narrative": BailianYellowNarrativeAdapter(
            client,
            api_key,
            narrative_prompt,
            base_url=base_url,
            model=models["narrative"],
        ),
        "reviews": reviews,
        "models": models,
    }
