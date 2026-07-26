from __future__ import annotations

from typing import Any, Mapping


TASKS = ("narrative", "format_review", "data_validation", "semantic_review")


def resolve_llm_models(config: Mapping[str, Any] | None, env: Mapping[str, str] | None) -> dict[str, str]:
    """Resolve one model per LLM task, with environment overrides."""
    config = config or {}
    env = env or {}
    llm = config.get("llm", {}) if isinstance(config, Mapping) else {}
    configured_default = llm.get("default_model", "qwen3.7-flash") if isinstance(llm, Mapping) else "qwen3.7-flash"
    default = env.get("APPRAISAL_LLM_MODEL") or configured_default or "qwen3.7-flash"
    task_config = llm.get("tasks", {}) if isinstance(llm, Mapping) else {}
    result: dict[str, str] = {}
    for task in TASKS:
        env_key = f"APPRAISAL_LLM_MODEL_{task.upper()}"
        result[task] = env.get(env_key) or (task_config.get(task) if isinstance(task_config, Mapping) else None) or default
    return result
