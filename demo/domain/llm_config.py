from __future__ import annotations

from typing import Any, Mapping


TASKS = ("narrative",)
DEFAULT_LLM_MODEL = "deepseek-v4-flash-0731"
DEFAULT_LLM_FALLBACK_MODEL = "qwen3.8-max"


def resolve_llm_models(config: Mapping[str, Any] | None, env: Mapping[str, str] | None) -> dict[str, str]:
    """Resolve one model per LLM task, with environment overrides."""
    config = config or {}
    env = env or {}
    llm = config.get("llm", {}) if isinstance(config, Mapping) else {}
    configured_default = (
        llm.get("default_model", DEFAULT_LLM_MODEL)
        if isinstance(llm, Mapping)
        else DEFAULT_LLM_MODEL
    )
    default = env.get("APPRAISAL_LLM_MODEL") or configured_default or DEFAULT_LLM_MODEL
    task_config = llm.get("tasks", {}) if isinstance(llm, Mapping) else {}
    result: dict[str, str] = {}
    for task in TASKS:
        env_key = f"APPRAISAL_LLM_MODEL_{task.upper()}"
        result[task] = env.get(env_key) or (task_config.get(task) if isinstance(task_config, Mapping) else None) or default
    return result


def resolve_llm_fallback_model(
    config: Mapping[str, Any] | None,
    env: Mapping[str, str] | None,
) -> str:
    """Resolve the single safety-net model used after a provider/model failure."""
    config = config or {}
    env = env or {}
    llm = config.get("llm", {}) if isinstance(config, Mapping) else {}
    configured = llm.get("fallback_model") if isinstance(llm, Mapping) else None
    return env.get("APPRAISAL_LLM_FALLBACK_MODEL") or configured or DEFAULT_LLM_FALLBACK_MODEL
