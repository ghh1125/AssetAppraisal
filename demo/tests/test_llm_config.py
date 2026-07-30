from demo.domain.llm_config import resolve_llm_models


def test_llm_models_use_current_project_default_when_no_override_exists():
    models = resolve_llm_models({}, {})

    assert set(models.values()) == {"qwen3.7-max-2026-05-17"}


def test_llm_models_default_to_qwen_and_allow_task_override():
    models = resolve_llm_models(
        {"llm": {"default_model": "qwen3.7-flash", "tasks": {"narrative": "qwen-max"}}},
        {},
    )

    assert models["narrative"] == "qwen-max"
    assert set(models) == {"narrative"}


def test_llm_environment_override_wins_over_project_config():
    models = resolve_llm_models(
        {"llm": {"default_model": "project-model"}},
        {"APPRAISAL_LLM_MODEL": "env-model", "APPRAISAL_LLM_MODEL_NARRATIVE": "narrative-model"},
    )

    assert models["narrative"] == "narrative-model"
