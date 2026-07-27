from demo.domain.llm_config import resolve_llm_models


def test_llm_models_use_current_project_default_when_no_override_exists():
    models = resolve_llm_models({}, {})

    assert set(models.values()) == {"qwen3.7-max-2026-05-17"}


def test_llm_models_default_to_qwen_and_allow_task_override():
    models = resolve_llm_models(
        {"llm": {"default_model": "qwen3.7-flash", "tasks": {"semantic_review": "qwen-max"}}},
        {},
    )

    assert models["narrative"] == "qwen3.7-flash"
    assert models["format_review"] == "qwen3.7-flash"
    assert models["semantic_review"] == "qwen-max"


def test_llm_environment_override_wins_over_project_config():
    models = resolve_llm_models(
        {"llm": {"default_model": "project-model"}},
        {"APPRAISAL_LLM_MODEL": "env-model", "APPRAISAL_LLM_MODEL_DATA_VALIDATION": "data-model"},
    )

    assert models["narrative"] == "env-model"
    assert models["data_validation"] == "data-model"
