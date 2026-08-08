import json

from demo.adapters.bailian_glm import BailianYellowNarrativeAdapter


class FakeResponse:
    def __init__(self, payload):
        self.payload = payload

    def raise_for_status(self):
        return None

    def json(self):
        return self.payload


class FakeClient:
    def __init__(self, content):
        self.content = content
        self.request = None

    def post(self, *args, **kwargs):
        self.request = {"args": args, "kwargs": kwargs}
        return FakeResponse({"choices": [{"message": {"content": self.content}}]})


class FieldwiseClient:
    def __init__(self):
        self.requests = []

    def post(self, *args, **kwargs):
        self.requests.append(kwargs)
        user_payload = json.loads(kwargs["json"]["messages"][-1]["content"])
        field = user_payload["requested_field"]
        evidence_id = user_payload["evidence"][0]["evidence_id"]
        return FakeResponse(
            {
                "choices": [
                    {
                        "message": {
                            "content": json.dumps(
                                {
                                    field: {
                                        "value": f"{field}内容",
                                        "evidence_ids": [evidence_id],
                                    }
                                },
                                ensure_ascii=False,
                            )
                        }
                    }
                ]
            }
        )


def test_glm_accepts_only_seven_fields_and_known_evidence_ids():
    client = FakeClient(
        json.dumps(
            {
                "fields": {
                    "company_profile_section": {
                        "value": "示例概述",
                        "evidence_ids": ["pdf:p1:b1"],
                    },
                    "industry_overview": {
                        "value": "无依据行业内容",
                        "evidence_ids": ["pdf:p99:b9"],
                    },
                    "book_net_assets": {"value": "999", "evidence_ids": []},
                }
            },
            ensure_ascii=False,
        )
    )
    adapter = BailianYellowNarrativeAdapter(client=client, api_key="test-key", prompt="规则")

    values, issues = adapter.generate(
        {"evidence": [{"evidence_id": "pdf:p1:b1", "text": "示例材料"}]}
    )

    assert values == {"company_profile_section": "示例概述"}
    assert any("book_net_assets" in issue for issue in issues)
    assert any("pdf:p99:b9" in issue for issue in issues)
    request = client.request["kwargs"]
    assert request["json"]["model"] == "qwen3.7-max-2026-05-17"
    # qwen3.7-max-2026-05-17 is a thinking-only snapshot; DashScope rejects
    # an explicit ``enable_thinking=false`` for this model.
    assert "enable_thinking" not in request["json"]
    assert request["json"]["response_format"]["type"] == "json_object"
    assert request["headers"]["Authorization"] == "Bearer test-key"


def test_glm_accepts_flat_json_object_used_by_qwen_flash():
    client = FakeClient(
        json.dumps(
            {
                "company_profile_section": {
                    "value": "",
                    "evidence_ids": [],
                },
                "main_products": {
                    "value": "主营工业滤波器。",
                    "evidence_ids": ["document:reference_report:p10"],
                },
            },
            ensure_ascii=False,
        )
    )
    adapter = BailianYellowNarrativeAdapter(
        client=client,
        api_key="test-key",
        prompt="规则",
    )

    values, issues = adapter.generate(
        {
            "evidence": [
                {
                    "evidence_id": "document:reference_report:p10",
                    "text": "主营工业滤波器。",
                }
            ]
        }
    )

    assert values["main_products"] == "主营工业滤波器。"
    assert issues == []


def test_hybrid_model_explicitly_disables_thinking_for_short_narrative_calls():
    client = FakeClient(
        json.dumps(
            {
                "company_profile_section": {
                    "value": "示例概述",
                    "evidence_ids": ["pdf:p1:b1"],
                }
            },
            ensure_ascii=False,
        )
    )
    adapter = BailianYellowNarrativeAdapter(
        client=client,
        api_key="test-key",
        prompt="规则",
        model="qwen3.7-max-2026-05-20",
    )

    adapter.generate({"evidence": [{"evidence_id": "pdf:p1:b1", "text": "示例材料"}]})

    assert client.request["kwargs"]["json"]["enable_thinking"] is False


def test_glm_generates_selected_modules_field_by_field():
    client = FieldwiseClient()
    adapter = BailianYellowNarrativeAdapter(
        client=client,
        api_key="test-key",
        prompt="规则",
    )

    values, issues = adapter.generate(
        {
            "selected_modules": ["main_products"],
            "evidence": [
                {
                    "evidence_id": "document:reference_report:p10",
                    "text": "示例有限公司主要生产工业滤波器。",
                }
            ],
        }
    )

    assert values == {
        "company_profile_section": "company_profile_section内容",
        "main_products": "main_products内容",
    }
    assert issues == []
    assert len(client.requests) == 2


def test_glm_generates_all_seven_fixed_word_candidates():
    client = FieldwiseClient()
    adapter = BailianYellowNarrativeAdapter(
        client=client,
        api_key="test-key",
        prompt="规则",
    )

    values, issues = adapter.generate(
        {
            "selected_modules": [
                "industry_overview",
                "business_and_segments",
                "main_products",
                "customers_suppliers",
                "profit_model_swot",
                "comparable_list",
            ],
            "evidence": [
                {
                    "evidence_id": "api:qichacha:target:735:profile",
                    "text": "示例有限公司经营工业设备业务。",
                },
                {
                    "evidence_id": "api:qichacha:target:915:peer:工业设备",
                    "text": "上市公司公告候选：上市公司甲；股票代码：600001。",
                },
            ],
        }
    )

    assert issues == []
    assert set(values) == {
        "company_profile_section",
        "industry_overview",
        "business_and_segments",
        "main_products",
        "customers_suppliers",
        "profit_model_swot",
        "comparable_list",
    }
    assert len(client.requests) == 7


def test_glm_failure_returns_empty_fields_without_exposing_api_key():
    class BrokenClient:
        def post(self, *args, **kwargs):
            raise RuntimeError("request failed")

    values, issues = BailianYellowNarrativeAdapter(
        client=BrokenClient(), api_key="secret-never-print", prompt="规则"
    ).generate({"evidence": []})

    assert values == {}
    assert issues == ["百炼 GLM 失败：request failed"]
    assert "secret-never-print" not in issues[0]
