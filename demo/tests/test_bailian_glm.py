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
    assert request["json"]["model"] == "qwen3.7-flash"
    assert request["json"]["enable_thinking"] is False
    assert request["json"]["response_format"]["type"] == "json_schema"
    assert request["headers"]["Authorization"] == "Bearer test-key"


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
