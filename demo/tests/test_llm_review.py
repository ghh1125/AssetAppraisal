import json

from demo.adapters.llm_review import BailianReviewAdapter


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


def test_review_adapter_returns_structured_findings_and_uses_task_model():
    client = FakeClient(
        json.dumps(
            {
                "status": "completed",
                "summary": "发现一项数据问题",
                "findings": [
                    {
                        "location": "第9页表格2",
                        "severity": "high",
                        "category": "amount",
                        "problem": "金额不一致",
                        "evidence": "source.xlsx!Sheet1!A1",
                        "suggestion": "人工复核",
                    }
                ],
            },
            ensure_ascii=False,
        )
    )
    adapter = BailianReviewAdapter(
        client=client,
        api_key="test-key",
        prompt="审核规则",
        task="data_validation",
        model="qwen3.7-flash",
    )

    result, issues = adapter.review({"report": "示例"})

    assert issues == []
    assert result["status"] == "completed"
    assert result["findings"][0]["severity"] == "high"
    request = client.request["kwargs"]
    assert request["json"]["model"] == "qwen3.7-flash"
    assert request["json"]["response_format"]["type"] == "json_schema"


def test_review_adapter_failure_is_a_review_issue_without_secret_leak():
    class BrokenClient:
        def post(self, *args, **kwargs):
            raise RuntimeError("request failed")

    result, issues = BailianReviewAdapter(
        client=BrokenClient(),
        api_key="secret-never-print",
        prompt="审核规则",
        task="semantic_review",
    ).review({"report": "示例"})

    assert result["status"] == "failed"
    assert issues == ["LLM semantic_review 失败：request failed"]
    assert "secret-never-print" not in issues[0]
