from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from demo.domain.llm_config import DEFAULT_LLM_MODEL


REVIEW_SCHEMA = json.loads(
    (Path(__file__).resolve().parents[1] / "prompts/review_output.v1.json").read_text(
        encoding="utf-8"
    )
)


class BailianReviewAdapter:
    """调用百炼兼容接口执行一个结构化审核任务。"""

    def __init__(
        self,
        client: Any,
        api_key: str,
        prompt: str,
        *,
        task: str,
        base_url: str = "https://dashscope.aliyuncs.com/compatible-mode/v1",
        model: str = DEFAULT_LLM_MODEL,
        prompt_version: str = "review.v1",
    ):
        self.client = client
        self.api_key = api_key
        self.prompt = prompt
        self.task = task
        self.base_url = base_url.rstrip("/")
        self.model = model
        self.prompt_version = prompt_version

    def review(self, evidence: dict[str, Any]) -> tuple[dict[str, Any], list[str]]:
        try:
            response = self.client.post(
                f"{self.base_url}/chat/completions",
                headers={"Authorization": f"Bearer {self.api_key}"},
                json={
                    "model": self.model,
                    "enable_thinking": False,
                    "messages": [
                        {"role": "system", "content": self.prompt},
                        {"role": "user", "content": json.dumps(evidence, ensure_ascii=False)},
                    ],
                    "response_format": {
                        "type": "json_schema",
                        "json_schema": {
                            "name": "appraisal_review",
                            "strict": True,
                            "schema": REVIEW_SCHEMA,
                        },
                    },
                },
            )
            response.raise_for_status()
            payload = json.loads(response.json()["choices"][0]["message"]["content"])
        except Exception as exc:
            return {
                "review_type": self.task,
                "status": "failed",
                "summary": "",
                "findings": [],
                "model": self.model,
                "prompt_version": self.prompt_version,
            }, [f"LLM {self.task} 失败：{exc}"]

        findings: list[dict[str, Any]] = []
        for item in payload.get("findings", []):
            if not isinstance(item, dict):
                continue
            findings.append(
                {
                    "location": str(item.get("location", "")),
                    "severity": str(item.get("severity", "low")),
                    "category": str(item.get("category", "")),
                    "problem": str(item.get("problem", "")),
                    "evidence": str(item.get("evidence", "")),
                    "suggestion": str(item.get("suggestion", "")),
                }
            )
        status = str(payload.get("status", "completed"))
        if status not in {"completed", "completed_with_issues", "failed"}:
            status = "completed"
        return {
            "review_type": self.task,
            "status": status,
            "summary": str(payload.get("summary", "")),
            "findings": findings,
            "model": self.model,
            "prompt_version": self.prompt_version,
        }, []
