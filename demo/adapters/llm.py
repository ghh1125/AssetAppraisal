from __future__ import annotations

import json
from typing import Any


class LlmAdapter:
    def __init__(self, client: Any = None, endpoint: str | None = None, api_key: str | None = None, model: str | None = None, prompt: str = ""):
        self.client, self.endpoint, self.api_key, self.model, self.prompt = client, endpoint, api_key, model, prompt

    def generate(self, evidence: dict[str, Any]) -> tuple[dict[str, Any], list[str]]:
        if not self.client or not self.endpoint or not self.api_key or not self.model:
            return {}, ["LLM 未配置，叙述字段按规则留空"]
        try:
            response = self.client.post(self.endpoint, headers={"Authorization": f"Bearer {self.api_key}"}, json={
                "model": self.model,
                "messages": [{"role": "system", "content": self.prompt}, {"role": "user", "content": json.dumps(evidence, ensure_ascii=False)}],
                "response_format": {"type": "json_object"},
            })
            response.raise_for_status()
            content = response.json()["choices"][0]["message"]["content"]
            return json.loads(content), []
        except Exception as exc:
            return {}, [f"LLM 失败：{exc}"]
