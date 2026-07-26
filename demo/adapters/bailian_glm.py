from __future__ import annotations

import json
from pathlib import Path
from typing import Any


ALLOWED_FIELDS = frozenset(
    {
        "company_profile_section",
        "industry_overview",
        "business_and_segments",
        "main_products",
        "customers_suppliers",
        "profit_model_swot",
        "comparable_list",
    }
)


OUTPUT_SCHEMA = json.loads(
    (Path(__file__).resolve().parents[1] / "prompts/yellow_narratives_output.v1.json").read_text(
        encoding="utf-8"
    )
)


class BailianYellowNarrativeAdapter:
    """百炼 OpenAI 兼容接口；不读取环境变量，只处理获准的七个字段。"""

    def __init__(
        self,
        client: Any,
        api_key: str,
        prompt: str,
        *,
        base_url: str = "https://dashscope.aliyuncs.com/compatible-mode/v1",
        model: str = "qwen3.7-flash",
        prompt_version: str = "yellow_narratives.v1",
    ):
        self.client = client
        self.api_key = api_key
        self.prompt = prompt
        self.base_url = base_url.rstrip("/")
        self.model = model
        self.prompt_version = prompt_version

    def generate(self, evidence: dict[str, Any]) -> tuple[dict[str, str], list[str]]:
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
                            "name": "yellow_narratives",
                            "strict": True,
                            "schema": OUTPUT_SCHEMA,
                        },
                    },
                },
            )
            response.raise_for_status()
            content = response.json()["choices"][0]["message"]["content"]
            payload = json.loads(content)
        except Exception as exc:
            return {}, [f"百炼 GLM 失败：{exc}"]

        known_evidence = {
            str(item.get("evidence_id"))
            for item in evidence.get("evidence", [])
            if item.get("evidence_id")
        }
        values: dict[str, str] = {}
        issues: list[str] = []
        for field_key, generated in payload.get("fields", {}).items():
            if field_key not in ALLOWED_FIELDS:
                issues.append(f"GLM 返回未授权字段，已丢弃：{field_key}")
                continue
            value = str(generated.get("value", "")).strip()
            evidence_ids = [str(item) for item in generated.get("evidence_ids", [])]
            unknown = sorted(set(evidence_ids) - known_evidence)
            if unknown:
                issues.append(f"GLM 字段 {field_key} 引用了未知证据：{'、'.join(unknown)}")
                continue
            if value and not evidence_ids:
                issues.append(f"GLM 字段 {field_key} 没有证据编号，已丢弃")
                continue
            values[field_key] = value
        return values, issues
