from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from demo.domain.llm_config import DEFAULT_LLM_MODEL


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
FIELD_ORDER = (
    "company_profile_section",
    "industry_overview",
    "business_and_segments",
    "main_products",
    "customers_suppliers",
    "profit_model_swot",
    "comparable_list",
)
# These snapshots are thinking-only models.  DashScope rejects an explicit
# ``enable_thinking=false`` for them; omitting the parameter is required.
THINKING_ONLY_MODELS = frozenset(
    {
        "qwen3.7-max-preview",
        "qwen3.7-max-2026-05-17",
    }
)
FIELD_KEYWORDS = {
    "company_profile_section": ("公司", "企业", "成立", "注册资本", "工商", "经营范围", "集团"),
    "industry_overview": ("行业", "市场", "应用领域", "机械", "建筑", "能源", "汽车", "医疗", "通信"),
    "business_and_segments": ("业务", "生产模式", "销售模式", "采购模式", "研发模式", "经营"),
    "main_products": ("产品", "主要生产", "滤波器", "电感", "电抗器", "元器件", "服务"),
    "customers_suppliers": ("客户", "供应商", "采购", "销售", "原材料"),
    "profit_model_swot": ("盈利", "利润", "优势", "劣势", "风险", "竞争", "生产模式", "销售模式"),
    "comparable_list": ("可比", "同行", "上市公司", "竞争", "行业"),
}


OUTPUT_SCHEMA = json.loads(
    (Path(__file__).resolve().parents[1] / "prompts/yellow_narratives_output.v2.json").read_text(
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
        model: str = DEFAULT_LLM_MODEL,
        prompt_version: str = "yellow_narratives.v2",
    ):
        self.client = client
        self.api_key = api_key
        self.prompt = prompt
        self.base_url = base_url.rstrip("/")
        self.model = model
        self.prompt_version = prompt_version

    def _request(self, evidence: dict[str, Any], *, requested_field: str | None = None) -> dict[str, Any]:
        system_prompt = self.prompt
        if requested_field:
            system_prompt += (
                f"\n本次只处理字段 {requested_field}。"
                f"返回扁平 JSON 对象，只包含键 {requested_field}，"
                "其值必须是包含 value 和 evidence_ids 的对象。"
            )
        request_payload = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": json.dumps(evidence, ensure_ascii=False)},
            ],
            # Compatible-mode models may accept JSON Schema syntax but
            # ignore nested ``required`` constraints.  Request JSON and
            # enforce the seven-field/evidence contract locally.
            "response_format": {"type": "json_object"},
        }
        if self.model not in THINKING_ONLY_MODELS:
            request_payload["enable_thinking"] = False
        response = self.client.post(
            f"{self.base_url}/chat/completions",
            headers={"Authorization": f"Bearer {self.api_key}"},
            json=request_payload,
        )
        response.raise_for_status()
        content = response.json()["choices"][0]["message"]["content"]
        return json.loads(content)

    @staticmethod
    def _relevant_evidence(
        field_key: str,
        evidence: list[dict[str, Any]],
        *,
        limit: int = 12,
    ) -> list[dict[str, Any]]:
        keywords = FIELD_KEYWORDS[field_key]
        scored: list[tuple[int, int, dict[str, Any]]] = []
        for index, item in enumerate(evidence):
            text = str(item.get("text", ""))
            score = sum(1 for keyword in keywords if keyword.lower() in text.lower())
            if field_key == "company_profile_section" and str(
                item.get("evidence_id", "")
            ).startswith("api:qichacha:"):
                score += 20
            if score:
                scored.append((score, index, item))
        if not scored:
            return evidence[: min(limit, len(evidence))]
        return [
            item
            for _, _, item in sorted(scored, key=lambda value: (-value[0], value[1]))[:limit]
        ]

    @staticmethod
    def _validated_values(
        payload: dict[str, Any],
        known_evidence: set[str],
    ) -> tuple[dict[str, str], list[str]]:
        values: dict[str, str] = {}
        issues: list[str] = []
        generated_fields = payload.get("fields")
        if not isinstance(generated_fields, dict):
            generated_fields = payload
        for field_key, generated in generated_fields.items():
            if field_key not in ALLOWED_FIELDS:
                issues.append(f"GLM 返回未授权字段，已丢弃：{field_key}")
                continue
            if not isinstance(generated, dict):
                issues.append(f"GLM 字段 {field_key} 结构无效，已丢弃")
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

    def generate(self, evidence: dict[str, Any]) -> tuple[dict[str, str], list[str]]:
        selected_modules = evidence.get("selected_modules")
        if isinstance(selected_modules, list):
            requested = [
                field_key
                for field_key in FIELD_ORDER
                if field_key == "company_profile_section"
                or field_key in selected_modules
            ]
            values: dict[str, str] = {}
            issues: list[str] = []
            all_evidence = list(evidence.get("evidence", []))
            for field_key in requested:
                field_evidence = self._relevant_evidence(field_key, all_evidence)
                field_payload = {
                    "requested_field": field_key,
                    "evidence": field_evidence,
                }
                try:
                    payload = self._request(
                        field_payload,
                        requested_field=field_key,
                    )
                except Exception as exc:
                    issues.append(f"百炼 GLM 字段 {field_key} 失败：{exc}")
                    continue
                known_evidence = {
                    str(item.get("evidence_id"))
                    for item in field_evidence
                    if item.get("evidence_id")
                }
                field_values, field_issues = self._validated_values(
                    payload,
                    known_evidence,
                )
                if field_key in field_values:
                    values[field_key] = field_values[field_key]
                issues.extend(field_issues)
            return values, issues

        try:
            payload = self._request(evidence)
        except Exception as exc:
            return {}, [f"百炼 GLM 失败：{exc}"]

        known_evidence = {
            str(item.get("evidence_id"))
            for item in evidence.get("evidence", [])
            if item.get("evidence_id")
        }
        return self._validated_values(payload, known_evidence)
