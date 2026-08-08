from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

from demo.domain.llm_config import DEFAULT_LLM_FALLBACK_MODEL, DEFAULT_LLM_MODEL


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


# These texts are deliberately statements about the available record, not
# substitutes for facts that were not supplied.  They keep the six optional
# narrative modules readable without turning missing source material into a
# yellow ``XXX`` or an invented business fact.
NO_EVIDENCE_STATEMENTS = {
    "industry_overview": "现有已上传材料及已调用企业信息接口未提供可直接核验的所处行业及行业介绍，未据此作进一步描述。",
    "business_and_segments": "现有已上传材料及已调用企业信息接口未提供可直接核验的业务内容及细分市场信息，未据此作进一步描述。",
    "main_products": "现有已上传材料及已调用企业信息接口未提供可直接核验的主要产品信息，未据此作进一步描述。",
    "customers_suppliers": "现有已上传材料及已调用企业信息接口未披露主要客户及供应商，未据此识别具体交易对手。",
    "comparable_list": "现有材料及已调用企业信息接口未获得同时包含证券代码、公告标题和公告日期的上市公司候选，未据此形成对标上市公司清单。",
}
SWOT_FALLBACKS = (
    ("盈利模式：", "现有材料未提供可直接核验的盈利模式证据。"),
    ("优势：", "现有材料未提供可直接核验的竞争优势证据。"),
    ("劣势：", "现有材料未提供可直接核验的竞争劣势证据。"),
    ("机会：", "现有材料未提供可直接核验的发展机会证据。"),
    ("风险：", "现有材料未提供可直接核验的主要风险证据。"),
)


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
        fallback_model: str = DEFAULT_LLM_FALLBACK_MODEL,
        prompt_version: str = "yellow_narratives.v2",
    ):
        self.client = client
        self.api_key = api_key
        self.prompt = prompt
        self.base_url = base_url.rstrip("/")
        self.model = model
        self.fallback_model = fallback_model
        self.prompt_version = prompt_version

    def _request_for_model(
        self,
        model: str,
        evidence: dict[str, Any],
        *,
        requested_field: str | None = None,
    ) -> dict[str, Any]:
        system_prompt = self.prompt
        if requested_field:
            system_prompt += (
                f"\n本次只处理字段 {requested_field}。"
                f"返回扁平 JSON 对象，只包含键 {requested_field}，"
                "其值必须是包含 value 和 evidence_ids 的对象。"
            )
        request_payload = {
            "model": model,
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": json.dumps(evidence, ensure_ascii=False)},
            ],
            # Compatible-mode models may accept JSON Schema syntax but
            # ignore nested ``required`` constraints.  Request JSON and
            # enforce the seven-field/evidence contract locally.
            "response_format": {"type": "json_object"},
        }
        if model not in THINKING_ONLY_MODELS:
            request_payload["enable_thinking"] = False
        response = self.client.post(
            f"{self.base_url}/chat/completions",
            headers={"Authorization": f"Bearer {self.api_key}"},
            json=request_payload,
        )
        response.raise_for_status()
        content = response.json()["choices"][0]["message"]["content"]
        return json.loads(content)

    def _request(self, evidence: dict[str, Any], *, requested_field: str | None = None) -> dict[str, Any]:
        try:
            return self._request_for_model(self.model, evidence, requested_field=requested_field)
        except Exception as primary_error:
            if not self.fallback_model or self.fallback_model == self.model:
                raise primary_error
            try:
                return self._request_for_model(
                    self.fallback_model,
                    evidence,
                    requested_field=requested_field,
                )
            except Exception as fallback_error:
                raise RuntimeError(
                    f"主模型 {self.model} 失败：{primary_error}；"
                    f"降级模型 {self.fallback_model} 失败：{fallback_error}"
                ) from fallback_error

    @staticmethod
    def _relevant_evidence(
        field_key: str,
        evidence: list[dict[str, Any]],
        *,
        limit: int = 12,
    ) -> list[dict[str, Any]]:
        # The six narrative modules describe the target company.  Once its
        # QCC profile is available, exclude the commissioning party rather
        # than letting a similarly worded business scope leak into the
        # target's industry or product description.
        target_profile_present = any(
            "api:qichacha:target:" in str(item.get("evidence_id", ""))
            for item in evidence
        )
        eligible_evidence = (
            [
                item
                for item in evidence
                if "api:qichacha:commissioning:" not in str(item.get("evidence_id", ""))
            ]
            if target_profile_present
            else evidence
        )
        if field_key == "comparable_list":
            # A target-company profile can be evidence about the target's
            # industry, but it can never be evidence that the target is its
            # own comparable.  Permit only actual API peer candidates or an
            # uploaded material explicitly labelled as a comparable list.
            peer_evidence = []
            for item in eligible_evidence:
                evidence_id = str(item.get("evidence_id", ""))
                text = str(item.get("text", ""))
                # A QCC fuzzy-search result is a business-name lead only. It
                # must not be treated as a listed comparable.  Api 915
                # records are sufficient only when the report can show the
                # company, stock code, announcement title and date.
                announcement = ":915:" in evidence_id and all(
                    marker in text for marker in ("股票代码：", "公告：", "日期：")
                )
                uploaded_list = "可比上市公司" in text and all(
                    marker in text for marker in ("股票代码", "公告", "日期")
                )
                if announcement or uploaded_list:
                    peer_evidence.append(item)
            return peer_evidence[:limit]
        keywords = FIELD_KEYWORDS[field_key]
        scored: list[tuple[int, int, dict[str, Any]]] = []
        for index, item in enumerate(eligible_evidence):
            text = str(item.get("text", ""))
            score = sum(1 for keyword in keywords if keyword.lower() in text.lower())
            if field_key == "company_profile_section" and str(
                item.get("evidence_id", "")
            ).startswith("api:qichacha:"):
                score += 20
            if score:
                scored.append((score, index, item))
        if not scored:
            return eligible_evidence[: min(limit, len(eligible_evidence))]
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

    @staticmethod
    def _normalize_profit_model_swot(value: str) -> tuple[str, bool]:
        """Require the five report headings without inferring missing facts."""
        text = str(value or "").strip()
        if all(label in text for label, _ in SWOT_FALLBACKS):
            return text, False
        sections: list[str] = []
        for label, fallback in SWOT_FALLBACKS:
            if label in text:
                # Preserve the model's evidence-backed wording as-is.  The
                # report only needs labels absent from a sparse response.
                continue
            if label == "盈利模式：" and text:
                sections.append(f"{label}{text.rstrip('。')}。")
            else:
                sections.append(f"{label}{fallback}")
        return "".join([text, *sections]), True

    @staticmethod
    def _industry_fallback(field_evidence: list[dict[str, Any]]) -> str:
        """Render a direct industry field from target API evidence only."""
        patterns = (
            r"(?:所属行业为|所属行业[：:]|行业[：:])\s*([^，；。\n}\]\"']{2,80})",
            r"[\"']Industry[\"']\s*[:：]\s*[\"']?([^，；。\n}\]\"']{2,80})",
        )
        for item in field_evidence:
            text = str(item.get("text", ""))
            for pattern in patterns:
                match = re.search(pattern, text)
                if match:
                    industry = match.group(1).strip()
                    if industry:
                        return f"根据已调用企业信息接口，被评估单位所属行业为{industry}。"
        return NO_EVIDENCE_STATEMENTS["industry_overview"]

    @staticmethod
    def _comparable_fallback(field_evidence: list[dict[str, Any]]) -> str:
        records = [str(item.get("text", "")).strip() for item in field_evidence]
        if not records:
            return NO_EVIDENCE_STATEMENTS["comparable_list"]
        return "以下为按经营关键词命中的上市公司公告候选，不等于可比性最终认定。" + "\n".join(records)

    @staticmethod
    def _normalize_selected_value(
        field_key: str,
        value: str,
        field_evidence: list[dict[str, Any]],
    ) -> tuple[str, str | None]:
        value = str(value or "").strip()
        if not value:
            if field_key == "industry_overview":
                return BailianYellowNarrativeAdapter._industry_fallback(field_evidence), "未返回内容，已从企业信息接口回退生成行业说明"
            if field_key == "profit_model_swot":
                normalized, _ = BailianYellowNarrativeAdapter._normalize_profit_model_swot("")
                return normalized, "未返回内容，已按五个维度写入可核验缺失说明"
            if field_key == "comparable_list":
                return BailianYellowNarrativeAdapter._comparable_fallback(field_evidence), "未返回内容，已按企查查公告证据生成候选说明"
            return NO_EVIDENCE_STATEMENTS.get(field_key, ""), "未返回内容，已写入材料未披露说明"
        if field_key == "profit_model_swot":
            normalized, changed = BailianYellowNarrativeAdapter._normalize_profit_model_swot(value)
            return normalized, "未覆盖盈利模式及 SWOT 五个维度，已补充可核验缺失说明" if changed else None
        if field_key == "comparable_list" and not all(
            marker in value for marker in ("股票代码", "公告", "日期")
        ):
            return BailianYellowNarrativeAdapter._comparable_fallback(field_evidence), "未按多维公告格式输出，已使用企查查公告证据生成候选清单"
        return value, None

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
                if not field_evidence:
                    if field_key != "company_profile_section":
                        value, issue = self._normalize_selected_value(field_key, "", [])
                        if value:
                            values[field_key] = value
                        if issue:
                            issues.append(f"GLM 字段 {field_key}：{issue}")
                    continue
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
                value, normalization_issue = self._normalize_selected_value(
                    field_key,
                    field_values.get(field_key, ""),
                    field_evidence,
                )
                if value:
                    values[field_key] = value
                issues.extend(field_issues)
                if normalization_issue:
                    issues.append(f"GLM 字段 {field_key}：{normalization_issue}")
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
