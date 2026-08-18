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
        self.requests = []

    def post(self, *args, **kwargs):
        self.request = {"args": args, "kwargs": kwargs}
        self.requests.append(self.request)
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


class EvidenceReviewClient:
    def __init__(self):
        self.requests = []

    def post(self, *args, **kwargs):
        self.requests.append({"args": args, "kwargs": kwargs})
        user_payload = json.loads(kwargs["json"]["messages"][-1]["content"])
        field_key = user_payload["fields"][0]["field_key"]
        return FakeResponse(
            {
                "choices": [
                    {
                        "message": {
                            "content": json.dumps(
                                {
                                    "reviews": [
                                        {
                                            "field_key": field_key,
                                            "status": "accept",
                                            "reason": "PDF和表格的期间、口径及单位一致。",
                                            "comment": "已核对审计报告页码与对应工作表。",
                                        }
                                    ]
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
    assert request["json"]["model"] == "deepseek-v4-pro-0813"
    assert request["json"]["enable_thinking"] is False
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


def test_glm_passes_feedback_to_single_module_regeneration_request():
    client = FieldwiseClient()
    adapter = BailianYellowNarrativeAdapter(
        client=client,
        api_key="test-key",
        prompt="规则",
    )

    values, issues = adapter.generate(
        {
            "selected_modules": ["main_products"],
            "regenerate_field": "main_products",
            "feedback": "补充产品应用场景并分段",
            "evidence": [
                {"evidence_id": "pdf:p1:b1", "text": "主要产品：工业滤波器。"}
            ],
        }
    )

    request = json.loads(client.requests[0]["json"]["messages"][-1]["content"])
    assert request["user_feedback"] == "补充产品应用场景并分段"
    assert request["requested_field"] == "main_products"
    assert len(client.requests) == 1
    assert values["main_products"] == "main_products内容"
    assert issues == []


def test_glm_reviews_extracted_data_without_returning_a_replacement_value():
    client = FakeClient(
        json.dumps(
            {
                "reviews": [
                    {
                        "field_key": "book_net_assets",
                        "status": "needs_review",
                        "reason": "PDF与Excel期间标签不一致。",
                        "value": "不得接受的替换金额",
                    }
                ]
            },
            ensure_ascii=False,
        )
    )
    adapter = BailianYellowNarrativeAdapter(
        client=client,
        api_key="test-key",
        prompt="叙述规则",
        review_prompt="复核规则",
        review_model="review-model",
    )

    reviews, issues = adapter.review_extracted_evidence(
        [
            {
                "field_key": "book_net_assets",
                "selected": {
                    "value": "100.00",
                    "source_kind": "pdf_ocr",
                    "source_file": "审计报告.pdf",
                    "source_locator": "第10页",
                },
                "candidates": [],
            }
        ],
        {"book_net_assets": "账面净资产"},
    )

    assert reviews == [
        {
            "field_key": "book_net_assets",
            "status": "needs_review",
            "reason": "PDF与Excel期间标签不一致。",
        }
    ]
    assert issues == []
    assert client.request["kwargs"]["json"]["model"] == "review-model"
    assert client.request["kwargs"]["json"]["messages"][0]["content"] == "复核规则"


def test_glm_reviews_each_table_group_with_a_separate_request():
    client = EvidenceReviewClient()
    adapter = BailianYellowNarrativeAdapter(
        client=client,
        api_key="test-key",
        prompt="叙述规则",
        review_prompt="复核规则",
        review_model="review-model",
    )

    reviews, issues = adapter.review_extracted_evidence(
        [
            {
                "field_key": "historical_balance_sheet_table",
                "selected": {"value": {"rows": [["项目", "金额"], ["总资产", "100"]]}, "source_kind": "pdf_ocr", "source_file": "审计.pdf", "source_locator": "审计 PDF 第18页：历史资产负债表"},
                "candidates": [],
            },
            {
                "field_key": "historical_income_statement_table",
                "selected": {"value": {"rows": [["项目", "金额"], ["营业收入", "90"]]}, "source_kind": "pdf_ocr", "source_file": "审计.pdf", "source_locator": "审计 PDF 第19页：历史利润表"},
                "candidates": [],
            },
        ],
        {"historical_balance_sheet_table": "历史资产负债表", "historical_income_statement_table": "历史利润表"},
    )

    assert len(client.requests) == 2
    assert [item["field_key"] for item in reviews] == [
        "historical_balance_sheet_table",
        "historical_income_statement_table",
    ]
    assert issues == []


def test_glm_can_locate_unresolved_pdf_pages_from_numbered_ocr_pages():
    client = FakeClient(
        json.dumps(
            {
                "locations": [
                    {
                        "field_key": "historical_income_statement_table",
                        "page_number": 27,
                        "reason": "第27页包含利润表标题和营业收入、净利润行。",
                    }
                ]
            },
            ensure_ascii=False,
        )
    )
    adapter = BailianYellowNarrativeAdapter(
        client=client,
        api_key="test-key",
        prompt="叙述规则",
        review_prompt="复核规则",
        review_model="review-model",
    )

    locations, issues = adapter.locate_pdf_pages(
        {
            "text_blocks": [{"page_number": 27, "text": "利润表 营业收入 净利润"}],
            "table_cells": [],
        },
        {"historical_income_statement_table": "历史利润表"},
        ["historical_income_statement_table"],
    )

    assert locations == {"historical_income_statement_table": 27}
    assert issues == []
    request = client.request["kwargs"]["json"]
    assert request["model"] == "review-model"
    assert request["response_format"]["type"] == "json_object"
    assert request["messages"][1]["content"].find("第27页") >= 0


def test_glm_ignores_target_context_pseudo_citation_when_real_evidence_remains():
    values, issues = BailianYellowNarrativeAdapter._validated_values(
        {
            "profit_model_swot": {
                "value": "盈利模式：以加工服务费取得收入。",
                "evidence_ids": [
                    "field:target_context",
                    "field:historical_income_statement_table",
                ],
            }
        },
        {"field:historical_income_statement_table"},
    )

    assert values == {"profit_model_swot": "盈利模式：以加工服务费取得收入。"}
    assert issues == []


def test_glm_rejects_target_context_as_the_only_substantive_evidence():
    values, issues = BailianYellowNarrativeAdapter._validated_values(
        {
            "profit_model_swot": {
                "value": "无材料支持的具体结论。",
                "evidence_ids": ["field:target_context"],
            }
        },
        set(),
    )

    assert values == {}
    assert any("没有证据编号" in issue for issue in issues)


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


def test_glm_retries_with_qwen_fallback_after_primary_request_failure():
    class PrimaryFailureThenSuccessClient:
        def __init__(self):
            self.requests = []

        def post(self, *args, **kwargs):
            self.requests.append(kwargs)
            if len(self.requests) == 1:
                raise RuntimeError("primary model unavailable")
            return FakeResponse(
                {
                    "choices": [
                        {
                            "message": {
                                "content": json.dumps(
                                    {
                                        "company_profile_section": {
                                            "value": "示例概述",
                                            "evidence_ids": ["pdf:p1:b1"],
                                        }
                                    },
                                    ensure_ascii=False,
                                )
                            }
                        }
                    ]
                }
            )

    client = PrimaryFailureThenSuccessClient()
    adapter = BailianYellowNarrativeAdapter(
        client=client,
        api_key="test-key",
        prompt="规则",
        model="deepseek-v4-flash-0731",
        fallback_model="qwen3.8-max",
    )

    values, issues = adapter.generate({"evidence": [{"evidence_id": "pdf:p1:b1", "text": "示例材料"}]})

    assert values == {"company_profile_section": "示例概述"}
    assert issues == []
    assert [request["json"]["model"] for request in client.requests] == [
        "deepseek-v4-flash-0731",
        "qwen3.8-max",
    ]


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


def test_glm_sends_target_context_with_each_selected_module():
    client = FieldwiseClient()
    adapter = BailianYellowNarrativeAdapter(
        client=client,
        api_key="test-key",
        prompt="规则",
    )

    adapter.generate(
        {
            "selected_modules": ["main_products"],
            "evidence": [
                {
                    "evidence_id": "field:target_company_name",
                    "text": "被评估单位名称：通富热处理（昆山）有限公司",
                },
                {
                    "evidence_id": "field:valuation_date_year",
                    "text": "评估基准日年：2025",
                },
                {
                    "evidence_id": "field:valuation_date_month",
                    "text": "评估基准日月：6",
                },
                {
                    "evidence_id": "field:valuation_date_day",
                    "text": "评估基准日日：30",
                },
                {
                    "evidence_id": "field:registered_capital",
                    "text": "注册资本：1,000万元",
                },
                {
                    "evidence_id": "field:main_products",
                    "text": "主要产品：汽车零部件热处理加工服务",
                },
            ],
        }
    )

    for request in client.requests:
        payload = json.loads(request["json"]["messages"][-1]["content"])
        assert payload["target_context"] == {
            "target_company_name": "通富热处理（昆山）有限公司",
            "registered_capital": "1,000万元",
            "valuation_date": "2025年6月30日",
        }


def test_glm_uses_a_factual_no_disclosure_statement_for_missing_customer_supplier_evidence():
    class CustomerSupplierEmptyClient(FieldwiseClient):
        def post(self, *args, **kwargs):
            self.requests.append(kwargs)
            request = json.loads(kwargs["json"]["messages"][-1]["content"])
            field = request["requested_field"]
            evidence_id = request["evidence"][0]["evidence_id"]
            value = "" if field == "customers_suppliers" else f"{field}内容"
            return FakeResponse(
                {
                    "choices": [
                        {
                            "message": {
                                "content": json.dumps(
                                    {
                                        field: {
                                            "value": value,
                                            "evidence_ids": [evidence_id] if value else [],
                                        }
                                    },
                                    ensure_ascii=False,
                                )
                            }
                        }
                    ]
                }
            )

    values, issues = BailianYellowNarrativeAdapter(
        client=CustomerSupplierEmptyClient(), api_key="test-key", prompt="规则"
    ).generate(
        {
            "selected_modules": ["customers_suppliers"],
            "evidence": [
                {
                    "evidence_id": "api:qichacha:target:735:profile",
                    "text": "示例有限公司经营工业设备业务。",
                }
            ],
        }
    )

    assert values["customers_suppliers"] == (
        "现有已上传材料及已调用企业信息接口未披露主要客户及供应商，"
        "未据此识别具体交易对手。"
    )
    assert any("customers_suppliers" in issue for issue in issues)


def test_glm_uses_target_company_industry_when_model_omits_evidence_ids():
    class EmptyIndustryClient(FieldwiseClient):
        def post(self, *args, **kwargs):
            self.requests.append(kwargs)
            request = json.loads(kwargs["json"]["messages"][-1]["content"])
            field = request["requested_field"]
            evidence_id = request["evidence"][0]["evidence_id"]
            value = "" if field == "industry_overview" else f"{field}内容"
            return FakeResponse(
                {"choices": [{"message": {"content": json.dumps({field: {
                    "value": value, "evidence_ids": [evidence_id] if value else []
                }}, ensure_ascii=False)}}]}
            )

    values, _ = BailianYellowNarrativeAdapter(
        client=EmptyIndustryClient(), api_key="test-key", prompt="规则"
    ).generate(
        {
            "selected_modules": ["industry_overview"],
            "evidence": [
                {
                    "evidence_id": "api:qichacha:commissioning:735:profile",
                    "text": "委托方工商信息：行业：贸易业。",
                },
                {
                    "evidence_id": "api:qichacha:target:735:profile",
                    "text": "被评估单位工商信息：行业：专用设备制造业；经营范围：热处理加工。",
                },
            ],
        }
    )

    assert values["industry_overview"] == "根据已调用企业信息接口，被评估单位所属行业为专用设备制造业。"


def test_glm_excludes_commissioning_party_evidence_when_target_profile_exists():
    target = {
        "evidence_id": "api:qichacha:target:735:profile",
        "text": "被评估单位工商信息：经营范围：汽车零部件热处理加工。",
    }
    commissioning = {
        "evidence_id": "api:qichacha:commissioning:735:profile",
        "text": "委托方工商信息：经营范围：涂料销售。",
    }

    selected = BailianYellowNarrativeAdapter._relevant_evidence(
        "business_and_segments", [commissioning, target]
    )

    assert selected == [target]


def test_glm_normalizes_profit_model_swot_to_all_required_dimensions():
    class SparseSwotClient(FieldwiseClient):
        def post(self, *args, **kwargs):
            self.requests.append(kwargs)
            request = json.loads(kwargs["json"]["messages"][-1]["content"])
            field = request["requested_field"]
            evidence_id = request["evidence"][0]["evidence_id"]
            return FakeResponse(
                {
                    "choices": [
                        {
                            "message": {
                                "content": json.dumps(
                                    {
                                        field: {
                                            "value": "2024年度营业收入为100万元。",
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

    values, issues = BailianYellowNarrativeAdapter(
        client=SparseSwotClient(), api_key="test-key", prompt="规则"
    ).generate(
        {
            "selected_modules": ["profit_model_swot"],
            "evidence": [
                {
                    "evidence_id": "field:historical_income_statement_table",
                    "text": "营业收入：100万元。",
                }
            ],
        }
    )

    assert all(label in values["profit_model_swot"] for label in ("盈利模式：", "优势：", "劣势：", "机会：", "风险："))
    assert any("profit_model_swot" in issue for issue in issues)


def test_glm_removes_unsupported_valuation_method_and_finance_inferences():
    value, issue = BailianYellowNarrativeAdapter._normalize_selected_value(
        "profit_model_swot",
        (
            "盈利模式：公司通过加工服务取得收入。"
            "优势：财务费用为负，说明公司资金状况良好，利息收入较多；"
            "劣势：收入存在波动。"
            "机会：采用收益法表明公司未来盈利能力受到认可。"
            "风险：客户集中度未披露。"
        ),
        [],
    )

    assert "盈利能力受到认可" not in value
    assert "资金状况良好" not in value
    assert "具体原因现有材料未披露" in value
    assert issue == "已移除缺少事实依据的评估方法或财务费用推断"


def test_glm_does_not_duplicate_complete_swot_when_model_uses_aspect_headings():
    value, changed = BailianYellowNarrativeAdapter._normalize_profit_model_swot(
        (
            "公司通过加工服务取得收入。"
            "优势方面，公司保持持续经营。"
            "劣势方面，收入存在波动。"
            "机会方面，业务仍有拓展空间。"
            "风险方面，客户结构尚未披露。"
        )
    )

    assert changed is True
    assert value.startswith("盈利模式：公司通过加工服务取得收入。")
    assert value.count("公司通过加工服务取得收入。") == 1
    assert "优势：现有材料未提供" not in value


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
                    "text": "上市公司公告候选：上市公司甲；股票代码：600001；公告：工业设备业务公告；日期：2026-01-01。",
                },
            ],
        }
    )

    assert any("profit_model_swot" in issue for issue in issues)
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
    assert len(issues) == 1
    assert "主模型 deepseek-v4-pro-0813 失败" in issues[0]
    assert "降级模型 qwen3.8-max 失败" in issues[0]
    assert "secret-never-print" not in issues[0]
