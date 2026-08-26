import hashlib
import json
import re
import zipfile
from pathlib import Path

import pytest
from docx import Document
from openpyxl import Workbook, load_workbook

import demo.pipeline as pipeline_module
from demo.pipeline import _apply_ocr_overrides_to_table, _company_profile_table, _default_ocr_field_resolver, _human_source_locator, _ocr_ownership_matrix, _ownership_matrix_summary, _resolve_unresolved_pdf_page_locators, _table_overview_annotations, _table_trace_annotations, _trace_comment, _validated_qcc_payload, _valuation_date_ownership_matrix, run_pipeline


def test_internal_ocr_locator_is_rendered_as_a_business_table_label():
    assert _human_source_locator(
        "审计报告.pdf",
        "OCR结构化结果.xlsx / asset_scope_summary_table",
        "资产负债范围表",
    ) == "审计 PDF：资产负债范围表"


def test_unresolved_pdf_page_locators_are_completed_by_closed_world_llm():
    class PageLocator:
        def locate_pdf_pages(self, normalized, field_names, field_keys):
            assert field_keys == ["asset_scope_summary_table"]
            return {"asset_scope_summary_table": 44}, []

    locators, issues = _resolve_unresolved_pdf_page_locators(
        {"asset_scope_summary_table": "资产负债范围表"},
        {"asset_scope_summary_table": "审计 PDF：资产负债范围表"},
        {"text_blocks": [{"page_number": 44, "text": "资产负债范围"}]},
        PageLocator(),
    )

    assert locators["asset_scope_summary_table"] == "审计 PDF 第44页：资产负债范围表"
    assert issues == []


def test_ocr_financial_history_prefers_formal_statement_over_prepared_history_table():
    formal_balance = {
        "caption": "通富热处理（昆山）有限公司近年资产负债状况见下表：",
        "rows": [
            ["项目\\报表日", "2023年12月31日", "2024年12月31日", "2025年6月30日"],
            ["总资产", "182,941,174.94", "180,824,246.17", "163,719,131.79"],
            ["负债", "10,043,136.64", "4,871,154.35", "117,737,555.13"],
            ["所有者权益", "172,898,038.30", "175,953,091.82", "45,981,576.66"],
        ],
    }
    prepared_balance = {
        "caption": "被评估单位近年资产负债状况见下表：",
        "rows": [
            ["项目\\报表日", "2023年度", "2024年度", "2025年1-6月"],
            ["总资产", "187,407,553.39", "185,985,773.82", "169,478,239.93"],
            ["负债", "10,043,136.64", "4,871,154.35", "117,737,555.13"],
            ["所有者权益", "177,364,416.75", "181,114,619.47", "51,740,684.80"],
        ],
    }

    values, issues = _default_ocr_field_resolver(
        {
            "financial_data": [
                {
                    "field_key": "historical_balance_sheet_table",
                    "value": formal_balance,
                    "evidence_id": "06N_资产负债表!D3:F3、D76:F76",
                },
                {
                    "field_key": "historical_balance_sheet_table",
                    "value": prepared_balance,
                    "evidence_id": "历资表!A47；历资表!C47:E47",
                },
            ]
        },
        {},
    )

    assert issues == []
    assert values["historical_balance_sheet_table"]["rows"][1][1:] == [
        "182,941,174.94",
        "180,824,246.17",
        "163,719,131.79",
    ]


def test_company_profile_table_writes_credit_code_and_profile_values():
    rows = _company_profile_table(
        {
            "credit_code": "91320000608319749X",
            "name": "示例有限公司",
            "company_type": "有限责任公司",
            "registered_capital": "1,000万元",
            "status": "存续",
        }
    )
    assert rows[0] == [
        "统一社会信用代码：91320000608319749X",
        "企业名称：示例有限公司",
    ]
    assert rows[1][0] == "类型：有限责任公司"
    assert rows[2][0] == "注册资本：1,000万元"


def test_company_profile_table_prefers_valuation_date_material_capital():
    rows = _company_profile_table(
        {
            "name": "示例有限公司",
            "registered_capital": "500万元",
        },
        fallback_capital="1,000万元",
    )

    assert rows[2][0] == "注册资本：1,000万元"


def test_company_profile_table_formats_qcc_dates_for_word():
    rows = _company_profile_table(
        {
            "establish_date": "2011-02-25 00:00:00",
            "term_start": "2011-02-25",
            "approval_date": "2025/06/30",
        }
    )

    assert rows[2][1] == "成立日期：2011年02月25日"
    assert rows[3][0] == "营业期限自：2011年02月25日"
    assert rows[4][1] == "核准日期：2025年06月30日"


def test_qichacha_and_llm_trace_comments_name_the_real_provider_and_review_scope():
    qcc_comment = _trace_comment(
        field_key="target_company_profile",
        field_name="被评估单位工商信息",
        status="verified",
        source={"kind": "qichacha_api", "file": "企查查 API", "locator": "735"},
        llm_review=None,
        model_name="deepseek-v4-pro-0813",
    )
    llm_comment = _trace_comment(
        field_key="industry_overview",
        field_name="所处行业及行业介绍",
        status="verified",
        source={"kind": "bailian_glm", "file": "百炼大模型", "locator": "行业介绍"},
        llm_review=None,
        model_name="deepseek-v4-pro-0813",
    )

    assert qcc_comment.startswith("【来源已核验】")
    assert "企查查API" in qcc_comment and "主体名称匹配" in qcc_comment
    assert llm_comment.startswith("【来源已核验】")
    assert "deepseek-v4-pro-0813" in llm_comment
    assert "经用户选择" in llm_comment


def test_verified_trace_comment_explicitly_names_successful_llm_review():
    comment = _trace_comment(
        field_key="historical_balance_sheet_table",
        field_name="被评估单位历史资产负债表",
        status="verified",
        source={
            "kind": "pdf_ocr_xlsx",
            "file": "审计报告.pdf",
            "locator": "审计 PDF 第28页",
        },
        llm_review={"status": "accept", "reason": "科目、期间和单位一致"},
        model_name="deepseek-v4-pro-0813",
    )

    assert "系统语义规则已完成唯一匹配" in comment
    assert "经大模型核对" in comment
    assert "复核结论为通过" in comment


def test_table_trace_annotations_keep_cell_colours_but_consolidate_same_status_comment():
    annotations = _table_trace_annotations(
        {
            2: [
                ["项目", "2024年度", "2025年1-6月"],
                ["营业收入", "100.00", "120.00"],
                ["投资收益", "XXX", "XXX"],
            ]
        },
        {2: ("historical_income_statement_table", 1, set())},
        {
            "historical_income_statement_table": {
                "kind": "pdf_ocr_xlsx",
                "file": "审计报告.pdf",
                "locator": "审计 PDF 第28页",
            }
        },
        {"historical_income_statement_table": "历史利润表"},
        {"historical_income_statement_table": "verified"},
        {"historical_income_statement_table": {"status": "accept"}},
        "deepseek-v4-pro-0813",
    )

    assert len(annotations) == 6
    assert all(not item["add_comment"] for item in annotations if item["status"] == "verified")
    assert all(item["add_comment"] for item in annotations if item["status"] == "missing")
    assert all(
        item["field_name"] == "历史利润表"
        for item in annotations
        if item["status"] == "verified"
    )
    assert all("投资收益" in item["field_name"] for item in annotations if item["status"] == "missing")
    assert all("整表综合批注" not in item["comment"] for item in annotations)


def test_table_trace_annotations_keep_nonconflicting_cells_verified_in_conflict_table():
    annotations = _table_trace_annotations(
        {
            4: [
                ["项目", "2024年度"],
                ["总资产", "100.00"],
                ["负债", "40.00"],
            ]
        },
        {4: ("historical_balance_sheet_table", 1, set())},
        {"historical_balance_sheet_table": {"kind": "pdf_ocr_xlsx"}},
        {"historical_balance_sheet_table": "历史资产负债表"},
        {"historical_balance_sheet_table": "review"},
        {"historical_balance_sheet_table": {"status": "conflict"}},
        "deepseek-v4-pro-0813",
        conflict_fields={"historical_balance_sheet_table"},
    )

    assert all(item["status"] == "verified" for item in annotations)


def test_table_level_llm_review_marks_the_complete_table_for_attention():
    common = {
        4: [
            ["项目", "2024年度"],
            ["总资产", "100.00"],
        ],
        5: [
            ["项目", "账面金额"],
            ["电子设备", "20.00"],
        ],
    }
    annotations = _table_trace_annotations(
        common,
        {
            4: ("historical_balance_sheet_table", 1, set()),
            5: ("long_term_assets_table", 1, set()),
        },
        {
            "historical_balance_sheet_table": {"kind": "pdf_ocr_xlsx"},
            "long_term_assets_table": {"kind": "asset_workbook"},
        },
        {
            "historical_balance_sheet_table": "历史资产负债表",
            "long_term_assets_table": "主要长期资产表",
        },
        {
            "historical_balance_sheet_table": "review",
            "long_term_assets_table": "review",
        },
        {
            "historical_balance_sheet_table": {"status": "needs_review"},
            "long_term_assets_table": {"status": "needs_review"},
        },
        "deepseek-v4-pro-0813",
        review_fields={"historical_balance_sheet_table", "long_term_assets_table"},
    )

    by_field = {}
    for item in annotations:
        by_field.setdefault(item["field_key"], set()).add(item["status"])
    assert by_field["historical_balance_sheet_table"] == {"review"}
    assert by_field["long_term_assets_table"] == {"review"}


def test_table_level_llm_review_targets_the_complete_table_without_colouring_it():
    cells = _table_trace_annotations(
        {
            4: [
                ["项目", "2024年度"],
                ["总资产", "100.00"],
                ["负债", "40.00"],
            ]
        },
        {4: ("historical_balance_sheet_table", 1, set())},
        {
            "historical_balance_sheet_table": {
                "kind": "pdf_ocr_xlsx",
                "file": "审计报告.pdf",
                "locator": "审计 PDF 第28页",
            }
        },
        {"historical_balance_sheet_table": "历史资产负债表"},
        {"historical_balance_sheet_table": "review"},
        {
            "historical_balance_sheet_table": {
                "status": "needs_review",
                "comment": "请确认单体或合并口径。",
            }
        },
        "deepseek-v4-pro-0813",
        review_fields={"historical_balance_sheet_table"},
    )
    annotations = _table_overview_annotations(
        cells,
        {4: ("historical_balance_sheet_table", 1, set())},
        {
            "historical_balance_sheet_table": {
                "kind": "pdf_ocr_xlsx",
                "file": "审计报告.pdf",
                "locator": "审计 PDF 第28页",
            }
        },
        {"historical_balance_sheet_table": "历史资产负债表"},
        {
            "historical_balance_sheet_table": {
                "status": "needs_review",
                "comment": "请确认单体或合并口径。",
            }
        },
        "deepseek-v4-pro-0813",
        [
            {
                "field_key": "historical_balance_sheet_table",
                "field_name": "历史资产负债表 / 负债 / 2024年度",
            }
        ],
    )

    assert len(annotations) == 1
    assert annotations[0]["whole_table_index"] == 4
    assert annotations[0]["field_key"] == "historical_balance_sheet_table__table_overview"
    assert annotations[0]["status"] == "review"
    assert annotations[0]["colorize"] is False
    assert "共4个填充位置" in annotations[0]["comment"]
    assert "通过0项，不通过4项，待人工核对0项，缺失0项" in annotations[0]["comment"]
    assert "历史资产负债表 / 负债 / 2024年度" in annotations[0]["comment"]
    assert annotations[0]["required_comment_fragments"] == [
        "来源",
        "通过0项",
        "不通过4项",
        "待人工核对0项",
        "缺失0项",
        "历史资产负债表 / 负债 / 2024年度",
    ]


def test_all_missing_table_does_not_double_count_missing_as_review_failure():
    cells = _table_trace_annotations(
        {6: [["无形资产账面金额：", "XXX"], ["使用权资产账面金额：", "XXX"]]},
        {6: ("asset_scope_summary_table", 0, {1})},
        {"asset_scope_summary_table": {"kind": "missing"}},
        {"asset_scope_summary_table": "资产负债范围表"},
        {"asset_scope_summary_table": "missing"},
        {"asset_scope_summary_table": {"status": "missing", "reason": "材料未披露"}},
        "deepseek-v4-pro-0813",
        review_fields={"asset_scope_summary_table"},
    )
    annotations = _table_overview_annotations(
        cells,
        {6: ("asset_scope_summary_table", 0, {1})},
        {"asset_scope_summary_table": {"kind": "missing"}},
        {"asset_scope_summary_table": "资产负债范围表"},
        {"asset_scope_summary_table": {"status": "missing", "reason": "材料未披露"}},
        "deepseek-v4-pro-0813",
    )

    assert "通过0项，不通过0项，待人工核对0项，缺失2项" in annotations[0]["comment"]
    assert "无形资产账面金额：" in annotations[0]["comment"]
    assert "使用权资产账面金额：" in annotations[0]["comment"]


def test_failed_llm_table_review_marks_all_cells_pending_instead_of_inventing_one_failure():
    review = {
        "historical_balance_sheet_table": {
            "status": "needs_review",
            "reason": "LLM未返回有效复核结论，已转为人工复核。",
        }
    }
    cells = _table_trace_annotations(
        {4: [["项目", "2024年度"], ["总资产", "100.00"], ["负债", "40.00"]]},
        {4: ("historical_balance_sheet_table", 1, set())},
        {"historical_balance_sheet_table": {"kind": "pdf_ocr_xlsx"}},
        {"historical_balance_sheet_table": "历史资产负债表"},
        {"historical_balance_sheet_table": "review"},
        review,
        "deepseek-v4-pro-0813",
        review_fields={"historical_balance_sheet_table"},
    )
    annotations = _table_overview_annotations(
        cells,
        {4: ("historical_balance_sheet_table", 1, set())},
        {"historical_balance_sheet_table": {"kind": "pdf_ocr_xlsx"}},
        {"historical_balance_sheet_table": "历史资产负债表"},
        review,
        "deepseek-v4-pro-0813",
    )

    assert {item["status"] for item in cells} == {"review"}
    assert annotations[0]["required_title"] == "【LLM调用失败，需人工复核】"
    assert annotations[0]["comment"].startswith("【LLM调用失败，需人工复核】")
    assert "请检查API Key、模型权限/额度、网络或返回结构" in annotations[0]["comment"]
    assert "通过0项，不通过0项，待人工核对4项，缺失0项" in annotations[0]["comment"]
    assert "整张表已统一标红" in annotations[0]["comment"]


def test_qcc_identity_mismatch_is_rejected_instead_of_filling_wrong_profile():
    issues = []
    assert _validated_qcc_payload(
        {"profile": {"name": "另一家有限公司"}, "fields": {"commissioning_party_profile": "错误"}},
        "目标有限公司",
        "被评估单位",
        issues,
    ) == {}
    assert "企查查身份核验失败" in issues[0]


def test_ocr_ownership_matrix_joins_split_shareholder_names():
    matrix, names = _ocr_ownership_matrix(
        {
            "table_cells": [
                {"table_id": "p39-t4", "row": 2, "column": 1, "text": "上海上大热处理有"},
                {"table_id": "p39-t4", "row": 3, "column": 1, "text": "限公司"},
                {"table_id": "p39-t4", "row": 3, "column": 2, "text": "5,000,000.00"},
                {"table_id": "p39-t4", "row": 4, "column": 1, "text": "富士和机械工业（昆"},
                {"table_id": "p39-t4", "row": 5, "column": 1, "text": "5,000,000.00"},
                {"table_id": "p39-t4", "row": 6, "column": 1, "text": "山）有限公司"},
            ]
        },
        {
            "table_id": "p39-t4",
            "target_table_index": 2,
            "rows": [
                {"name_cells": [[2, 1], [3, 1]], "capital_cell": [3, 2], "percent": "50%"},
                {"name_cells": [[4, 1], [6, 1]], "capital_cell": [5, 1], "percent": "50%"},
            ],
        },
    )
    assert names == ["上海上大热处理有限公司", "富士和机械工业（昆山）有限公司"]
    assert matrix[-1] == ["合计", "合计", "10,000,000.00", "100%"]


def test_valuation_date_ownership_prefers_audit_material_over_current_api():
    historical = [
        ["序号", "股东名称", "总出资（元）", "股权比例"],
        ["1", "股东甲", "5,000,000.00", "50%"],
        ["2", "股东乙", "5,000,000.00", "50%"],
        ["合计", "合计", "10,000,000.00", "100%"],
    ]
    current_api = [
        {"name": "现股东丙", "capital": "500万元", "percent": "100%"},
    ]

    assert _valuation_date_ownership_matrix(historical, current_api) == historical


def test_ownership_summary_keeps_audited_shareholders_and_ignores_total_row():
    matrix = [
        ["序号", "股东名称", "总出资（元）", "股权比例"],
        ["1", "甲公司", "5,000,000.00", "50%"],
        ["2", "乙公司", "5,000,000.00", "50%"],
        ["合计", "合计", "10,000,000.00", "100%"],
    ]

    assert _ownership_matrix_summary(matrix) == (
        "甲公司：出资额5,000,000.00元，股权比例50%；"
        "乙公司：出资额5,000,000.00元，股权比例50%"
    )


def test_ocr_amount_overrides_keep_word_table_in_sync():
    matrix = [["其中：固定资产账面金额：", "4,993,561.04"], ["无形资产账面金额：", "0.00"]]
    spec = {"rows": [
        {"label": "其中：固定资产账面金额：", "ocr_field_key": "fixed"},
        {"label": "无形资产账面金额：", "ocr_field_key": "intangibles"},
    ]}
    assert _apply_ocr_overrides_to_table(matrix, spec, {"fixed": "5,050,511.04", "intangibles": "96,508.64"}) == [
        ["其中：固定资产账面金额：", "5,050,511.04"],
        ["无形资产账面金额：", "96,508.64"],
    ]


class FixtureOcrAdapter:
    def extract(self, pdf_path):
        assert pdf_path.suffix == ".pdf"
        return [
            {
                "page_number": 1,
                "page_count": 1,
                "blocks": [
                    {
                        "block_id": "p1-b1",
                        "block_type": "text",
                        "text": "通富昆山审计材料",
                        "confidence": 0.99,
                        "bbox": [1, 2, 3, 4],
                    }
                ],
                "tables": [],
            }
        ], []


class FixtureLlmAdapter:
    prompt_version = "yellow_narratives.test"
    model = "fixture-model"

    def generate(self, evidence):
        assert evidence["evidence"][0]["evidence_id"] == "pdf:p1:b1"
        return {
            "company_profile_section": "基于 OCR 证据生成的公司概述。",
            "main_products": "热处理服务。",
            "ownership_history": "不应采用的 LLM 越权内容",
        }, ["模拟 LLM 返回了越权字段"]

    def write_traceability_comments(self, annotations):
        return {
            index: "由测试模型根据已确认的来源证据组织；" + str(item["comment"]).split("】", 1)[-1]
            for index, item in enumerate(annotations)
        }, []


class NoPdfEvidenceLlmAdapter:
    prompt_version = "yellow_narratives.test"

    def generate(self, evidence):
        by_id = {item["evidence_id"]: item["text"] for item in evidence["evidence"]}
        assert "field:target_company_name" in by_id
        assert "示例有限公司" in by_id["field:target_company_name"]
        return {"company_profile_section": "根据已上传结构化材料生成。"}, []


class ReferenceDocumentEvidenceLlmAdapter:
    prompt_version = "yellow_narratives.test"

    def generate(self, evidence):
        document_blocks = [
            item
            for item in evidence["evidence"]
            if item["evidence_id"].startswith("document:reference_report:")
        ]
        assert any("主营产品为工业滤波器" in item["text"] for item in document_blocks)
        return {"main_products": "主营产品为工业滤波器。"}, []


class FixtureQichachaAdapter:
    def fetch(self, company_name):
        assert "通富" in company_name
        return {
            "ownership_history": "企查查返回的历史股权沿革。",
            "industry_overview": "不应采用的 API 越权内容",
        }, []


class CapitalQichachaAdapter:
    def fetch(self, company_name):
        return {
            "profile": {
                "name": company_name,
                "registered_capital": "1,250万元",
            },
            "fields": {},
        }, []


class TrackingQichachaAdapter:
    def __init__(self):
        self.calls = []

    def fetch(self, company_name):
        self.calls.append(company_name)
        return {
            "profile": {
                "name": company_name,
                "credit_code": "91320000608319749X",
                "registered_capital": "1,000万元",
                "status": "存续",
            },
            "fields": {
                "commissioning_party_profile": f"企业名称：{company_name}；统一社会信用代码：91320000608319749X；登记状态：存续",
                "target_company_profile": f"企业名称：{company_name}；统一社会信用代码：91320000608319749X；登记状态：存续",
            },
        }, []


class FixtureTemplatePageReader:
    def extract(self, template_path):
        assert template_path.suffix == ".docx"
        return ["XXX有限责任公司拟收购", "纳入评估范围的全部资产和负债。"], []


def fixture_ocr_fields(normalized, config):
    assert normalized["text_blocks"]
    return {
        "tax_rates": "增值税税率13%，企业所得税税率15%。",
        "valuation_scope": "纳入评估范围的全部资产和负债。",
    }, []


def test_pipeline_creates_ocr_xlsx_and_word_with_excel_fallback_when_pdf_field_is_absent(tmp_path):
    config = Path("demo/projects/tongfu.yaml")
    template = Path("templates/评估报告版式_0817确认.docx")
    pdf = Path("资产评估工作流/通富2025.6.30合并及母公司审计报告.pdf")
    template_hash = hashlib.sha256(template.read_bytes()).hexdigest()

    result = run_pipeline(
        project_config=config,
        pdf_path=pdf,
        output_dir=tmp_path,
        ocr_adapter=FixtureOcrAdapter(),
        llm_adapter=FixtureLlmAdapter(),
        qichacha_adapter=FixtureQichachaAdapter(),
        manual_inputs_override={
            "target_company_name": "通富热处理（昆山）有限公司",
            "target_company_short_name": "通富昆山",
            "transaction_type": "收购",
            "valuation_subject_type": "股东全部权益价值",
            "selected_valuation_method": "收益法、资产基础法",
            "final_valuation_method": "收益法",
            "valuation_base_date": "2025-06-30",
            "ownership_history_strategy": "qichacha",
            "unrecorded_intangibles_strategy": "qichacha",
            "company_profile_strategy": "qichacha",
        },
        node_inputs={
            "selected_valuation_method": "收益法、资产基础法",
            "valuation_purpose_inputs": "用于股权收购决策。",
            "company_profile_section": "不应采用的节点越权内容",
        },
        ocr_field_resolver=fixture_ocr_fields,
        template_page_reader=FixtureTemplatePageReader(),
    )

    assert result.ocr_workbook_path.exists()
    assert result.report_path.exists()
    assert not (tmp_path / "字段审计清单.xlsx").exists()
    assert hashlib.sha256(template.read_bytes()).hexdigest() == template_hash
    fields = json.loads((tmp_path / "normalized_fields.json").read_text(encoding="utf-8"))
    assert fields["company_profile_section"] == "基于 OCR 证据生成的公司概述。"
    assert fields["main_products"] == "热处理服务。"
    assert fields["ownership_history"] == "企查查返回的历史股权沿革。"
    # A missing LLM module is never backfilled from a different company's
    # project configuration; it remains unresolved for the reviewer.
    assert fields["industry_overview"] == ""
    with zipfile.ZipFile(result.report_path) as archive:
        document_xml = archive.read("word/document.xml").decode("utf-8")
        comments_xml = archive.read("word/comments.xml").decode("utf-8")
    assert 'w:fill="C6E0B4"' in document_xml
    assert "【来源已核验】" in comments_xml
    assert "【未找到数据】" in comments_xml
    assert "人工基础信息" in comments_xml
    assert "百炼模型" in comments_xml
    assert "由测试模型根据已确认的来源证据组织" in comments_xml
    assert "增值税税率" in fields["tax_rates"]
    assert "13%" in fields["tax_rates"]
    assert "15%" in fields["tax_rates"]
    assert fields["selected_valuation_method"] == "收益法、资产基础法"
    assert fields["commissioning_party_profile"] == ""
    required_monetary = json.loads(config.read_text(encoding="utf-8"))["required_monetary_fields"]
    assert required_monetary == [
        "book_net_assets",
        "income_approach_value",
        "asset_approach_value",
        "historical_balance_sheet_table",
        "historical_income_statement_table",
        "major_long_term_assets",
        "asset_approach_result_section",
    ]
    # The fixture PDF deliberately has no financial tables.  When the
    # requested audit field is absent from PDF OCR, a uniquely identified
    # uploaded workbook may fill it, but the run must explicitly record that
    # no PDF cross-check was available.
    assert fields["income_approach_value"] not in (None, "", [], {})
    assert fields["asset_approach_value"] not in (None, "", [], {})
    for field in (
        "book_net_assets",
        "historical_balance_sheet_table",
        "historical_income_statement_table",
        "major_long_term_assets",
    ):
        assert fields[field] not in (None, "", [], {})
        assert any(
            f"{field}：审计PDF已上传，但OCR未识别到该字段" in issue
            for issue in result.issues
        )
    assert "节点输入 返回越权字段，已丢弃：company_profile_section" in result.issues
    assert not any("commissioning_party_name" in issue for issue in result.issues)

    report = Document(result.report_path)
    paragraph_text = "\n".join(paragraph.text for paragraph in report.paragraphs)
    assert "单体层面各类资产负债的金额为：货币资金" not in paragraph_text
    assert "单体层面各类资产负债的金额为：" in paragraph_text
    balance_text = "\n".join(cell.text for table in report.tables for row in table.rows for cell in row.cells)
    assert "148,537,259.26" in balance_text
    ownership_text = "\n".join(
        cell.text for table_index in (2, 3) for row in report.tables[table_index].rows for cell in row.cells
    )
    assert "富士和机械工业（昆山）有限公司" not in ownership_text
    assert "上海上大热处理有限公司" not in ownership_text

    with zipfile.ZipFile(result.report_path) as archive:
        xml = "".join(
            archive.read(name).decode("utf-8")
            for name in archive.namelist()
            if re.fullmatch(r"word/(document|header\d+|footer\d+)\.xml", name)
        )
    assert 'w:val="yellow"' in xml
    assert "待人工补充" not in xml
    assert "不应采用" not in xml
    assert not (tmp_path / "生成问题清单.xlsx").exists()
    assert not (tmp_path / "生成问题清单.json").exists()

    manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    assert manifest["yellow_route_version"] == "yellow_routes.v1"
    assert manifest["prompt_version"] == "yellow_narratives.test"


def test_pipeline_uses_node_input_names_before_qichacha_lookup(tmp_path):
    qichacha = TrackingQichachaAdapter()

    run_pipeline(
        project_config=Path("demo/projects/tongfu.yaml"),
        pdf_path=None,
        output_dir=tmp_path,
        ocr_adapter=None,
        llm_adapter=None,
        qichacha_adapter=qichacha,
        source_overrides={
            "audit_pdf": None,
            "reference_report": None,
            "audited_financials": None,
            "income_workbook": None,
            "reporting_workbook": None,
        },
        manual_inputs_override={
            "commissioning_party_name": "委托方有限公司",
            "commissioning_party_short_name": "委托方",
            "target_company_name": "被评估单位有限公司",
            "target_company_short_name": "被评估单位",
            "transaction_type": "收购",
            "valuation_subject_type": "股东全部权益价值",
            "selected_valuation_method": "收益法、资产基础法",
            "final_valuation_method": "收益法",
            "valuation_base_date": "2025-06-30",
            "registry_info_strategy": "qichacha",
        },
    )

    assert qichacha.calls[:2] == ["委托方有限公司", "被评估单位有限公司"]
    fields = json.loads((tmp_path / "normalized_fields.json").read_text(encoding="utf-8"))
    assert "统一社会信用代码：91320000608319749X" in fields["commissioning_party_profile"]


def test_pipeline_does_not_run_llm_reviews_and_exports_four_node_trace(tmp_path):
    result = run_pipeline(
        project_config=Path("demo/projects/tongfu.yaml"),
        pdf_path=Path("资产评估工作流/通富2025.6.30合并及母公司审计报告.pdf"),
        output_dir=tmp_path,
        ocr_adapter=FixtureOcrAdapter(),
        llm_adapter=FixtureLlmAdapter(),
        qichacha_adapter=FixtureQichachaAdapter(),
        ocr_field_resolver=fixture_ocr_fields,
        template_page_reader=FixtureTemplatePageReader(),
    )

    assert not (tmp_path / "格式审核.json").exists()
    assert not (tmp_path / "数据校验.json").exists()
    assert not (tmp_path / "语义审核.json").exists()
    assert not (tmp_path / "审核汇总.json").exists()
    manifest = json.loads((tmp_path / "run_manifest.json").read_text(encoding="utf-8"))
    assert "reviews" not in manifest
    trace_path = tmp_path / "workflow_trace.json"
    assert trace_path.exists()
    trace = json.loads(trace_path.read_text(encoding="utf-8"))
    assert trace["contract_version"] == "workflow_contract.v2"
    assert [node["node_name"] for node in trace["nodes"]] == [
        "start_input",
        "ocr_llm_candidates",
        "fill_word",
        "output",
    ]
    assert str(trace_path) in manifest["outputs"]


def test_pipeline_generates_review_report_when_monetary_fields_are_missing(
    tmp_path,
    monkeypatch,
):
    real_run_project = pipeline_module.run_project

    def run_project_with_missing_fields(*args, **kwargs):
        result = real_run_project(*args, **kwargs)
        fields_path = result.report_path.parent / "normalized_fields.json"
        fields = json.loads(fields_path.read_text(encoding="utf-8"))
        fields["book_net_assets"] = ""
        fields["historical_balance_sheet_table"] = ""
        fields_path.write_text(
            json.dumps(fields, ensure_ascii=False),
            encoding="utf-8",
        )
        return result

    monkeypatch.setattr(
        pipeline_module,
        "run_project",
        run_project_with_missing_fields,
    )

    result = run_pipeline(
        project_config=Path("demo/projects/tongfu.yaml"),
        pdf_path=Path("资产评估工作流/通富2025.6.30合并及母公司审计报告.pdf"),
        output_dir=tmp_path,
        ocr_adapter=FixtureOcrAdapter(),
        llm_adapter=FixtureLlmAdapter(),
        qichacha_adapter=FixtureQichachaAdapter(),
        ocr_field_resolver=fixture_ocr_fields,
        template_page_reader=FixtureTemplatePageReader(),
    )

    assert result.report_path.exists()
    assert not (tmp_path / "资产评估报告_最终候选.docx").exists()
    assert "高优先级：金额及财务结果字段未匹配到，Word已保留黄色占位符：book_net_assets" in result.issues
    fields = json.loads(
        (tmp_path / "normalized_fields.json").read_text(encoding="utf-8")
    )
    assert fields["book_net_assets"] == ""
    report = Document(result.report_path)
    assert all(
        cell.text == "XXX"
        for row in report.tables[4].rows[1:]
        for cell in row.cells[1:]
    )

    manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    assert manifest["financial_validation"]["valid"] is False
    assert set(manifest["financial_validation"]["missing_fields"]) == {
        "book_net_assets",
        "historical_balance_sheet_table",
    }
    trace = json.loads(
        (tmp_path / "workflow_trace.json").read_text(encoding="utf-8")
    )
    fill_word_node = next(
        node for node in trace["nodes"] if node["node_name"] == "fill_word"
    )
    assert fill_word_node["status"] == "completed_with_issues"


def test_pipeline_without_pdf_exports_report_only(tmp_path):
    result = run_pipeline(
        project_config=Path("demo/projects/tongfu.yaml"),
        pdf_path=None,
        output_dir=tmp_path,
        ocr_adapter=None,
        source_overrides={
            "audit_pdf": None,
            "reference_report": None,
            "audited_financials": None,
            "income_workbook": None,
            "reporting_workbook": None,
        },
        manual_inputs_override={
            "target_company_name": "示例有限公司",
            "registry_info_strategy": "qichacha",
        },
    )

    assert result.report_path.exists()
    assert result.ocr_workbook_path is None
    assert not (tmp_path / "生成问题清单.xlsx").exists()
    assert not (tmp_path / "生成问题清单.json").exists()
    manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
    assert manifest["financial_validation"]["valid"] is False
    assert manifest["generation_validation"]["valid"] is False


def test_pipeline_does_not_use_semantic_excel_as_audit_pdf_substitute(tmp_path):
    reporting = tmp_path / "任意资产表.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "汇总表"
    sheet.append(["金额单位：人民币万元"])
    sheet.append(["项目", "账面价值", "评估价值"])
    sheet.append(["流动资产", 100, 110])
    sheet.append(["负债合计", 40, 42])
    sheet.append(["净资产", 60, 68])
    workbook.save(reporting)

    output = tmp_path / "run"
    result = run_pipeline(
        project_config=Path("demo/projects/tongfu.yaml"),
        pdf_path=None,
        output_dir=output,
        ocr_adapter=None,
        source_overrides={
            "audit_pdf": None,
            "reference_report": None,
            "audited_financials": None,
            "income_workbook": None,
            "reporting_workbook": reporting,
        },
        manual_inputs_override={"target_company_name": "示例有限公司"},
    )

    trace = json.loads(
        (output / "workflow_trace.json").read_text(encoding="utf-8")
    )
    assert [node["node_name"] for node in trace["nodes"]] == [
        "start_input", "ocr_llm_candidates", "fill_word", "output"
    ]
    normalized_evidence = json.loads(
        (output / "normalized_evidence.json").read_text(encoding="utf-8")
    )
    # A missing PDF does not discard uniquely matched uploaded workbook data;
    # its evidence carries the workbook source so the Word note can state that
    # a PDF cross-check has not been performed.
    assert normalized_evidence["asset_scope_summary_table"]["file"] == reporting.name
    assert normalized_evidence["asset_scope_summary_table"]["kind"] == "asset_workbook"
    manifest = json.loads(
        (output / "run_manifest.json").read_text(encoding="utf-8")
    )
    assert str(output / "normalized_evidence.json") in manifest["outputs"]


def test_pipeline_keeps_unfinished_appraisal_reason_in_issue_list(
    tmp_path,
    monkeypatch,
):
    reporting = tmp_path / "尚未完成评估.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "汇总表"
    sheet.append(["金额单位：人民币万元"])
    sheet.append(["项目", "账面价值", "评估价值"])
    sheet.append(["资产总计", 100, 0])
    sheet.append(["负债合计", 40, 0])
    sheet.append(["净资产", 60, 0])
    workbook.save(reporting)

    def page_mapping(records, page_texts, paragraph_texts):
        location_ids = {
            str(item.get("location_id", ""))
            for item in records
            if isinstance(item, dict)
        }
        return (
            {"DOCUMENT-P0558-H01": 16}
            if "DOCUMENT-P0558-H01" in location_ids
            else {}
        )

    monkeypatch.setattr(
        pipeline_module,
        "map_location_pages",
        page_mapping,
    )
    output = tmp_path / "run"
    result = run_pipeline(
        project_config=Path("demo/projects/tongfu.yaml"),
        pdf_path=None,
        output_dir=output,
        ocr_adapter=None,
        source_overrides={
            "audit_pdf": None,
            "reference_report": None,
            "audited_financials": None,
            "income_workbook": None,
            "reporting_workbook": reporting,
        },
        manual_inputs_override={"target_company_name": "示例有限公司"},
        template_page_reader=FixtureTemplatePageReader(),
    )

    assert any("尚未完成评估" in issue for issue in result.issues)


def test_pipeline_without_pdf_passes_material_fields_to_llm(tmp_path):
    output = tmp_path / "run"
    result = run_pipeline(
        project_config=Path("demo/projects/tongfu.yaml"),
        pdf_path=None,
        output_dir=output,
        ocr_adapter=None,
        llm_adapter=NoPdfEvidenceLlmAdapter(),
        source_overrides={
            "audit_pdf": None,
            "reference_report": None,
            "audited_financials": None,
            "income_workbook": None,
            "reporting_workbook": None,
        },
        manual_inputs_override={
            "target_company_name": "示例有限公司",
            "narrative_modules": ["industry_overview"],
        },
    )

    fields = json.loads((output / "normalized_fields.json").read_text(encoding="utf-8"))
    assert fields["company_profile_section"] == "根据已上传结构化材料生成。"
    assert fields["industry_overview"] == ""


def test_pipeline_passes_uploaded_reference_word_to_llm(tmp_path):
    reference = tmp_path / "任意名称的参考报告.docx"
    document = Document()
    document.add_paragraph("4.2、主要产品")
    document.add_paragraph("该公司主营产品为工业滤波器，并面向工业自动化客户销售。")
    document.save(reference)
    output = tmp_path / "run"

    run_pipeline(
        project_config=Path("demo/projects/tongfu.yaml"),
        pdf_path=None,
        output_dir=output,
        ocr_adapter=None,
        llm_adapter=ReferenceDocumentEvidenceLlmAdapter(),
        source_overrides={
            "audit_pdf": None,
            "reference_report": reference,
            "audited_financials": None,
            "income_workbook": None,
            "reporting_workbook": None,
        },
        manual_inputs_override={
            "target_company_name": "示例有限公司",
            "narrative_modules": ["main_products"],
        },
    )

    fields = json.loads((output / "normalized_fields.json").read_text(encoding="utf-8"))
    assert fields["main_products"] == "主营产品为工业滤波器。"


def test_qichacha_profile_supplies_missing_registered_capital(tmp_path):
    output = tmp_path / "run"
    run_pipeline(
        project_config=Path("demo/projects/tongfu.yaml"),
        pdf_path=None,
        output_dir=output,
        ocr_adapter=None,
        qichacha_adapter=CapitalQichachaAdapter(),
        source_overrides={
            "audit_pdf": None,
            "reference_report": None,
            "audited_financials": None,
            "income_workbook": None,
            "reporting_workbook": None,
        },
        manual_inputs_override={
            "target_company_name": "示例有限公司",
            "registry_info_strategy": "qichacha",
        },
    )

    fields = json.loads((output / "normalized_fields.json").read_text(encoding="utf-8"))
    assert fields["registered_capital"] == "1,250万元"


def test_pipeline_keeps_semantic_excel_scope_as_reconciliation_only_without_pdf(tmp_path):
    reporting = tmp_path / "changed-layout.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "汇总表"
    sheet.append(["金额单位：人民币万元"])
    sheet.append(["项目", "序号", "账面价值", "评估价值"])
    sheet.append(["流动资产", 1, 100, 101])
    sheet.append(["非流动资产", 2, 200, 210])
    sheet.append(["固定资产", 21, 120, 130])
    sheet.append(["无形资产", 22, 30, 35])
    sheet.append(["长期待摊费用", 23, 10, 10])
    sheet.append(["资产总计", 3, 300, 311])
    sheet.append(["流动负债", 4, 50, 50])
    sheet.append(["非流动负债", 5, 20, 20])
    sheet.append(["负债总计", 6, 70, 70])
    sheet.append(["净资产", 7, 230, 241])
    detail = workbook.create_sheet("固定资产汇总表")
    detail.append(["金额单位：人民币元"])
    detail.append(["项目", "账面价值", "评估价值"])
    detail.append(["电子设备", 320_000, 350_000])
    workbook.save(reporting)

    result = run_pipeline(
        project_config=Path("demo/projects/tongfu.yaml"),
        pdf_path=None,
        output_dir=tmp_path / "run",
        ocr_adapter=None,
        source_overrides={
            "audit_pdf": None,
            "reference_report": None,
            "audited_financials": reporting,
            "income_workbook": None,
            "reporting_workbook": reporting,
        },
        manual_inputs_override={"target_company_name": "示例有限公司"},
    )

    fields = json.loads(
        (tmp_path / "run/normalized_fields.json").read_text(encoding="utf-8")
    )
    assert fields["asset_scope_summary_table"] not in (None, "", [], {})
    assert any(
        "asset_scope_summary_table：未上传审计PDF，无法获取审计数据" in issue
        for issue in result.issues
    )
    long_term_text = "\n".join(
        cell.text
        for row in Document(tmp_path / "run/资产评估报告_待复核.docx").tables[7].rows
        for cell in row.cells
    )
    assert "320,000.00" in long_term_text


def test_pipeline_does_not_read_legacy_coordinates_without_semantic_evidence(tmp_path):
    """A workbook without semantic tables must not fall back to old cells."""
    workbook_path = tmp_path / "arbitrary-layout.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "06N_资产负债表"
    # Populate the coordinates used by the old Tongfu configuration with a
    # deliberately misleading value.  No semantic financial table is present.
    sheet["F28"] = 999
    sheet["F49"] = 999
    sheet["F76"] = 999
    sheet["L32"] = 999
    sheet["L46"] = 999
    sheet["L75"] = 999
    detail = workbook.create_sheet("表4-6")
    detail["E15"] = 999
    workbook.save(workbook_path)

    output = tmp_path / "run"
    run_pipeline(
        project_config=Path("demo/projects/tongfu.yaml"),
        pdf_path=None,
        output_dir=output,
        ocr_adapter=None,
        source_overrides={
            "audit_pdf": None,
            "reference_report": None,
            "audited_financials": workbook_path,
            "income_workbook": None,
            "reporting_workbook": workbook_path,
        },
        manual_inputs_override={"target_company_name": "示例有限公司"},
    )

    fields = json.loads((output / "normalized_fields.json").read_text(encoding="utf-8"))
    assert fields["asset_scope_summary_table"] == ""
    generated = Document(output / "资产评估报告_待复核.docx")
    long_term_rows = [
        [cell.text.strip() for cell in row.cells]
        for row in generated.tables[7].rows
    ]
    assert all(row[1] == "XXX" for row in long_term_rows[1:])


def test_pipeline_rejects_invalid_workflow_before_ocr(tmp_path):
    invalid_workflow = tmp_path / "invalid-workflow.json"
    invalid_workflow.write_text(
        json.dumps(
            {
                "version": "test",
                "contract_version": "workflow_contract.v1",
                "nodes": [
                    {
                        "name": "ocr_pdf",
                        "input_model": "MissingInput",
                        "output_model": "OcrPdfOutput",
                        "depends_on": [],
                        "human_checkpoint": None,
                    }
                ],
            }
        ),
        encoding="utf-8",
    )

    class FailIfCalledOcr:
        def extract(self, pdf_path):
            raise AssertionError("invalid workflow must stop before OCR")

    with pytest.raises(ValueError, match="工作流契约校验失败"):
        run_pipeline(
            project_config=Path("demo/projects/tongfu.yaml"),
            pdf_path=Path("资产评估工作流/通富2025.6.30合并及母公司审计报告.pdf"),
            output_dir=tmp_path / "run",
            ocr_adapter=FailIfCalledOcr(),
            workflow_path=invalid_workflow,
        )


def test_pipeline_keeps_review_report_when_cloud_ocr_fails(tmp_path):
    class FailingCloudOcr:
        def extract(self, _pdf_path):
            return [], ["阿里云 OCR 提交失败：NoPermission"]

    result = run_pipeline(
        project_config=Path("demo/projects/tongfu.yaml"),
        pdf_path=Path(
            "资产评估工作流/通富2025.6.30合并及母公司审计报告.pdf"
        ),
        output_dir=tmp_path,
        ocr_adapter=FailingCloudOcr(),
        template_page_reader=FixtureTemplatePageReader(),
    )

    assert result.report_path.exists()
    assert "阿里云 OCR 提交失败：NoPermission" in result.issues
    assert not (tmp_path / "生成问题清单.xlsx").exists()


def test_pipeline_only_fills_selected_narrative_modules(tmp_path):
    result = run_pipeline(
        project_config=Path("demo/projects/tongfu.yaml"),
        pdf_path=Path("资产评估工作流/通富2025.6.30合并及母公司审计报告.pdf"),
        output_dir=tmp_path,
        ocr_adapter=FixtureOcrAdapter(),
        llm_adapter=FixtureLlmAdapter(),
        qichacha_adapter=FixtureQichachaAdapter(),
        ocr_field_resolver=fixture_ocr_fields,
        template_page_reader=FixtureTemplatePageReader(),
        manual_inputs_override={"narrative_modules": ["main_products"]},
    )

    fields = json.loads((tmp_path / "normalized_fields.json").read_text(encoding="utf-8"))
    assert fields["main_products"] == "热处理服务。"
    assert fields["industry_overview"] == ""
    assert fields["customers_suppliers"] == ""
