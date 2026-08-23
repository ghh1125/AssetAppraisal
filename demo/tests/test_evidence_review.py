from demo.domain.evidence_review import (
    build_evidence_review_request,
    needs_review_fallback,
    validate_evidence_review_response,
)
from demo.pipeline import _llm_review_rows


def test_evidence_review_request_contains_selected_value_and_all_source_candidates():
    request = build_evidence_review_request(
        [
            {
                "field_key": "book_net_assets",
                "selected": {
                    "value": "100.00",
                    "source_kind": "pdf_ocr",
                    "source_file": "审计报告.pdf",
                    "source_locator": "第10页",
                },
                "candidates": [
                    {
                        "value": "100.00",
                        "source_kind": "pdf_ocr",
                        "source_file": "审计报告.pdf",
                        "source_locator": "第10页",
                    },
                    {
                        "value": "90.00",
                        "source_kind": "asset_workbook",
                        "source_file": "资产基础法.xlsx",
                        "source_locator": "汇总表!C8",
                    },
                ],
            }
        ],
        {"book_net_assets": "账面净资产"},
    )

    assert request["version"] == "evidence_review.v1"
    assert request["fields"] == [
        {
            "field_key": "book_net_assets",
            "field_name": "账面净资产",
            "selected": {
                "value": "100.00",
                "source_kind": "pdf_ocr",
                "source_file": "审计报告.pdf",
                "source_locator": "审计 PDF 第10页：账面净资产",
            },
            "candidates": [
                {
                    "value": "100.00",
                    "source_kind": "pdf_ocr",
                    "source_file": "审计报告.pdf",
                    "source_locator": "审计 PDF 第10页：账面净资产",
                },
                {
                    "value": "90.00",
                    "source_kind": "asset_workbook",
                    "source_file": "资产基础法.xlsx",
                    "source_locator": "《资产基础法.xlsx》工作表“汇总表”（账面净资产）",
                },
            ],
            "review_required": True,
        }
    ]


def test_evidence_review_request_uses_readable_ocr_table_name():
    request = build_evidence_review_request(
        [
            {
                "field_key": "asset_scope_summary_table",
                "selected": {
                    "value": {"rows": [["项目", "金额"], ["资产合计", "1"]]},
                    "source_kind": "pdf_ocr",
                    "source_file": "审计报告.pdf",
                    "source_locator": "OCR结构化结果.xlsx / asset_scope_summary_table",
                },
                "candidates": [],
            }
        ],
        {"asset_scope_summary_table": "资产负债范围表"},
    )

    assert request["fields"][0]["selected"]["source_locator"] == "审计 PDF：资产负债范围表"


def test_evidence_review_rejects_unknown_fields_and_invalid_statuses():
    reviews, issues = validate_evidence_review_response(
        {
            "reviews": [
                {
                    "field_key": "book_net_assets",
                    "status": "needs_review",
                    "reason": "期间标签不一致，需人工核对。",
                    "comment": "请核对《资产基础法.xlsx》“汇总表”C8与审计报告第10页的期间和口径。",
                },
                {
                    "field_key": "invented",
                    "status": "accept",
                    "reason": "不应保留。",
                },
                {
                    "field_key": "book_net_assets",
                    "status": "replace_value",
                    "reason": "模型不能改数。",
                },
            ]
        },
        allowed_field_keys=["book_net_assets"],
    )

    assert reviews == [
        {
            "field_key": "book_net_assets",
            "status": "needs_review",
            "reason": "期间标签不一致，需人工核对。",
            "comment": "请核对《资产基础法.xlsx》“汇总表”C8与审计报告第10页的期间和口径。",
        }
    ]
    assert len(issues) == 2


def test_needs_review_fallback_preserves_human_readable_source():
    review = needs_review_fallback(
        {
            "field_key": "book_net_assets",
            "field_name": "账面净资产",
            "selected": {
                "source_file": "审计报告.pdf",
                "source_locator": "审计 PDF 第10页：资产负债表",
            },
        }
    )

    assert review["status"] == "needs_review"
    assert review["field_key"] == "book_net_assets"
    assert "《审计报告.pdf》" in review["comment"]
    assert "第10页" in review["comment"]


def test_llm_review_rows_include_every_resolved_pdf_or_workbook_field_once():
    rows = _llm_review_rows(
        {"book_net_assets": "100.00", "income_approach_value": "120.00", "transaction_type": "收购"},
        {
            "book_net_assets": {
                "kind": "pdf_ocr_xlsx",
                "file": "审计报告.pdf",
                "locator": "第10页",
            },
            "income_approach_value": {
                "kind": "income_workbook",
                "file": "收益法.xlsx",
                "locator": "汇总表!D8",
            },
            "transaction_type": {"kind": "node_input", "file": "", "locator": ""},
        },
        [
            {
                "field_key": "book_net_assets",
                "selected": {"value": "100.00", "source_kind": "pdf_ocr", "source_file": "审计报告.pdf", "source_locator": "第10页"},
                "candidates": [],
            }
        ],
    )

    assert [item["field_key"] for item in rows] == ["book_net_assets", "income_approach_value"]
