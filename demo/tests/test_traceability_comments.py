from demo.domain.traceability_comments import (
    build_traceability_comment_request,
    validate_traceability_comment_response,
)


def test_traceability_comment_request_groups_repeated_cells_without_losing_source():
    request, index_map = build_traceability_comment_request(
        [
            {
                "field_key": "historical_income_table",
                "field_name": "历史利润表 / 营业收入",
                "status": "verified",
                "comment": "【来源已核验】来源：《审计报告.pdf》第28页。LLM复核通过。",
            },
            {
                "field_key": "historical_income_table",
                "field_name": "历史利润表 / 净利润",
                "status": "verified",
                "comment": "【来源已核验】来源：《审计报告.pdf》第28页。LLM复核通过。",
            },
        ]
    )

    assert len(request["comments"]) == 1
    assert request["comments"][0]["required_title"] == "【来源已核验】"
    assert "审计报告.pdf" in request["comments"][0]["evidence_summary"]
    assert index_map == {"trace-1": [0, 1]}


def test_traceability_comment_response_rejects_unknown_ids_and_strips_model_title():
    comments, issues = validate_traceability_comment_response(
        {
            "comments": [
                {"comment_id": "trace-1", "comment": "【审核通过】来源于审计报告第28页。"},
                {"comment_id": "trace-x", "comment": "越权内容"},
            ]
        },
        allowed_comment_ids={"trace-1"},
    )

    assert comments == {"trace-1": "来源于审计报告第28页。"}
    assert any("未知编号" in issue for issue in issues)
