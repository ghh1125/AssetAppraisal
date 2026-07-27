from demo.domain.source_precedence import prefer_semantic_result


def test_semantic_source_replaces_fixed_coordinate_result():
    selected = prefer_semantic_result(
        fixed_value=1,
        fixed_evidence={
            "kind": "reporting_workbook",
            "file": "旧表.xlsx",
            "locator": "旧表!A1",
        },
        semantic_value=2,
        semantic_evidence={
            "kind": "semantic_excel",
            "file": "新表.xlsx",
            "locator": "汇总表!D22",
        },
    )

    assert selected == {
        "value": 2,
        "evidence": {
            "kind": "semantic_excel",
            "file": "新表.xlsx",
            "locator": "汇总表!D22",
        },
    }


def test_fixed_coordinate_is_used_when_semantic_value_is_missing():
    selected = prefer_semantic_result(
        fixed_value=1,
        fixed_evidence={
            "kind": "reporting_workbook",
            "file": "旧表.xlsx",
            "locator": "旧表!A1",
        },
        semantic_value=None,
        semantic_evidence={},
    )

    assert selected["value"] == 1
    assert selected["evidence"]["locator"] == "旧表!A1"


def test_semantic_empty_container_does_not_erase_fixed_result():
    selected = prefer_semantic_result(
        fixed_value={"rows": [["项目", "1"]]},
        fixed_evidence={"kind": "reporting_workbook"},
        semantic_value={},
        semantic_evidence={"kind": "semantic_excel"},
    )

    assert selected["value"] == {"rows": [["项目", "1"]]}
