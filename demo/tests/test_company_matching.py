from demo.domain.company_matching import company_match_score, matching_company_records


def test_exact_legal_name_has_priority_over_filename_fallback():
    records = [
        {"name": "上海明悦医疗科技有限公司", "source": "未命名审计报告.pdf"},
        {"name": "", "source": "上海明悦-其他材料.pdf"},
    ]

    matches = matching_company_records(
        records,
        "上海明悦医疗科技有限公司",
        name_getter=lambda item: item["name"],
        source_getter=lambda item: item["source"],
    )

    assert matches == [records[0]]


def test_filename_abbreviation_matches_when_ocr_name_is_missing():
    assert company_match_score(
        "上海明悦医疗科技有限公司",
        "",
        "00审计报告-测试/2年一期/赋码-上自贸-2026-0214-上海明悦.pdf",
    ) > 0


def test_filename_fallback_does_not_match_another_company():
    assert company_match_score(
        "上海明悦医疗科技有限公司",
        "",
        "赋码-上自贸-2026-0297北京嘉云升.pdf",
    ) == 0


def test_short_generic_overlap_is_not_accepted():
    assert company_match_score(
        "明悦科技有限公司",
        "",
        "另一家科技有限公司审计报告.pdf",
    ) == 0
