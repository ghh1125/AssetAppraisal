from demo.domain.source_labels import human_source_locator, locate_pdf_field_page


def test_excel_locator_uses_uploaded_filename_and_human_sheet_name():
    assert human_source_locator(
        "收益法输入.xlsx",
        "历利表!A6；历利表!C6；历利表!E46",
        "历史利润表",
    ) == "《收益法输入.xlsx》工作表“历利表”（历史利润表）"


def test_ocr_locator_uses_the_original_pdf_page_when_table_markers_match():
    normalized = {
        "text_blocks": [
            {"page_number": 7, "text": "审计报告正文"},
            {"page_number": 18, "text": "资产负债表"},
        ],
        "table_cells": [
            {"page_number": 18, "table_id": "t1", "row": 1, "column": 1, "text": "总资产"},
            {"page_number": 18, "table_id": "t1", "row": 2, "column": 1, "text": "负债"},
            {"page_number": 18, "table_id": "t1", "row": 3, "column": 1, "text": "所有者权益"},
        ],
    }

    assert locate_pdf_field_page(
        normalized,
        "historical_balance_sheet_table",
        "历史资产负债表",
    ) == "审计 PDF 第18页：历史资产负债表"


def test_ocr_locator_uses_financial_data_evidence_id_to_find_pdf_page():
    normalized = {
        "financial_data": [
            {
                "field_key": "historical_income_statement_table",
                "evidence_id": "p22-t3!A1:C12",
            }
        ],
        "table_cells": [
            {"page_number": 22, "table_id": "p22-t3", "row": 1, "column": 1, "text": "利润表"},
        ],
        "text_blocks": [],
    }

    assert locate_pdf_field_page(
        normalized,
        "historical_income_statement_table",
        "历史利润表",
    ) == "审计 PDF 第22页：历史利润表"


def test_ocr_locator_does_not_invent_a_page_without_enough_evidence():
    assert locate_pdf_field_page(
        {"text_blocks": [], "table_cells": []},
        "historical_income_statement_table",
        "历史利润表",
    ) == "审计 PDF：历史利润表"


def test_pdf_locator_accepts_a_single_strong_table_marker():
    normalized = {
        "text_blocks": [],
        "table_cells": [
            {"page_number": 31, "table_id": "p31-t8", "row": 1, "column": 1, "text": "电子设备"},
            {"page_number": 31, "table_id": "p31-t8", "row": 2, "column": 1, "text": "账面净值"},
        ],
    }

    assert locate_pdf_field_page(
        normalized,
        "long_term_assets_table",
        "主要长期资产账面记录表",
    ) == "审计 PDF 第31页：主要长期资产账面记录表"


def test_pdf_locator_extracts_page_from_pdf_evidence_id_even_without_table_cells():
    normalized = {
        "financial_data": [
            {
                "field_key": "asset_scope_summary_table",
                "evidence_id": "pdf:p44:t2:r1:c1",
            }
        ],
        "text_blocks": [],
        "table_cells": [],
    }

    assert locate_pdf_field_page(
        normalized,
        "asset_scope_summary_table",
        "资产负债范围表",
    ) == "审计 PDF 第44页：资产负债范围表"


def test_pdf_source_label_never_exposes_internal_unlocated_fallback():
    assert human_source_locator(
        "审计报告.pdf",
        "asset_scope_summary_table",
        "资产负债范围表",
    ) == "审计 PDF：资产负债范围表"
