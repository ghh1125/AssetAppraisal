import json
from pathlib import Path

import pytest

from demo.adapters.word import inventory_template
from demo.domain.yellow_routing import RouteKind, load_yellow_routes, validate_yellow_routes


EXPECTED_FIELDS = {
    RouteKind.NODE_INPUT: {
        "selected_valuation_method",
        "valuation_purpose_inputs",
    },
    RouteKind.QICHACHA_API: {
        "commissioning_party_profile",
        "ownership_history",
        "ownership_at_valuation_date",
        "unrecorded_intangibles",
        "software_copyrights",
    },
    RouteKind.PDF_OCR_XLSX: {
        "historical_balance_sheet_table",
        "historical_income_statement_table",
        "tax_rates",
        "valuation_scope",
        "major_long_term_assets",
        "asset_approach_result_section",
    },
    RouteKind.BAILIAN_GLM: {
        "company_profile_section",
        "industry_overview",
        "business_and_segments",
        "main_products",
        "customers_suppliers",
        "profit_model_swot",
        "comparable_list",
    },
}


def test_project_config_locks_every_yellow_location_to_exactly_one_source():
    config_path = Path("demo/projects/tongfu.yaml")
    project = json.loads(config_path.read_text(encoding="utf-8"))
    routes = load_yellow_routes(project["yellow_routes"])

    template = config_path.parent / project["template"]
    yellow_ids = {
        item["location_id"]
        for item in inventory_template(template)
        if item["record_type"] == "黄色标注内容块"
    }

    validate_yellow_routes(routes, expected_location_ids=yellow_ids)
    assert len(routes) == len(yellow_ids) == 20
    for route_kind, expected in EXPECTED_FIELDS.items():
        assert {route.field_key for route in routes if route.route_kind == route_kind} == expected


def test_route_validation_rejects_duplicate_location_and_field():
    duplicated = [
        {
            "location_id": "DOCUMENT-P0001-H01",
            "field_key": "field_a",
            "route_kind": "pdf_ocr_xlsx",
        },
        {
            "location_id": "DOCUMENT-P0001-H01",
            "field_key": "field_a",
            "route_kind": "bailian_glm",
        },
    ]

    with pytest.raises(ValueError, match="重复"):
        load_yellow_routes(duplicated)


def test_route_validation_rejects_missing_or_extra_template_location():
    routes = load_yellow_routes(
        [
            {
                "location_id": "DOCUMENT-P0001-H01",
                "field_key": "field_a",
                "route_kind": "node_input",
            }
        ]
    )

    with pytest.raises(ValueError, match="不一致"):
        validate_yellow_routes(routes, expected_location_ids={"DOCUMENT-P0002-H01"})
