from demo.domain.template_pagination import map_location_pages


def test_maps_repeated_contexts_monotonically_to_original_template_pages():
    locations = [
        {"location_id": "DOCUMENT-P0001-X01", "context": "甲公司"},
        {"location_id": "DOCUMENT-P0002-X01", "context": "唯一中间段"},
        {"location_id": "DOCUMENT-P0003-X01", "context": "甲公司"},
        {"location_id": "FOOTER4-P0001-X01", "context": "页脚"},
    ]

    pages = map_location_pages(locations, ["甲公司", "唯一中间段", "甲公司"])

    assert pages == {
        "DOCUMENT-P0001-X01": 1,
        "DOCUMENT-P0002-X01": 2,
        "DOCUMENT-P0003-X01": 3,
        "FOOTER4-P0001-X01": "多页页脚",
    }
