from pathlib import Path

from demo.adapters.word import extract_word_comments, inventory_template
from demo.domain.comment_mapping import build_comment_aware_locations
import json


ROOT = Path(__file__).resolve().parents[1]


def test_comment_annotations_are_retained_on_word_locations() -> None:
    template = Path("/Users/ghh/Downloads/评估报告版式-沟通标注版_v2.docx")
    if not template.is_file():
        template = ROOT.parent / "templates/评估报告版式-沟通标注版.docx"
    locations = inventory_template(template)
    first = next(item for item in locations if item["location_id"] == "DOCUMENT-P0005-X01")
    if "comment_texts" not in first:
        raise AssertionError("inventory records must expose comment metadata")
    if template.name.endswith("_v2.docx"):
        assert any("委托方" in text for text in first["comment_texts"])
        assert len(extract_word_comments(template)) == 100


def test_legacy_highlight_template_remains_comment_compatible() -> None:
    locations = inventory_template(ROOT.parent / "templates/评估报告版式-沟通标注版.docx")
    assert locations
    assert all(isinstance(item.get("comment_texts", []), list) for item in locations)


def test_comment_template_maps_all_placeholder_locations_to_business_fields() -> None:
    template = Path("/Users/ghh/Downloads/评估报告版式-沟通标注版_v2.docx")
    if not template.is_file():
        return
    base = json.loads((ROOT / "mappings/appraisal_report_v1.yaml").read_text(encoding="utf-8"))["locations"]
    mapped = build_comment_aware_locations(inventory_template(template), base)
    assert len(mapped) == 132
    assert not [item for item in mapped if not item.get("field_key")]


def test_comment_template_maps_repeated_placeholders_in_order() -> None:
    template = Path("/Users/ghh/Downloads/评估报告版式-沟通标注版_v2.docx")
    if not template.is_file():
        return
    base = json.loads((ROOT / "mappings/appraisal_report_v1.yaml").read_text(encoding="utf-8"))["locations"]
    mapped = {
        item["location_id"]: item["field_key"]
        for item in build_comment_aware_locations(inventory_template(template), base)
    }
    assert [mapped[f"DOCUMENT-P0323-X{i:02d}"] for i in range(1, 5)] == [
        "commissioning_party_name", "target_company_name",
        "target_company_name", "valuation_subject_type",
    ]
    assert [mapped[f"DOCUMENT-P0550-X{i:02d}"] for i in range(1, 6)] == [
        "target_company_name", "book_net_assets", "income_approach_value",
        "income_increment", "income_increment_rate",
    ]
    assert [mapped[f"DOCUMENT-P0587-X{i:02d}"] for i in range(1, 4)] == [
        "report_date_year", "report_date_month", "report_date_day",
    ]
