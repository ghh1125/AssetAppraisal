import json
from pathlib import Path

from demo.adapters.word import inventory_template


def test_template_and_mapping_cover_147_locations():
    template = Path("资产评估工作流/评估报告版式-沟通标注版.docx")
    inventory = inventory_template(template)
    mapping = json.loads(Path("demo/mappings/appraisal_report_v1.yaml").read_text(encoding="utf-8"))
    records = mapping["locations"]
    static_records = mapping["static_locations"]
    assert len(inventory) == 147
    assert sum(x["record_type"] == "占位符" for x in inventory) == 127
    assert sum(x["record_type"] == "黄色标注内容块" for x in inventory) == 20
    assert sum(x["part"].startswith("word/footer") for x in inventory) == 2
    assert {x["location_id"] for x in inventory} == {x["location_id"] for x in records}
    assert {x["location_id"] for x in static_records} == {"DOCUMENT-P0380-S01"}
