import json
import tomllib
from pathlib import Path

from demo.adapters.paddle_ocr import PaddleStructureOcrAdapter


class FakeResult:
    json = {
        "res": {
            "page_index": 0,
            "page_count": 1,
            "overall_ocr_res": {
                "rec_texts": ["资产总计"],
                "rec_scores": [0.9],
                "rec_boxes": [[1, 2, 30, 12]],
            },
            "table_res_list": [
                {
                    "table_ocr_pred": {
                        "rec_texts": ["项目", "期末余额", "资产总计", "1,234.50"],
                        "rec_scores": [0.99, 0.98, 0.97, 0.96],
                        "rec_boxes": [
                            [1, 20, 30, 30],
                            [40, 20, 80, 30],
                            [1, 35, 30, 45],
                            [40, 35, 80, 45],
                        ],
                    }
                }
            ],
        }
    }


class FakePipeline:
    def predict(self, input):
        assert isinstance(input, str)
        assert Path(input).name == "scan.pdf"
        yield FakeResult()


def test_paddle_adapter_returns_serializable_pages(tmp_path):
    pdf = tmp_path / "scan.pdf"
    pdf.write_bytes(b"%PDF-1.4")

    pages, issues = PaddleStructureOcrAdapter(FakePipeline()).extract(pdf)

    assert not issues
    assert pages[0]["page_number"] == 1
    assert pages[0]["blocks"][0]["text"] == "资产总计"
    assert pages[0]["tables"][0]["cells"][3]["row"] == 2
    assert pages[0]["tables"][0]["cells"][3]["column"] == 2
    assert pages[0]["tables"][0]["cells"][3]["text"] == "1,234.50"
    json.dumps(pages, ensure_ascii=False)


def test_paddle_adapter_reports_failure_without_leaking_exception(tmp_path):
    class BrokenPipeline:
        def predict(self, input):
            raise RuntimeError("bad page")

    pages, issues = PaddleStructureOcrAdapter(BrokenPipeline()).extract(tmp_path / "scan.pdf")

    assert pages == []
    assert issues == ["PaddleOCR 失败：bad page"]


def test_ocr_extra_installs_pp_structure_runtime_dependencies():
    project = tomllib.loads(Path("pyproject.toml").read_text(encoding="utf-8"))
    dependencies = project["project"]["optional-dependencies"]["ocr"]
    assert any(item.startswith("paddlex[ocr]") for item in dependencies)
