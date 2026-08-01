from pathlib import Path

from openpyxl import Workbook

from demo.adapters.company_api import CompanyApiAdapter, QichachaApiAdapter
from demo.adapters.excel import read_cells, try_read_cells
from demo.adapters.llm import LlmAdapter
from demo.adapters.ocr import OcrAdapter


def test_excel_adapter_reads_exact_cells_without_modifying_source(tmp_path: Path):
    path = tmp_path / "source.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "项目信息"
    ws["B5"] = "示例有限公司"
    wb.save(path)
    before = path.read_bytes()
    assert read_cells(path, ["项目信息!B5"])["项目信息!B5"] == "示例有限公司"
    assert path.read_bytes() == before


def test_optional_excel_cell_returns_issue_instead_of_raising(tmp_path: Path):
    workbook_path = tmp_path / "other-layout.xlsx"
    workbook = Workbook()
    workbook.active.title = "资产负债表"
    workbook.save(workbook_path)

    values, issues = try_read_cells(
        workbook_path,
        ["06N_资产负债表!L75"],
    )

    assert values == {}
    assert "缺少工作表" in issues[0]


def test_optional_adapters_degrade_without_clients_or_credentials(tmp_path: Path):
    assert OcrAdapter().extract(tmp_path / "missing.pdf")[0] == []
    assert CompanyApiAdapter().fetch("示例有限公司")[0] == {}
    assert LlmAdapter().generate({"name": "示例有限公司"})[0] == {}


class _Response:
    def __init__(self, payload):
        self.payload = payload

    def raise_for_status(self):
        return None

    def json(self):
        return self.payload


class _Client:
    def __init__(self):
        self.calls = []

    def get(self, url, **kwargs):
        self.calls.append((url, kwargs))
        if url.endswith("ECIInfoVerify/GetInfo"):
            return _Response({"Status": "200", "Data": {"Name": "示例有限公司", "Partners": [{"StockName": "甲", "StockPercent": "60%"}], "ChangeRecords": [{"ChangeDate": "2024-01-01", "ProjectName": "股东变更", "BeforeContent": "甲", "AfterContent": "乙"}]}})
        if url.endswith("tm/SearchByApplicant"):
            return _Response({"Status": "200", "Result": [{"Name": "示例商标", "RegNo": "123", "Category": "35", "FlowStatusDesc": "注册"}]})
        if url.endswith("PatentV4/Search") or url.endswith("PatentV4/SearchMultiPatents"):
            return _Response({"Status": "200", "Data": [{"Title": "示例专利", "PatentType": "发明", "LegalStatus": "有效"}]})
        return _Response({"Status": "200", "Result": [{"Name": "示例软件", "RegisterNo": "2024SR0001", "VersionNo": "V1.0"}]})


def test_qichacha_adapter_signs_and_maps_only_authorized_fields():
    client = _Client()
    adapter = QichachaApiAdapter(client, "app", "secret", extra_api_codes=())
    payload, issues = adapter.fetch("示例有限公司")
    assert not issues
    assert set(payload["fields"]) == {
        "commissioning_party_profile", "ownership_history", "ownership_at_valuation_date",
        "unrecorded_intangibles", "software_copyrights",
    }
    assert "统一社会信用代码" not in payload["fields"]["commissioning_party_profile"]
    assert "示例专利" in payload["fields"]["unrecorded_intangibles"]
    assert "示例商标" in payload["fields"]["unrecorded_intangibles"]
    assert "示例软件" in payload["fields"]["software_copyrights"]
    assert len(client.calls) == 4
    first_url, first_kwargs = client.calls[0]
    assert first_url.endswith("ECIInfoVerify/GetInfo")
    timespan = first_kwargs["headers"]["Timespan"]
    assert first_kwargs["headers"]["Token"] == QichachaApiAdapter.token("app", timespan, "secret")


def test_graphical_trademark_uses_text_name_and_marks_software_query_successfully():
    class GraphicClient(_Client):
        def get(self, url, **kwargs):
            if url.endswith("tm/SearchByApplicant"):
                return _Response({"Status": "200", "Result": [{"ImageUrl": "https://example.test/mark.png", "RegNo": "1"}]})
            if url.endswith("CopyRight/SearchCopyRight"):
                return _Response({"Status": "200", "Result": []})
            return super().get(url, **kwargs)

    payload, issues = QichachaApiAdapter(GraphicClient(), "app", "secret", extra_api_codes=()).fetch("示例有限公司")
    assert not issues
    assert payload["trademark_rows"][0]["name"] == "图形"
    assert payload["trademark_rows"][0]["image"].startswith("https://")
    assert payload["software_query_ok"] is True


def test_qichacha_non_review_apis_are_default_and_expose_evidence():
    class ExtendedClient(_Client):
        def get(self, url, **kwargs):
            if url.endswith("EnterpriseInfo/Verify"):
                self.calls.append((url, kwargs))
                return _Response({"Status": "200", "Data": {"Name": "示例有限公司", "QccIndustry": "电子元件", "Scope": "芯片封装测试"}})
            if url.endswith("AR/GetAnnualReport"):
                self.calls.append((url, kwargs))
                return _Response({"Status": "200", "Data": [{"Year": "2025", "Asset": "100"}]})
            return super().get(url, **kwargs)

    client = ExtendedClient()
    adapter = QichachaApiAdapter(client, "app", "secret")
    payload, issues = adapter.fetch("示例有限公司")
    assert not issues
    assert {item["api_code"] for item in payload["evidence"]} >= {"2001", "213"}
    assert len(client.calls) == 6


def test_qichacha_default_does_not_call_optional_paid_apis():
    client = _Client()
    payload, issues = QichachaApiAdapter(client, "app", "secret", extra_api_codes=()).fetch("示例有限公司")
    assert not issues
    assert "evidence" not in payload or not any(item["api_code"] in {"2001", "213"} for item in payload["evidence"])
    assert len(client.calls) == 4


def test_qichacha_default_calls_all_non_review_apis():
    client = _Client()
    payload, issues = QichachaApiAdapter(client, "app", "secret").fetch("示例有限公司")
    assert not issues
    called_paths = {url.split("api.qichacha.com/", 1)[-1] for url, _ in client.calls}
    assert called_paths == {
        "ECIInfoVerify/GetInfo", "tm/SearchByApplicant", "PatentV4/Search", "CopyRight/SearchCopyRight",
        "EnterpriseInfo/Verify", "AR/GetAnnualReport",
    }
    assert len(client.calls) == 6


def test_qichacha_does_not_register_review_only_apis():
    assert "962" not in QichachaApiAdapter.DEFAULT_ENDPOINTS
    assert "886" not in QichachaApiAdapter.DEFAULT_ENDPOINTS


def test_narrative_prompt_forbids_model_world_knowledge():
    prompt = Path("demo/prompts/yellow_narratives.v2.txt").read_text(encoding="utf-8")
    assert "不得使用模型自身知识" in prompt
