from pathlib import Path

from openpyxl import Workbook

from demo.adapters.company_api import (
    CompanyApiAdapter,
    QichachaApiAdapter,
    comparable_search_terms,
)
from demo.adapters.bailian_glm import BailianYellowNarrativeAdapter
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


def test_qichacha_699_default_endpoint_uses_stock_code_query():
    class ComparableClient(_Client):
        def get(self, url, **kwargs):
            self.calls.append((url, kwargs))
            if url.endswith("FuzzySearch/GetList"):
                return _Response({"Status": "200", "Result": []})
            if url.endswith("IPOAnnouncement/GetList"):
                return _Response({"Status": "200", "Result": {"Data": [{
                    "CompanyName": "上市公司甲",
                    "StockCode": "600001",
                    "Title": "热处理业务公告",
                    "PublishDate": "2026-01-01",
                }]}})
            if url.endswith("IPO/GetIPODetail"):
                return _Response({"Status": "200", "Result": {}})
            return super().get(url, **kwargs)

    adapter = QichachaApiAdapter(ComparableClient(), "app", "secret")
    adapter.discover_listed_comparables(["汽车零部件热处理"])

    assert QichachaApiAdapter.DEFAULT_ENDPOINTS["699"] == "/IPO/GetIPODetail"
    detail_call = next(
        kwargs for url, kwargs in adapter.client.calls if url.endswith("IPO/GetIPODetail")
    )
    assert detail_call["params"]["stockCode"] == "600001"
    assert "searchKey" not in detail_call["params"]


def test_qichacha_discovers_listed_comparable_candidates_when_api_699_has_no_detail():
    class ComparableClient(_Client):
        def get(self, url, **kwargs):
            self.calls.append((url, kwargs))
            if url.endswith("FuzzySearch/GetList"):
                return _Response({"Status": "200", "Result": [{"Name": "同行企业甲", "Industry": "汽车零部件"}]})
            if url.endswith("IPOAnnouncement/GetList"):
                return _Response({"Status": "200", "Result": {"Data": [{"CompanyName": "上市公司甲", "KeyNo": "peer-key", "StockCode": "600001", "Title": "热处理业务公告", "PublishDate": "2026-01-01"}]}})
            if url.endswith("IPO/GetIPODetail"):
                return _Response({"Status": "200", "Result": {}})
            return super().get(url, **kwargs)

    adapter = QichachaApiAdapter(ComparableClient(), "app", "secret")
    evidence, issues = adapter.discover_listed_comparables(["汽车零部件热处理"])

    assert issues == []
    assert {item["api_code"] for item in evidence} == {"886", "915"}
    assert "上市公司甲" in next(item["text"] for item in evidence if item["api_code"] == "915")
    assert "600001" in next(item["text"] for item in evidence if item["api_code"] == "915")
    assert any("IPO/GetIPODetail" in url for url, _ in adapter.client.calls)


def test_qichacha_enriches_verified_listed_candidates_with_flat_api_699_response():
    class ComparableClient(_Client):
        def get(self, url, **kwargs):
            self.calls.append((url, kwargs))
            if url.endswith("FuzzySearch/GetList"):
                return _Response({"Status": "200", "Result": [{"Name": "同行企业甲", "Industry": "汽车零部件"}]})
            if url.endswith("IPOAnnouncement/GetList"):
                return _Response({"Status": "200", "Result": {"Data": [{
                    "CompanyName": "上市公司甲", "StockCode": "600001", "Title": "热处理业务公告", "PublishDate": "2026-01-01"
                }]}})
            if url.endswith("IPO/GetIPODetail"):
                return _Response({"Status": "200", "Result": {
                    "CompanyName": "上市公司甲", "StockCode": "600001", "Industry": "汽车零部件制造",
                    "MainBusiness": "汽车零部件热处理"
                }})
            return super().get(url, **kwargs)

    adapter = QichachaApiAdapter(ComparableClient(), "app", "secret")
    evidence, issues = adapter.discover_listed_comparables(["汽车零部件热处理"])

    assert issues == []
    listed_detail = next(item for item in evidence if item["api_code"] == "699")
    assert "上市公司甲" in listed_detail["text"]
    assert "600001" in listed_detail["text"]
    assert "汽车零部件制造" in listed_detail["text"]
    assert "IPO/GetIPODetail" in {url.split("api.qichacha.com/", 1)[-1] for url, _ in adapter.client.calls}


def test_qichacha_699_nested_response_enriches_verified_candidate():
    class ComparableClient(_Client):
        def get(self, url, **kwargs):
            self.calls.append((url, kwargs))
            if url.endswith("FuzzySearch/GetList"):
                return _Response({"Status": "200", "Result": []})
            if url.endswith("IPOAnnouncement/GetList"):
                return _Response({"Status": "200", "Result": {"Data": [{
                    "CompanyName": "上市公司甲",
                    "StockCode": "600001",
                    "Title": "热处理业务公告",
                    "PublishDate": "2026-01-01",
                }]}})
            if url.endswith("IPO/GetIPODetail"):
                return _Response({"Status": "200", "Result": {
                    "BaseInfo": {
                        "Companyname": "上市公司甲",
                        "ACode": "600001",
                        "AShortName": "公司甲",
                        "Industry": "汽车制造业",
                        "StockCategory": "上交所主板",
                        "MarketDate": "2011-06-30",
                        "RegAddress": "上海市示例路1号",
                        "PBR": "4.05",
                        "PER": "28.64",
                    },
                    "IPOPublishInfo": {
                        "EstablishDate": "1995-02-10",
                        "ReleasesType": "网下询价配售",
                    },
                }})
            return super().get(url, **kwargs)

    evidence, issues = QichachaApiAdapter(
        ComparableClient(), "app", "secret"
    ).discover_listed_comparables(["汽车零部件热处理"])

    assert issues == []
    detail = next(item["text"] for item in evidence if item["api_code"] == "699")
    for expected in (
        "企业名称：上市公司甲",
        "股票代码：600001",
        "股票简称：公司甲",
        "所属行业：汽车制造业",
        "证券类别：上交所主板",
        "上市日期：2011-06-30",
        "注册地址：上海市示例路1号",
        "市净率：4.05",
        "市盈率：28.64",
        "成立日期：1995-02-10",
        "发行方式：网下询价配售",
    ):
        assert expected in detail


def test_qichacha_699_does_not_treat_business_scope_as_main_business():
    from demo.adapters.company_api import _listed_company_detail_text

    detail = _listed_company_detail_text(
        {"Status": "200", "Result": {"BaseInfo": {
            "Companyname": "上市公司甲",
            "ACode": "600001",
            "Industry": "汽车制造业",
            "BusinessScope": "汽车零部件制造及销售",
        }}},
        expected_name="上市公司甲",
        expected_code="600001",
    )

    assert "所属行业：汽车制造业" in detail
    assert "主营业务" not in detail


def test_listed_announcements_are_deduplicated_by_company_and_stock_code():
    from demo.adapters.company_api import _listed_announcements_text

    text = _listed_announcements_text(
        {"Status": "200", "Result": {"Data": [
            {"CompanyName": "上市公司甲", "StockCode": "600001", "Title": "公告一", "PublishDate": "2026-01-01"},
            {"CompanyName": "上市公司甲", "StockCode": "600001", "Title": "公告二", "PublishDate": "2026-01-02"},
            {"CompanyName": "上市公司乙", "StockCode": "600002", "Title": "公告三", "PublishDate": "2026-01-03"},
        ]}}
    )

    assert text.count("上市公司甲") == 1
    assert "上市公司乙" in text


def test_qichacha_can_disable_paid_comparable_discovery():
    client = _Client()
    adapter = QichachaApiAdapter(
        client,
        "app",
        "secret",
        enable_comparable_discovery=False,
    )
    evidence, issues = adapter.discover_listed_comparables(["汽车零部件热处理"])
    assert evidence == []
    assert issues == []
    assert client.calls == []


def test_comparable_search_terms_reduce_business_scope_to_compact_industry_phrases():
    assert comparable_search_terms(
        "各类汽车零部件的热处理加工并提供相关的技术开发、技术咨询及售后服务。"
    ) == ["汽车零部件热处理", "汽车零部件", "热处理"]


def test_comparable_narrative_evidence_requires_an_announcement_with_multiple_dimensions():
    evidence = [
        {
            "evidence_id": "api:qichacha:target:735:profile",
            "text": "被评估单位工商信息：行业：制造业；经营范围：汽车零部件热处理",
        },
        {
            "evidence_id": "api:qichacha:target:915:peer:汽车零部件热处理",
            "text": "上市公司公告候选：上市公司甲；股票代码：600001；公告：热处理业务公告；日期：2026-01-01",
        },
    ]

    selected = BailianYellowNarrativeAdapter._relevant_evidence(
        "comparable_list", evidence
    )

    assert selected == [evidence[1]]


def test_comparable_narrative_evidence_allows_api_699_only_as_verified_915_enrichment():
    announcement = {
        "evidence_id": "api:qichacha:target:915:peer:热处理",
        "text": "上市公司公告候选：上市公司甲；股票代码：600001；公告：热处理业务公告；日期：2026-01-01",
    }
    listed_detail = {
        "evidence_id": "api:qichacha:target:699:peer:上市公司甲:600001",
        "text": "上市公司补充信息（已由公告候选核验）：企业名称：上市公司甲；股票代码：600001；所属行业：汽车零部件制造；主营业务：汽车零部件热处理",
    }

    selected = BailianYellowNarrativeAdapter._relevant_evidence(
        "comparable_list", [announcement, listed_detail]
    )

    assert selected == [announcement, listed_detail]


def test_comparable_narrative_evidence_rejects_a_name_only_candidate():
    evidence = [
        {
            "evidence_id": "api:qichacha:target:915:peer:工业设备",
            "text": "上市公司公告候选：上市公司甲；股票代码：600001",
        },
    ]

    selected = BailianYellowNarrativeAdapter._relevant_evidence(
        "comparable_list", evidence
    )

    assert selected == []


def test_resolved_structured_financial_evidence_outranks_conflicting_raw_ocr():
    raw_ocr = {
        "evidence_id": "pdf:p12:b3",
        "text": "利润表营业收入65,506,460.03元，净利润20,869,022.09元",
    }
    resolved_table = {
        "evidence_id": "field:historical_income_statement_table",
        "text": "历史利润表：营业收入46,186,357.24元，净利润14,357,065.14元",
    }

    selected = BailianYellowNarrativeAdapter._relevant_evidence(
        "profit_model_swot",
        [raw_ocr, resolved_table],
        limit=1,
    )

    assert selected == [resolved_table]


def test_company_profile_always_keeps_target_qcc_profile_with_many_fields():
    target_profile = {
        "evidence_id": "api:qichacha:target:735:profile",
        "text": "被评估单位工商信息：统一社会信用代码91320000123456789X；经营范围：热处理加工",
    }
    field_evidence = [
        {
            "evidence_id": f"field:business_fact_{index}",
            "text": f"公司业务事实{index}",
        }
        for index in range(20)
    ]

    selected = BailianYellowNarrativeAdapter._relevant_evidence(
        "company_profile_section",
        [*field_evidence, target_profile],
    )

    assert target_profile in selected


def test_company_profile_excludes_current_qcc_shareholders_from_historical_narrative():
    profile = {
        "evidence_id": "api:qichacha:target:735:profile",
        "text": "被评估单位工商信息：统一社会信用代码91320000123456789X；经营范围：热处理加工",
    }
    current_partner = {
        "evidence_id": "api:qichacha:target:735:partner:1",
        "text": "当前股东：示例投资有限公司，持股100%",
    }

    selected = BailianYellowNarrativeAdapter._relevant_evidence(
        "company_profile_section",
        [profile, current_partner],
    )

    assert profile in selected
    assert current_partner not in selected


def test_comparable_fallback_formats_verified_915_and_699_records():
    evidence = [
        {
            "evidence_id": "api:qichacha:target:915:peer:汽车零部件",
            "text": (
                "关键词“汽车零部件”命中的上市公司公告候选："
                "宁波继峰汽车零部件股份有限公司；股票代码：603997；"
                "公告类别：股权质押；公告：股份质押公告；日期：2026-08-08"
            ),
        },
        {
            "evidence_id": "api:qichacha:target:699:peer:603997",
            "text": (
                "上市公司补充信息（已由公告候选核验）："
                "企业名称：宁波继峰汽车零部件股份有限公司；股票代码：603997；"
                "所属行业：汽车制造业；证券类别：上交所主板；"
                "上市日期：2015-03-02；市净率：3.08；市盈率：30.81"
            ),
        },
    ]

    value = BailianYellowNarrativeAdapter._comparable_fallback(evidence)

    assert "宁波继峰汽车零部件股份有限公司｜股票代码：603997" in value
    assert "所属行业：汽车制造业" in value
    assert "证券类别：上交所主板；上市日期：2015-03-02" in value
    assert "公告依据：股份质押公告（2026-08-08）" in value
    assert "估值指标：市盈率30.81；市净率3.08" in value
    assert "ApiCode" not in value


def test_narrative_target_context_keeps_subject_and_valuation_date_together():
    context = BailianYellowNarrativeAdapter._target_context(
        [
            {
                "evidence_id": "field:target_company_name",
                "text": "被评估单位名称：通富热处理（昆山）有限公司",
            },
            {
                "evidence_id": "field:target_company_short_name",
                "text": "被评估单位简称：通富昆山",
            },
            {
                "evidence_id": "field:valuation_date_year",
                "text": "评估基准日年：2025",
            },
            {
                "evidence_id": "field:valuation_date_month",
                "text": "评估基准日月：6",
            },
            {
                "evidence_id": "field:valuation_date_day",
                "text": "评估基准日日：30",
            },
            {
                "evidence_id": "field:registered_capital",
                "text": "注册资本：1,000万元",
            },
            {
                "evidence_id": "field:main_products",
                "text": "主要产品：通富新材料科技（仙桃）有限公司从事热处理加工",
            },
        ]
    )

    assert context == {
        "target_company_name": "通富热处理（昆山）有限公司",
        "target_company_short_name": "通富昆山",
        "registered_capital": "1,000万元",
        "valuation_date": "2025年6月30日",
    }


def test_narrative_prompt_forbids_model_world_knowledge():
    prompt = Path("demo/prompts/yellow_narratives.v2.txt").read_text(encoding="utf-8")
    assert "不得使用模型自身知识" in prompt
    assert "未披露主要客户及供应商" in prompt
    assert "不等于可比性最终认定" in prompt


def test_narrative_v3_prompt_requires_six_module_evidence_boundaries():
    prompt = Path("demo/prompts/yellow_narratives.v3.txt").read_text(encoding="utf-8")

    assert "699" in prompt
    assert "客户及供应商" in prompt
    assert "不得编造" in prompt
    assert "公告标题" in prompt


def test_narrative_v3_prompt_guides_professional_evidence_grounded_analysis():
    prompt = Path("demo/prompts/yellow_narratives.v3.txt").read_text(encoding="utf-8")

    assert "专业综合" in prompt
    assert "审慎分析" in prompt
    assert "自然、连贯" in prompt
    assert "不得虚构" in prompt
    assert "逐句复制" in prompt
    assert "target_context" in prompt
    assert "评估基准日之后" in prompt
    assert "其他公司" in prompt


def test_qichacha_699_needs_no_endpoint_setting_in_env_example():
    env_example = Path(".env.example").read_text(encoding="utf-8")

    assert "QICHACHA_APP_KEY=" in env_example
    assert "QICHACHA_SECRET_KEY=" in env_example
    assert "QICHACHA_ENDPOINT_699" not in env_example
