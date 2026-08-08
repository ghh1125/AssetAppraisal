from __future__ import annotations

"""Company-data adapters.

The Qichacha adapter deliberately keeps the provider-specific signing and
response shaping outside ``domain/``.  It returns only the five fields that
the project yellow-route configuration permits, so provider data cannot leak
into unrelated report fields.
"""

import hashlib
import json
import re
import time
from collections.abc import Mapping
from typing import Any


def _first(record: Mapping[str, Any], *names: str) -> Any:
    for name in names:
        value = record.get(name)
        if value not in (None, "", [], {}):
            return value
    return ""


def _records(payload: Any) -> list[dict[str, Any]]:
    """Normalize the common QCC ``Data``/``Result`` response shapes."""
    if isinstance(payload, list):
        return [item for item in payload if isinstance(item, dict)]
    if not isinstance(payload, dict):
        return []
    for key in ("Data", "Result", "data", "result"):
        value = payload.get(key)
        if isinstance(value, list):
            return [item for item in value if isinstance(item, dict)]
        if isinstance(value, dict):
            nested = _records(value)
            if nested:
                return nested
    return [dict(payload)] if payload else []


def _objects(payload: Any) -> list[dict[str, Any]]:
    """Extract nested object lists from a 735 response."""
    if not isinstance(payload, dict):
        return []
    for key in ("Data", "Result", "data", "result"):
        value = payload.get(key)
        if isinstance(value, dict):
            return value
    return payload


def _join(items: list[str]) -> str:
    return "；".join(item.strip("； ") for item in items if item and item.strip("； "))


def _format_profile(payload: Any) -> str:
    records = _records(payload)
    if not records:
        return ""
    row = records[0]
    labels = (
        ("企业名称", ("Name", "CompanyName", "name")),
        ("统一社会信用代码", ("CreditCode", "CreditCodeNo", "creditCode")),
        ("企业类型", ("EconKind", "EconomicType", "CompanyType")),
        ("注册资本", ("RegistCapi", "RegisteredCapital")),
        ("成立日期", ("StartDate", "EstablishDate")),
        ("法定代表人", ("OperName", "LegalRepresentative")),
        ("注册地址", ("Address", "RegistAddress")),
        ("登记状态", ("Status", "StatusDesc")),
        ("经营范围", ("Scope", "BusinessScope")),
    )
    return _join([f"{label}：{value}" for label, names in labels if (value := _first(row, *names))])


def _profile_record(payload: Any) -> dict[str, str]:
    records = _records(payload)
    if not records:
        return {}
    row = records[0]
    return {
        "credit_code": str(_first(row, "CreditCode", "CreditCodeNo", "creditCode")),
        "name": str(_first(row, "Name", "CompanyName", "name")),
        "company_type": str(_first(row, "EconKind", "EconomicType", "CompanyType")),
        "legal_representative": str(_first(row, "OperName", "LegalRepresentative")),
        "registered_capital": str(_first(row, "RegistCapi", "RegisteredCapital")),
        "establish_date": str(_first(row, "StartDate", "EstablishDate")),
        "term_start": str(_first(row, "TermStart", "BusinessTermStart", "From")),
        "term_end": str(_first(row, "TermEnd", "TeamEnd", "BusinessTermEnd", "EndDate", "To")),
        "registration_authority": str(_first(row, "BelongOrg", "RegistrationAuthority", "RegisterOrg")),
        "approval_date": str(_first(row, "CheckDate", "ApprovalDate")),
        "status": str(_first(row, "Status", "StatusDesc")),
        "address": str(_first(row, "Address", "RegistAddress")),
        "business_scope": str(_first(row, "Scope", "BusinessScope")),
    }


def _format_partners(payload: Any) -> str:
    root = _objects(payload)
    rows = []
    if isinstance(root, dict):
        rows = _records(root.get("Partners", root.get("partners", [])))
    if not rows:
        return ""
    result = []
    for row in rows:
        name = _first(row, "StockName", "PartnerName", "Name", "ShareholderName")
        percent = _first(row, "StockPercent", "Percent", "SharePercent", "PercentOfStock")
        capital = _first(row, "ShouldCapi", "SubscribedCapital", "SubConAm", "Capital")
        date = _first(row, "ShouldDate", "ShoudDate", "SubscribedDate")
        details = _join([
            f"持股比例{percent}" if percent else "",
            f"认缴出资{capital}" if capital else "",
            f"认缴日期{date}" if date else "",
        ])
        if name or details:
            result.append(f"{name}{('：' + details) if details else ''}")
    return "；".join(result)


def _format_changes(payload: Any) -> str:
    root = _objects(payload)
    rows = []
    if isinstance(root, dict):
        rows = _records(root.get("ChangeRecords", root.get("changeRecords", [])))
    if not rows:
        return ""
    result = []
    for row in rows:
        date = _first(row, "ChangeDate", "Date", "AuditDate")
        item = _first(row, "ProjectName", "ChangeItem", "Item")
        before = _first(row, "BeforeContent", "Before", "ContentBefore")
        after = _first(row, "AfterContent", "After", "ContentAfter")
        transition = f"{before}→{after}" if before or after else ""
        text = _join([str(date) if date else "", str(item) if item else "", transition])
        if text:
            result.append(text)
    return "；".join(result)


def _format_patents(payload: Any) -> str:
    rows = _records(payload)
    result = []
    for row in rows:
        title = _first(row, "Title", "PatentName", "InventionName", "Name")
        patent_type = _first(row, "PatentType", "Type", "Kind")
        status = _first(row, "LegalStatus", "Status", "LegalStatusDesc")
        application = _first(row, "ApplicationNumber", "ApplyNo", "ApplicationNo")
        publication = _first(row, "PublicationNumber", "PublicNo", "PublicationNo")
        text = _join([
            f"{title}" if title else "",
            f"类型：{patent_type}" if patent_type else "",
            f"法律状态：{status}" if status else "",
            f"申请号：{application}" if application else "",
            f"公开号：{publication}" if publication else "",
        ])
        if text:
            result.append(text)
    return "；".join(result)


def _patent_rows(payload: Any) -> list[dict[str, str]]:
    rows = []
    for row in _records(payload):
        rows.append(
            {
                "title": str(_first(row, "Title", "PatentName", "InventionName", "Name")),
                "number": str(_first(row, "ApplicationNumber", "ApplyNo", "ApplicationNo", "PublicationNumber", "PublicNo")),
                "application_date": str(_first(row, "ApplicationDate", "ApplyDate")),
                "grant_date": str(_first(row, "GrantDate", "PublicationDate", "PublicDate")),
                "status": str(_first(row, "LegalStatus", "Status", "LegalStatusDesc")),
            }
        )
    return [row for row in rows if any(row.values())]


def _format_trademarks(payload: Any) -> str:
    rows = _records(payload)
    result = []
    for row in rows:
        name = _first(row, "Name", "TmName", "TrademarkName")
        reg_no = _first(row, "RegNo", "RegisterNo", "ApplicationNo")
        category = _first(row, "Category", "IntCls", "InternationalClass")
        status = _first(row, "FlowStatusDesc", "Status", "StatusDesc")
        text = _join([
            name,
            f"注册号：{reg_no}" if reg_no else "",
            f"类别：{category}" if category else "",
            f"状态：{status}" if status else "",
        ])
        if text:
            result.append(text)
    return "；".join(result)


def _trademark_rows(payload: Any) -> list[dict[str, str]]:
    rows = []
    for row in _records(payload):
        image = str(_first(row, "ImageUrl", "Image", "ImgUrl"))
        name = str(_first(row, "Name", "TmName", "TrademarkName"))
        rows.append(
            {
                "application_date": str(_first(row, "ApplyDate", "ApplicationDate")),
                "image": image,
                "name": name or ("图形" if image else ""),
                "registration_number": str(_first(row, "RegNo", "RegisterNo", "ApplicationNo")),
                "class": str(_first(row, "Category", "IntCls", "InternationalClass")),
                "status": str(_first(row, "FlowStatusDesc", "Status", "StatusDesc")),
                "announcement_date": str(_first(row, "AnnouncementDate", "PubDate")),
            }
        )
    return [row for row in rows if any(row.values())]


def _format_software(payload: Any) -> str:
    rows = _records(payload)
    result = []
    for row in rows:
        name = _first(row, "Name", "SoftwareFullName", "FullName")
        short = _first(row, "ShortName", "SoftwareShortName")
        reg_no = _first(row, "RegisterNo", "RegistrationNo", "RegNo")
        date = _first(row, "RegisterAperDate", "RegisterDate", "RegistrationDate")
        version = _first(row, "VersionNo", "Version")
        text = _join([
            name,
            f"简称：{short}" if short else "",
            f"版本：{version}" if version else "",
            f"登记号：{reg_no}" if reg_no else "",
            f"登记日期：{date}" if date else "",
        ])
        if text:
            result.append(text)
    return "；".join(result)


def _software_rows(payload: Any) -> list[dict[str, str]]:
    rows = []
    for index, row in enumerate(_records(payload), 1):
        rows.append(
            {
                "index": str(index),
                "name": str(_first(row, "Name", "SoftwareFullName", "FullName")),
                "registration_number": str(_first(row, "RegisterNo", "RegistrationNo", "RegNo")),
                "first_publication_date": str(_first(row, "FirstPublishDate", "PublishDate", "FinishDate")),
                "approval_date": str(_first(row, "RegisterAperDate", "RegisterDate", "RegistrationDate")),
            }
        )
    return [row for row in rows if any(value for key, value in row.items() if key != "index")]


def _partner_rows(payload: Any) -> list[dict[str, str]]:
    """Return the normalized current shareholder rows from ApiCode 735."""
    root = _objects(payload)
    if not isinstance(root, dict):
        return []
    records = _records(root.get("Partners", root.get("partners", [])))
    result: list[dict[str, str]] = []
    for index, row in enumerate(records, 1):
        result.append(
            {
                "index": str(index),
                "name": str(_first(row, "StockName", "PartnerName", "Name", "ShareholderName")),
                "capital": str(_first(row, "ShouldCapi", "SubscribedCapital", "SubConAm", "Capital")),
                "percent": str(_first(row, "StockPercent", "Percent", "SharePercent", "PercentOfStock")),
                "date": str(_first(row, "ShoudDate", "ShouldDate", "SubscribedDate")),
            }
        )
    return [row for row in result if row["name"] or row["capital"] or row["percent"]]


def _profile_key_no(payload: Any) -> str:
    records = _records(payload)
    if not records:
        return ""
    return str(_first(records[0], "KeyNo", "KeyNo", "keyNo", "CompanyKeyNo"))


def _json_text(payload: Any, limit: int = 5000) -> str:
    """Keep provider evidence bounded while retaining its original facts."""
    if payload in (None, "", [], {}):
        return ""
    try:
        text = json.dumps(payload, ensure_ascii=False, default=str, separators=(",", ":"))
    except (TypeError, ValueError):
        text = str(payload)
    return text[:limit]


def _annual_report_text(payload: Any) -> str:
    rows = _records(payload)
    if not rows:
        return ""
    return "；".join(_json_text(row, 1800) for row in rows if _json_text(row, 1800))


def _fuzzy_candidates_text(payload: Any) -> str:
    result = []
    for row in _records(payload):
        name = _first(row, "Name", "CompanyName", "StockName")
        industry = _first(row, "Industry", "QccIndustry", "Category")
        if name:
            result.append(_join([str(name), f"行业：{industry}" if industry else ""]))
    return "；".join(result)


def _listed_announcements_text(payload: Any) -> str:
    result = []
    for row in _records(payload):
        name = _first(row, "CompanyName", "Name", "StockName")
        code = _first(row, "StockCode", "StockNumber", "Code")
        title = _first(row, "Title", "AnnouncementTitle")
        date = _first(row, "PublishDate", "Date", "PublishTime")
        if name:
            result.append(_join([
                str(name),
                f"股票代码：{code}" if code else "",
                f"公告：{title}" if title else "",
                f"日期：{date}" if date else "",
            ]))
    return "；".join(result)


def _listed_candidate_rows(payload: Any) -> list[dict[str, str]]:
    result = []
    for row in _records(payload):
        key_no = str(_first(row, "KeyNo", "keyNo"))
        name = str(_first(row, "CompanyName", "Name", "StockName"))
        if key_no and name:
            result.append({"key_no": key_no, "name": name})
    return result


def comparable_search_terms(business_scope: str, *, limit: int = 3) -> list[str]:
    """Derive short, searchable industry phrases from a business scope.

    QCC's listed-announcement endpoint rejects full legal business-scope
    sentences.  This uses only the uploaded/API-returned source text and
    generic business-action suffixes; it never contains company-specific
    coordinates or a pre-written peer list.
    """

    scope = re.sub(r"[（(][^）)]*[）)]", "", str(business_scope or ""))
    scope = re.sub(r"各类|相关的?|以及|并提供|提供", "", scope)
    scope = scope.replace("的", "")
    candidates: list[str] = []
    actions = "加工|制造|销售|服务|开发|咨询|研发|安装|维修|租赁|贸易|运营|设计|检测|生产"
    for clause in re.split(r"[，,；;。\n及与和、]", scope):
        clause = clause.strip()
        if not clause:
            continue
        for match in re.finditer(rf"([\u4e00-\u9fff]{{2,16}}?)(?:{actions})", clause):
            term = match.group(1).strip("等")
            if 2 <= len(term) <= 12 and term not in candidates:
                candidates.append(term)
            # A narrower parent term (for example, “汽车零部件” from
            # “汽车零部件热处理”) makes an acceptable fallback query.
            for suffix in ("热处理", "制造", "加工", "服务", "设备"):
                if term.endswith(suffix) and len(term) > len(suffix) + 1:
                    parent = term[: -len(suffix)]
                    if parent not in candidates:
                        candidates.append(parent)
                    if suffix not in candidates:
                        candidates.append(suffix)
    return candidates[:limit]


class QichachaApiAdapter:
    """Signed adapter for QCC company and IP APIs.

    Endpoint paths are configurable because QCC occasionally versions IP
    endpoints.  The 735 and 231 paths are the official current paths; 514 and
    233 can be overridden with ``QICHACHA_ENDPOINT_514`` and
    ``QICHACHA_ENDPOINT_233`` without changing business code.
    """

    DEFAULT_ENDPOINTS = {
        "735": "/ECIInfoVerify/GetInfo",
        "231": "/tm/SearchByApplicant",
        # ApiCode 514 exposes a company multi-patent search; ApiCode 233's
        # software-copyright submethod is the one needed by this report.
        "514": "/PatentV4/Search",
        "233": "/CopyRight/SearchCopyRight",
        # APIs available through normal per-call purchase.  Interfaces that
        # require enterprise real-name / scenario approval are deliberately
        # not registered here.
        "2001": "/EnterpriseInfo/Verify",
        "213": "/AR/GetAnnualReport",
        # Directly purchasable candidate-discovery APIs. They are called only
        # for the target company's business keywords, never for the client.
        "886": "/FuzzySearch/GetList",
        "915": "/IPOAnnouncement/GetList",
        # ApiCode 699 contains several subresources.  Both are only called
        # after 915 has returned an actual listed-company KeyNo.
        "699_detail": "/IPO/GetIPODetail",
        "699_indicator": "/IPO/GetMainIndicator",
    }

    def __init__(
        self,
        client: Any = None,
        app_key: str | None = None,
        secret_key: str | None = None,
        base_url: str = "https://api.qichacha.com",
        endpoints: Mapping[str, str] | None = None,
        extra_api_codes: tuple[str, ...] | list[str] | None = None,
        enable_comparable_discovery: bool = True,
        timeout: float = 120.0,
    ):
        self.client = client
        self.app_key = app_key or ""
        self.secret_key = secret_key or ""
        self.base_url = base_url.rstrip("/")
        self.endpoints = {**self.DEFAULT_ENDPOINTS, **dict(endpoints or {})}
        supported = set(self.DEFAULT_ENDPOINTS)
        # The two normal-purchase evidence APIs run by default in node 2.
        # Passing an explicit empty collection is useful to tests and callers
        # that want the four direct Word-fill APIs only.
        requested_extra = ("2001", "213") if extra_api_codes is None else extra_api_codes
        self.extra_api_codes = tuple(
            code for code in dict.fromkeys(str(item) for item in requested_extra)
            if code in {"2001", "213"}
        )
        self.enable_comparable_discovery = enable_comparable_discovery
        self.timeout = timeout

    @staticmethod
    def token(app_key: str, timespan: str, secret_key: str) -> str:
        return hashlib.md5(f"{app_key}{timespan}{secret_key}".encode("utf-8")).hexdigest().upper()

    def _get(self, code: str, company_name: str, key_no: str = "") -> tuple[Any, str | None]:
        if not self.client:
            return None, "企查查 API 客户端未配置"
        path = self.endpoints[code]
        url = path if path.startswith("http") else f"{self.base_url}/{path.lstrip('/')}"
        timespan = str(int(time.time()))
        params: dict[str, Any] = {"key": self.app_key}
        # ApiCode 735 accepts only searchKey; the IP list endpoints accept
        # pagination as documented.  Sending pageIndex/pageSize to 735 is
        # rejected by the current API with status 125.
        if code == "231":
            params.update({"keyword": company_name, "pageIndex": 1, "pageSize": 50})
        elif code == "213":
            # Annual-report API officially uses keyNo. Some QCC tenants also
            # accept searchKey, so retain a safe fallback for name-only calls.
            params["keyNo"] = key_no or company_name
        elif code in {"699_detail", "699_indicator"}:
            params["keyNo"] = key_no
        else:
            params["searchKey"] = company_name
            if code in {"514", "233", "886", "915"}:
                params.update({"pageIndex": 1, "pageSize": 50})
        headers = {"Token": self.token(self.app_key, timespan, self.secret_key), "Timespan": timespan}
        try:
            response = self.client.get(url, params=params, headers=headers, timeout=self.timeout)
            response.raise_for_status()
            payload = response.json()
            status = str(payload.get("Status", payload.get("status", "200"))) if isinstance(payload, dict) else "200"
            if status not in {"200", "0", "成功"}:
                message = payload.get("Message", payload.get("message", "接口返回失败")) if isinstance(payload, dict) else "接口返回失败"
                # Never parse the envelope itself as a company record.  QCC
                # error envelopes often contain ``Status`` (for example 125),
                # which would otherwise be written into the report as a fake
                # registration status.
                return None, f"企查查接口 {code} 返回 {status}：{message}"
            return payload, None
        except Exception as exc:
            return None, f"企查查接口 {code} 请求失败：{type(exc).__name__}"

    def fetch(self, company_name: str) -> tuple[dict[str, Any], list[str]]:
        if not company_name:
            return {}, ["企查查 API 未获得企业名称，相关字段留空"]
        if not self.app_key or not self.secret_key:
            return {}, ["企查查 API 未配置 AppKey/SecretKey，相关字段留空"]
        payloads: dict[str, Any] = {}
        issues: list[str] = []
        base_codes = ("735", "231", "514", "233")
        profile_payload, profile_issue = self._get("735", company_name)
        if profile_issue:
            issues.append(profile_issue)
        if profile_payload is not None:
            payloads["735"] = profile_payload
        key_no = _profile_key_no(profile_payload)
        for code in (*base_codes[1:], *self.extra_api_codes):
            payload, issue = self._get(code, company_name, key_no=key_no)
            if issue:
                issues.append(issue)
            if payload is not None:
                payloads[code] = payload
        profile_payload = payloads.get("735")
        patents = _format_patents(payloads.get("514"))
        trademarks = _format_trademarks(payloads.get("231"))
        fields = {
            "commissioning_party_profile": _format_profile(profile_payload),
            "ownership_history": _format_changes(profile_payload),
            "ownership_at_valuation_date": _format_partners(profile_payload),
            "unrecorded_intangibles": _join([patents, trademarks]),
            "software_copyrights": _format_software(payloads.get("233")),
        }
        evidence: list[dict[str, str]] = []
        evidence_by_topic: dict[str, str] = {}

        def add_evidence(api_code: str, topic: str, text: str) -> None:
            if not text:
                return
            item = {
                "evidence_id": f"api:qichacha:{api_code}:{topic}",
                "api_code": api_code,
                "topic": topic,
                "source_kind": "qichacha_api",
                "text": text,
            }
            evidence.append(item)
            evidence_by_topic[topic] = _join([evidence_by_topic.get(topic, ""), text])

        add_evidence("735", "company_profile_section", _format_profile(profile_payload))
        add_evidence("735", "ownership_history", _format_changes(profile_payload))
        add_evidence("735", "ownership_at_valuation_date", _format_partners(profile_payload))
        add_evidence("231", "unrecorded_intangibles", trademarks)
        add_evidence("514", "unrecorded_intangibles", patents)
        add_evidence("233", "software_copyrights", _format_software(payloads.get("233")))
        if "2001" in payloads:
            add_evidence("2001", "industry_overview", _json_text(payloads["2001"]))
            add_evidence("2001", "business_and_segments", _json_text(payloads["2001"]))
        if "213" in payloads:
            add_evidence("213", "profit_model_swot", _annual_report_text(payloads["213"]) or _json_text(payloads["213"]))
            add_evidence("213", "business_and_segments", _annual_report_text(payloads["213"]) or _json_text(payloads["213"]))
        return {
            "fields": {key: value for key, value in fields.items() if value},
            "profile": _profile_record(profile_payload),
            "partner_rows": _partner_rows(profile_payload),
            "trademark_rows": _trademark_rows(payloads.get("231")),
            "patent_rows": _patent_rows(payloads.get("514")),
            "software_rows": _software_rows(payloads.get("233")),
            "software_query_ok": "233" in payloads,
            "evidence": evidence,
            "evidence_by_topic": evidence_by_topic,
        }, issues

    def discover_listed_comparables(
        self,
        business_keywords: list[str] | tuple[str, ...],
    ) -> tuple[list[dict[str, str]], list[str]]:
        """Find evidence-backed peer candidates from supplied business terms.

        This deliberately returns candidate evidence, not a final peer list.
        The pipeline and LLM may only describe the returned companies; they
        must not introduce a company name of their own.
        """
        if not self.enable_comparable_discovery:
            return [], []
        keywords = list(dict.fromkeys(
            str(keyword).strip() for keyword in business_keywords if str(keyword).strip()
        ))[:3]
        evidence: list[dict[str, str]] = []
        issues: list[str] = []
        # ApiCode 699 is billed per detail/indicator request. Keep one
        # run within the purchased trial quota even when several keywords
        # return overlapping companies (at most 5 peers × 2 calls = 10).
        inspected_key_nos: set[str] = set()
        max_detailed_peers = 5
        for keyword in keywords:
            fuzzy_payload, fuzzy_issue = self._get("886", keyword)
            if fuzzy_issue:
                issues.append(f"对标候选企业检索“{keyword}”：{fuzzy_issue}")
            elif text := _fuzzy_candidates_text(fuzzy_payload):
                evidence.append({
                    "evidence_id": f"api:qichacha:886:peer:{keyword}",
                    "api_code": "886",
                    "topic": "comparable_list",
                    "source_kind": "qichacha_api",
                    "text": f"关键词“{keyword}”的同行业企业候选：{text}",
                })
            ipo_payload, ipo_issue = self._get("915", keyword)
            if ipo_issue:
                issues.append(f"对标上市公司检索“{keyword}”：{ipo_issue}")
            elif text := _listed_announcements_text(ipo_payload):
                evidence.append({
                    "evidence_id": f"api:qichacha:915:peer:{keyword}",
                    "api_code": "915",
                    "topic": "comparable_list",
                    "source_kind": "qichacha_api",
                    "text": f"关键词“{keyword}”命中的上市公司公告候选：{text}",
                })
            for candidate in _listed_candidate_rows(ipo_payload)[:5]:
                key_no = str(candidate.get("key_no") or "")
                if not key_no or key_no in inspected_key_nos:
                    continue
                if len(inspected_key_nos) >= max_detailed_peers:
                    break
                inspected_key_nos.add(key_no)
                detail_payload, detail_issue = self._get(
                    "699_detail", candidate["name"], key_no=key_no
                )
                indicator_payload, indicator_issue = self._get(
                    "699_indicator", candidate["name"], key_no=key_no
                )
                if detail_issue:
                    issues.append(f"上市候选“{candidate['name']}”简介查询：{detail_issue}")
                if indicator_issue:
                    issues.append(f"上市候选“{candidate['name']}”指标查询：{indicator_issue}")
                detail_text = _json_text(detail_payload, 2500)
                indicator_text = _json_text(indicator_payload, 2500)
                if detail_text or indicator_text:
                    evidence.append({
                        "evidence_id": f"api:qichacha:699:peer:{key_no}",
                        "api_code": "699",
                        "topic": "comparable_list",
                        "source_kind": "qichacha_api",
                        "text": _join([
                            f"上市候选公司：{candidate['name']}",
                            f"企业简介：{detail_text}" if detail_text else "",
                            f"主要指标：{indicator_text}" if indicator_text else "",
                        ]),
                    })
        return evidence, issues


class CompanyApiAdapter:
    """Backward-compatible generic adapter used by older Demo callers."""

    def __init__(self, client: Any = None, endpoint: str | None = None, api_key: str | None = None):
        self.client, self.endpoint, self.api_key = client, endpoint, api_key

    def fetch(self, company_name: str) -> tuple[dict[str, Any], list[str]]:
        if not self.client or not self.endpoint or not self.api_key:
            return {}, ["企业数据 API 未配置，相关字段按规则留空"]
        try:
            response = self.client.get(self.endpoint, params={"name": company_name}, headers={"Authorization": self.api_key})
            response.raise_for_status()
            return response.json(), []
        except Exception as exc:
            return {}, [f"企业数据 API 失败：{type(exc).__name__}"]
