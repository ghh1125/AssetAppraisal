"""Batch-generate role workbooks from an extracted audit-material directory.

Usage (the project entrypoint loads credentials from ``.env``):

    python -m demo.run_material_intake D:\\assets\\unpacked runs\\material-intake
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import tempfile
from typing import Any, Callable
from zipfile import ZipFile
from xml.etree import ElementTree

import pdfplumber
from docx import Document
from openpyxl import load_workbook

from demo.adapters.audit_intake import needs_ocr, tables_from_pages
from demo.adapters.template_workbook_mapper import build_template_workbooks
from demo.adapters.ocr_factory import create_ocr_adapter
from demo.domain.company_matching import matching_company_records, normalize_company_name
from demo.run import _load_local_env
from demo.template_bundle import load_workbook_template_bundle


AUDIT_HINTS = ("审计", "财务报表", "年审", "单体", "合并", "2年一期")
MATERIAL_SUFFIXES = {
    ".pdf", ".png", ".jpg", ".jpeg",
    ".doc", ".docx", ".xls", ".xlsx", ".xlsm", ".ppt", ".pptx",
}
# Legacy Office binary files are delegated to the configured layout OCR; OOXML
# files are read natively first so their sheets/tables stay addressable.
OCR_SUFFIXES = {".pdf", ".png", ".jpg", ".jpeg", ".doc", ".xls", ".ppt"}


def _write_generic_rule_graph(workbook_root: Path) -> None:
    """Persist the reusable evidence-to-template mapping contract as a graph."""
    workbook_root.mkdir(parents=True, exist_ok=True)
    mermaid = """flowchart LR
  subgraph S[材料证据层]
    audit[审计报告/财务报表/附注]
    registry[营业执照/企业信息报告/截图]
    ocr[版式 OCR：页码/表格/行列/原文]
    audit --> ocr
    registry --> ocr
  end
  subgraph F[标准事实层]
    classify[表类型识别：资产负债表/利润表/现金流量表/附注明细]
    fact[标准事实：主体+口径+期间+单位+标准科目+数值+证据坐标]
    ocr --> classify --> fact
  end
  subgraph G[安全门]
    account{目标与来源是否同一标准科目}
    period{期间是否明确}
    unit{金额单位是否明确}
    scope{母公司/合并口径是否唯一}
    fact --> account --> period --> unit --> scope
  end
  scope -->|全部通过| convert[金额×来源单位÷模板单位]
  account -.不通过.-> missing[目标单元格留空+补充资料清单]
  period -.不通过.-> missing
  unit -.不通过.-> missing
  scope -.不通过.-> missing
  classify -->|附注/明细尚无唯一记录主键| candidate[附注及明细候选：保留原表，不猜填]
  subgraph T[系统通用工作簿模板]
    asset[资产基础法/资产清查：只写审定账面值]
    income[收益法/市场法：只写审定历史三表]
    formula[原 Sheet、样式、公式坐标和公式文本保留]
    convert --> asset
    convert --> income
    asset --> formula
    income --> formula
  end
  formula --> result[两份业务工作簿]
  fact --> ledger[逐单元格映射规则]
  formula --> formulas[逐单元格公式规则]
  missing --> registrySheet[模板字段注册表/缺失资料]
  candidate --> trace[审计材料_规则与溯源.xlsx]
  ledger --> trace
  formulas --> trace
  registrySheet --> trace
"""
    (workbook_root / "通用映射规则图.mmd").write_text(mermaid, encoding="utf-8")
    graph = {
        "version": "generic_appraisal_mapping.v2",
        "fact_key": ["subject", "statement_type", "scope", "period", "unit", "canonical_account", "value"],
        "evidence_key": ["source_file", "page", "table_id", "row", "column", "raw_text"],
        "gates": [
            {"id": "same_account", "rule": "目标字段与来源字段规范化后必须完全同名；禁止模糊相似度自动填值"},
            {"id": "known_period", "rule": "来源金额列必须有明确日期、年度或本年/上年相对期间"},
            {"id": "known_unit", "rule": "来源页必须明示人民币元或万元；不得默认猜测"},
            {"id": "unique_scope", "rule": "母公司与合并口径同时存在且金额不一致时不得自动选择"},
            {"id": "target_role", "rule": "账面净额不得写入原值、余额、减值准备等目标字段；目标字段角色必须与证据金额语义一致"},
            {"id": "single_effective_mapping", "rule": "每个目标单元格只允许一条生效映射；其他来源仅作为校验、冲突或被替代证据"},
            {"id": "case_independence", "rule": "公司名、文件名、页码和案例金额不得作为通用匹配条件"},
        ],
        "transforms": [
            {"id": "direct_unit_conversion", "expression": "target_value = source_value * source_unit_to_yuan / target_unit_to_yuan"},
            {"id": "template_formula", "expression": "保留版本化通用模板的公式坐标与公式文本，由 Excel 计算"},
            {"id": "unsupported", "expression": "目标留空，并登记需要补充的证据材料"},
        ],
        "targets": {
            "资产基础法_资产清查.xlsx": "审计材料只支持审定账面值；评估值、设备明细和权属资料无证据时留空",
            "收益法_市场法.xlsx": "审计材料只支持历史三表；预测、WACC、资本开支和估值参数无证据时留空",
        },
    }
    (workbook_root / "通用映射规则.json").write_text(json.dumps(graph, ensure_ascii=False, indent=2), encoding="utf-8")


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _safe_name(path: Path) -> str:
    value = re.sub(r"[\\/:*?\"<>|]", "_", path.stem)
    value = re.sub(r"\s+", "_", value).strip("._")
    return value[:80] or "审计材料"


def _subject_key(value: Any) -> str:
    return normalize_company_name(value)


def _local_pages(path: Path) -> list[dict[str, Any]]:
    pages: list[dict[str, Any]] = []
    with pdfplumber.open(path) as pdf:
        for page_number, page in enumerate(pdf.pages, 1):
            tables = []
            for table_number, raw_table in enumerate(page.extract_tables() or [], 1):
                cells = []
                for row_number, row in enumerate(raw_table or [], 1):
                    for column_number, value in enumerate(row or [], 1):
                        cells.append(
                            {
                                "row": row_number,
                                "column": column_number,
                                "row_span": 1,
                                "column_span": 1,
                                "text": str(value or ""),
                                "confidence": None,
                                "bbox": [],
                            }
                        )
                tables.append({"table_id": f"local-p{page_number}-t{table_number}", "cells": cells})
            pages.append(
                {
                    "page_number": page_number,
                    "page_count": len(pdf.pages),
                    "blocks": [
                        {
                            "block_id": f"local-p{page_number}-text",
                            "block_type": "text",
                            "text": page.extract_text() or "",
                            "confidence": None,
                            "bbox": [],
                        }
                    ],
                    "tables": tables,
                }
            )
    return pages


def _table_page(*, page_number: int, title: str, cells: list[dict[str, Any]]) -> dict[str, Any]:
    return {
        "page_number": page_number,
        "page_count": 1,
        "blocks": [{
            "block_id": f"structured-p{page_number}-title",
            "block_type": "text",
            "text": title,
            "confidence": 1,
            "bbox": [],
        }],
        "tables": [{"table_id": f"structured-p{page_number}-t1", "cells": cells}] if cells else [],
    }


def _docx_pages(path: Path) -> list[dict[str, Any]]:
    document = Document(path)
    blocks = [paragraph.text.strip() for paragraph in document.paragraphs if paragraph.text.strip()]
    cells: list[dict[str, Any]] = []
    for table_number, table in enumerate(document.tables, 1):
        for row_number, row in enumerate(table.rows, 1):
            for column_number, cell in enumerate(row.cells, 1):
                text = cell.text.strip()
                if text:
                    cells.append({
                        "row": row_number,
                        "column": column_number,
                        "row_span": 1,
                        "column_span": 1,
                        "text": text,
                        "confidence": 1,
                        "bbox": [],
                        "table_number": table_number,
                    })
    page = _table_page(page_number=1, title="\n".join(blocks), cells=cells)
    page["page_count"] = 1
    return [page] if blocks or cells else []


def _xlsx_pages(path: Path) -> list[dict[str, Any]]:
    workbook = load_workbook(path, read_only=True, data_only=False)
    pages: list[dict[str, Any]] = []
    try:
        for page_number, sheet in enumerate(workbook.worksheets, 1):
            cells = []
            for row_number, row in enumerate(sheet.iter_rows(), 1):
                for column_number, cell in enumerate(row, 1):
                    if cell.value in (None, ""):
                        continue
                    cells.append({
                        "row": row_number,
                        "column": column_number,
                        "row_span": 1,
                        "column_span": 1,
                        "text": str(cell.value),
                        "confidence": 1,
                        "bbox": [],
                    })
            if cells:
                pages.append(_table_page(page_number=page_number, title=f"Excel 工作表：{sheet.title}", cells=cells))
    finally:
        workbook.close()
    page_count = len(pages)
    for page in pages:
        page["page_count"] = page_count
    return pages


def _pptx_pages(path: Path) -> list[dict[str, Any]]:
    """Read text from PPTX without relying on an Office installation."""
    pages: list[dict[str, Any]] = []
    namespace = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    with ZipFile(path) as archive:
        slide_names = sorted(
            (name for name in archive.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)),
            key=lambda name: int(re.search(r"slide(\d+)", name).group(1)),
        )
        for page_number, name in enumerate(slide_names, 1):
            root = ElementTree.fromstring(archive.read(name))
            text = "\n".join(node.text.strip() for node in root.findall(".//a:t", namespace) if node.text and node.text.strip())
            if text:
                pages.append(_table_page(page_number=page_number, title=text, cells=[]))
    page_count = len(pages)
    for page in pages:
        page["page_count"] = page_count
    return pages


def _structured_pages(path: Path) -> list[dict[str, Any]]:
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        return _local_pages(path)
    if suffix == ".docx":
        return _docx_pages(path)
    if suffix in {".xlsx", ".xlsm"}:
        return _xlsx_pages(path)
    if suffix == ".pptx":
        return _pptx_pages(path)
    return []


def _docx_embedded_image_pages(path: Path, ocr_adapter: Any) -> tuple[list[dict[str, Any]], list[str]]:
    """OCR images pasted into Word, such as a scanned business licence.

    Word does not expose a stable PDF page number without an Office layout
    pass, so the resulting evidence coordinates are explicitly labelled by
    embedded-image order.  This is still traceable to the uploaded Word file
    and materially better than silently dropping the image.
    """
    if ocr_adapter is None:
        return [], ["Word 内嵌图片需要 OCR，但当前 OCR 未配置"]
    pages: list[dict[str, Any]] = []
    issues: list[str] = []
    allowed = {".png", ".jpg", ".jpeg"}
    with ZipFile(path) as archive, tempfile.TemporaryDirectory(prefix="appraisal-docx-media-") as temporary:
        media_names = [
            name for name in archive.namelist()
            if name.startswith("word/media/") and Path(name).suffix.lower() in allowed
        ]
        for image_index, name in enumerate(media_names, 1):
            image_path = Path(temporary) / f"image-{image_index}{Path(name).suffix.lower()}"
            image_path.write_bytes(archive.read(name))
            image_pages, image_issues = ocr_adapter.extract(image_path)
            issues.extend(f"Word 内嵌图片 {image_index}：{issue}" for issue in image_issues)
            for page in image_pages:
                page = dict(page)
                page["page_number"] = image_index
                page["page_count"] = len(media_names)
                page["embedded_image_index"] = image_index
                pages.append(page)
    return pages, issues


def _is_audit_material(path: Path, pages: list[dict[str, Any]]) -> bool:
    hint = f"{path.parent.name} {path.name}"
    if any(token in hint for token in AUDIT_HINTS):
        return True
    sample = " ".join(
        str(block.get("text", ""))
        for page in pages[:12]
        for block in page.get("blocks", [])
    )
    return "审计报告" in sample or ("资产负债表" in sample and "利润表" in sample)


def _material_kind(path: Path, pages: list[dict[str, Any]]) -> str:
    if _is_audit_material(path, pages):
        return "审计财务报告"
    hint = f"{path.parent.name} {path.name}"
    text = " ".join(
        str(block.get("text", "")) for page in pages[:3] for block in page.get("blocks", [])
    )
    if "营业执照" in hint or "营业执照" in text:
        return "营业执照"
    if "天眼查" in hint or "天眼查" in text or "企业信用报告" in hint:
        return "企业信息报告"
    return "其他企业资料"


def _all_text(pages: list[dict[str, Any]]) -> str:
    return "\n".join(
        str(block.get("text", "")) for page in pages for block in page.get("blocks", [])
    )


def _company_name(text: str) -> str:
    candidates = re.findall(r"([^\n：:，,。；;()（）]{2,60}?(?:股份有限公司|有限责任公司|有限公司))", text)
    cleaned = []
    for item in candidates:
        item = re.sub(r"^(?:系由|系|本公司|公司名称|名称|被审计单位|被评估单位)\s*[：:]?", "", item).strip()
        if "会计师事务所" not in item and "律师事务所" not in item:
            cleaned.append(item)
    if not cleaned:
        return ""
    return max(set(cleaned), key=lambda item: (cleaned.count(item), len(item)))


def _one_field(text: str, patterns: tuple[str, ...]) -> str:
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if match:
            return re.sub(r"\s+", " ", match.group(1)).strip(" ：:")[:300]
    return ""


def _metadata_from_pages(pages: list[dict[str, Any]], path: Path | None = None) -> dict[str, str]:
    text = _all_text(pages)
    report_date_match = re.search(r"(20\d{2})年\s*(\d{1,2})月\s*(\d{1,2})日", text)
    report_date = ""
    if report_date_match:
        report_date = f"{report_date_match.group(1)}-{int(report_date_match.group(2)):02d}-{int(report_date_match.group(3)):02d}"
    company_name = _company_name(text)
    if not company_name and path is not None:
        filename_candidates = re.findall(r"([^\\/（）()_+\-]{2,60}?(?:股份有限公司|有限责任公司|有限公司))", path.stem)
        company_name = max(filename_candidates, key=len) if filename_candidates else ""
    return {
        "company_name": company_name,
        "credit_code": _one_field(text, (r"(?:统一社会信用代码|社会信用代码)\s*[：:]?\s*([0-9A-Z]{18})",)),
        "legal_representative": _one_field(text, (r"(?:法定代表人|法定代表人姓名)\s*[：:]?\s*([^\n，,；;]{2,40})",)),
        "address": _one_field(text, (r"(?:住所|注册地址|住\s*所|地址)\s*[：:]?\s*([^\n]{5,180})",)),
        "business_scope": _one_field(text, (r"经营范围\s*[：:]?\s*([^\n]{8,300})",)),
        "registered_capital": _one_field(text, (r"(?:注册资本|注册资金)\s*[：:]?\s*([^\n，,；;]{2,80})",)),
        "report_date": report_date,
    }


def _name_key(value: str) -> str:
    return re.sub(r"[\s（）()，,。·-]", "", value or "").replace("有限责任公司", "有限公司")


def _merge_project_metadata(audit: dict[str, str], materials: list[dict[str, Any]]) -> tuple[dict[str, str], list[dict[str, Any]]]:
    """Join company documents only when their recognised subject matches audit."""
    result = dict(audit)
    audit_name = _name_key(audit.get("company_name", ""))
    matched: list[dict[str, Any]] = []
    for material in materials:
        candidate = material.get("metadata", {})
        candidate_name = _name_key(candidate.get("company_name", ""))
        if not audit_name or not candidate_name:
            continue
        if audit_name != candidate_name:
            continue
        matched.append(material)
        for key, value in candidate.items():
            if key != "report_date" and value and not result.get(key):
                result[key] = value
    return result, matched


def _cached_pages(cache_dir: Path, path: Path) -> list[dict[str, Any]] | None:
    cache_file = cache_dir / f"{_sha256(path)}.json"
    if not cache_file.is_file():
        return None
    try:
        payload = json.loads(cache_file.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return None
    # A scanned PDF can produce one placeholder entry per page even though
    # local text extraction found neither text nor tables.  Older versions
    # cached that placeholder list and then treated the page count as an OCR
    # success forever.  A cache hit must contain actual evidence, not merely
    # page shells.
    return payload if isinstance(payload, list) and _has_usable_extraction(payload) else None


def _has_usable_extraction(pages: list[dict[str, Any]]) -> bool:
    """Return whether OCR/native extraction contains any reviewable content."""
    for page in pages:
        if any(str(block.get("text") or "").strip() for block in page.get("blocks", [])):
            return True
        for table in page.get("tables", []):
            if any(str(cell.get("text") or "").strip() for cell in table.get("cells", [])):
                return True
    return False


def _save_cache(cache_dir: Path, path: Path, pages: list[dict[str, Any]]) -> Path:
    cache_dir.mkdir(parents=True, exist_ok=True)
    target = cache_dir / f"{_sha256(path)}.json"
    target.write_text(json.dumps(pages, ensure_ascii=False), encoding="utf-8")
    return target


def _is_template_output_ready(path: Path) -> bool:
    """Avoid redoing completed template copies during bounded desktop runs."""
    if not path.is_file():
        return False
    try:
        with ZipFile(path) as archive:
            return "AI映射规则" in archive.read("xl/workbook.xml").decode("utf-8", errors="ignore")
    except (KeyError, OSError):
        return False


def generate_material_workbooks(
    root: Path,
    output_dir: Path,
    template_dir: Path | None = None,
    skip_ready: bool = False,
    progress_callback: Callable[[str, str, str, int], None] | None = None,
    cache_dir: Path | None = None,
    target_company_name: str | None = None,
) -> dict[str, Any]:
    root = root.resolve()
    output_dir = output_dir.resolve()
    _load_local_env()
    template_bundle = load_workbook_template_bundle(template_dir)
    cache_dir = (cache_dir or (output_dir / "ocr_cache")).resolve()
    # Do not overwrite a workbook the reviewer may currently have open.
    # Each complete intake is emitted to a separate, explicit deliverable
    # folder while the SHA OCR cache remains shared and reusable.
    workbook_root = output_dir / "workbooks_generic_safe_v6"
    documents: list[dict[str, Any]] = []
    ocr_adapter = create_ocr_adapter(os.environ)
    ocr_provider = str(os.environ.get("APPRAISAL_OCR_PROVIDER") or "aliyun").strip().lower()
    sources = sorted(path for path in root.rglob("*") if path.is_file() and path.suffix.lower() in MATERIAL_SUFFIXES)
    if progress_callback:
        progress_callback("inventory", "completed", f"材料清点完成，共发现 {len(sources)} 份可解析材料（PDF、图片、Word、Excel、PPT）", 30)
        progress_callback("ocr_materials", "running", "正在读取原生表格/文本，并对 PDF 和图片执行版式 OCR、记录页码和表格坐标", 31)
    material_records: list[dict[str, Any]] = []
    for source_index, source in enumerate(sources, start=1):
        cached = _cached_pages(cache_dir, source)
        suffix = source.suffix.lower()
        if cached is not None:
            parsing_method = "OCR 缓存"
        elif suffix not in OCR_SUFFIXES:
            parsing_method = "原生 Office 结构读取"
        elif ocr_provider == "paddle":
            parsing_method = "PaddleOCR 版式识别"
        elif ocr_provider == "aliyun":
            parsing_method = "阿里云文档智能版式 OCR"
        elif ocr_adapter is None:
            parsing_method = "本地文本提取（未启用 OCR）"
        else:
            parsing_method = f"{ocr_provider} OCR"
        if progress_callback:
            percent = 31 + int(24 * (source_index - 1) / max(1, len(sources)))
            progress_callback(
                "ocr_materials",
                "running",
                f"正在使用{parsing_method}解析 {source_index}/{len(sources)}：{source.name}",
                percent,
            )
        # Word/Excel/PPT are first parsed as native structures.  PDF and image
        # files still use layout OCR so table/page coordinates remain auditable.
        local_pages = _structured_pages(source)
        local_is_usable = _has_usable_extraction(local_pages)
        pages = local_pages if local_is_usable else []
        source_kind = "native_structure" if local_is_usable else "ocr_document"
        issues: list[str] = []
        if cached is not None:
            pages, source_kind = cached, "ocr_cache"
        elif source.suffix.lower() not in OCR_SUFFIXES:
            # Native document structures have explicit sheet/slide/table
            # locations; keep them even when a cloud OCR account is absent.
            source_kind = "native_structure"
        elif ocr_adapter is None:
            issues.append("该材料需要版式 OCR，但当前 OCR 未配置")
            pages = local_pages
        else:
            ocr_pages, ocr_issues = ocr_adapter.extract(source)
            issues.extend(ocr_issues)
            # Only a non-empty remote result is reusable as OCR evidence.  A
            # failed remote OCR may fall back to useful native PDF text, but
            # that fallback is deliberately not saved in the OCR cache so a
            # later run can retry the configured OCR provider.
            if _has_usable_extraction(ocr_pages):
                pages = ocr_pages
                source_kind = f"{ocr_provider}_ocr"
                _save_cache(cache_dir, source, pages)
            elif local_is_usable:
                pages = local_pages
                source_kind = "native_text_fallback"
            else:
                pages = []
                source_kind = "ocr_failed"
                if not issues:
                    issues.append("OCR 未返回任何可用文字或表格")
        if cached is None and source.suffix.lower() == ".docx":
            image_pages, image_issues = _docx_embedded_image_pages(source, ocr_adapter)
            issues.extend(image_issues)
            if image_pages:
                pages = [*pages, *image_pages]
                source_kind = "native_structure+embedded_image_ocr"
                _save_cache(cache_dir, source, pages)
        if progress_callback:
            percent = 31 + int(24 * source_index / max(1, len(sources)))
            progress_callback(
                "ocr_materials",
                "running" if source_index < len(sources) else "completed",
            (
                f"已用{parsing_method}完成 {source_index}/{len(sources)}：{source.name}"
                if _has_usable_extraction(pages)
                else f"{parsing_method}未识别到有效文字或表格 {source_index}/{len(sources)}：{source.name}"
            ),
            percent,
        )
        material_records.append({
            "source": source,
            "source_file": str(source.relative_to(root)),
            "document_name": source.name,
            "pages": pages,
            "source_kind": source_kind,
            "issues": issues,
            "kind": _material_kind(source, pages or local_pages),
            "metadata": _metadata_from_pages(pages or local_pages, source),
        })

    if progress_callback:
        counts: dict[str, int] = {}
        for record in material_records:
            counts[record["kind"]] = counts.get(record["kind"], 0) + 1
        summary = "、".join(f"{name}{count}份" for name, count in sorted(counts.items())) or "未识别到材料"
        progress_callback("classify_materials", "completed", f"材料分类完成：{summary}", 58)

    supporting = [record for record in material_records if record["kind"] != "审计财务报告"]
    audit_records = [record for record in material_records if record["kind"] == "审计财务报告"]
    records_to_build = audit_records
    if target_company_name is not None:
        detected = sorted({
            str(record["metadata"].get("company_name") or record["source_file"])
            for record in audit_records
        })
        target_key = _subject_key(target_company_name)
        if target_key:
            matches = matching_company_records(
                audit_records,
                target_company_name,
                name_getter=lambda record: record["metadata"].get("company_name"),
                source_getter=lambda record: record.get("source_file", ""),
            )
            if not matches:
                raise RuntimeError(
                    f"材料包中未找到与‘{target_company_name}’可唯一匹配的审计主体；已识别：{'、'.join(detected[:12])}"
                )
            selected_record = min(matches, key=lambda item: (len(item.get("issues", [])), item["source_file"]))
            if not selected_record["metadata"].get("company_name"):
                selected_record["metadata"]["company_name"] = str(target_company_name).strip()
            records_to_build = [selected_record]
        elif len(audit_records) == 1:
            records_to_build = audit_records
        else:
            raise RuntimeError(f"材料包包含多个审计主体，请先填写评估主体全称：{'、'.join(detected[:12])}")
    for record in records_to_build:
        source = record["source"]
        pages = record["pages"]
        issues = record["issues"]
        if not _has_usable_extraction(pages):
            detail = "；".join(str(issue) for issue in issues if issue)[:500]
            raise RuntimeError(
                f"目标公司的审计报告未识别到可用文字或表格：{source.name}"
                + (f"；{detail}" if detail else "")
            )
        statement_tables = [
            table
            for table in tables_from_pages(source.name, pages)
            if table.category in {"资产负债表", "利润表", "现金流量表"}
        ]
        if not statement_tables:
            raise RuntimeError(
                f"目标公司的审计报告已有 OCR 内容，但未识别到资产负债表、利润表或现金流量表：{source.name}"
            )
        target_dir = workbook_root / _safe_name(source)
        if skip_ready and _is_template_output_ready(target_dir / "资产基础法_资产清查.xlsx"):
            continue
        if progress_callback:
            progress_callback("match_subject", "running", f"正在匹配审计报告与营业执照/企业信息：{source.name}", 59)
        project_metadata, matched_supporting = _merge_project_metadata(record["metadata"], supporting)
        if progress_callback:
            progress_callback(
                "match_subject",
                "completed",
                f"主体匹配完成：{project_metadata.get('company_name') or source.stem}，匹配补充材料 {len(matched_supporting)} 份",
                63,
            )
            progress_callback(
                "load_mapping_rules",
                "running",
                f"正在读取系统通用工作簿模板 {template_bundle.version}、逐单元格规则、期间口径和公式依赖",
                63,
            )
        matched_ids = {id(item) for item in matched_supporting}
        # Every archive item is preserved in each output's evidence area.
        # Only an exact normalized subject match may populate metadata;
        # the rest remain visible as non-applicable source documents.
        evidence_materials = []
        for item in supporting:
            evidence = dict(item)
            evidence["match_status"] = (
                "主体已匹配：可作为项目主体资料证据；不参与财务金额取数"
                if id(item) in matched_ids
                else "主体未匹配：已完整 OCR 保留，不写入本审计项目"
            )
            evidence_materials.append(evidence)
        paths = build_template_workbooks(
            output_dir=target_dir,
            document_name=source.name,
            pages=pages,
            asset_template=template_bundle.asset,
            income_template=template_bundle.income,
            project_metadata=project_metadata,
            supporting_materials=evidence_materials,
            progress_callback=progress_callback,
        )
        documents.append(
            {
                "source_file": record["source_file"],
                "source_kind": record["source_kind"],
                "page_count": len(pages),
                "issues": issues,
                "metadata": record["metadata"],
                "workbooks": {role: str(path.relative_to(output_dir)) for role, path in paths.items()},
            }
        )
    manifest = {
        "version": "material_intake.v2",
        "root": str(root),
        "documents": documents,
        "all_materials": [
            {
                "source_file": record["source_file"], "material_type": record["kind"],
                "source_kind": record["source_kind"], "page_count": len(record["pages"]),
                "issues": record["issues"], "metadata": record["metadata"],
            }
            for record in material_records
        ],
    }
    output_dir.mkdir(parents=True, exist_ok=True)
    (output_dir / "生成清单.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return manifest


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="从解压的审计材料生成两类标准工作簿")
    parser.add_argument("root", type=Path)
    parser.add_argument("output_dir", type=Path)
    parser.add_argument(
        "--template-dir",
        type=Path,
        help="可选的版本化通用工作簿模板包目录；默认使用仓库内置模板包",
    )
    parser.add_argument("--skip-ready", action="store_true", help="跳过已生成模板并带映射规则的材料")
    args = parser.parse_args(argv)
    result = generate_material_workbooks(args.root, args.output_dir, args.template_dir, args.skip_ready)
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
