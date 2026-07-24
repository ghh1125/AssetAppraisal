# Asset Appraisal Word Mapping Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a reproducible, read-only analysis pipeline that inventories every marked location in the appraisal Word template, catalogs the PDF and Excel sources, exports all detected PDF tables to XLSX, and generates a traceable Word mapping document.

**Architecture:** The pipeline separates source extraction, semantic mapping, and Word report generation. Source files remain immutable; normalized catalogs and mapping records are written under `outputs/`. Each Word location receives a stable ID, semantic field key, source category, exact source locator when available, calculation/selection rule, and review status.

**Tech Stack:** Bundled Python 3, `lxml`, `python-docx`, `openpyxl`, `pdfplumber`, `pandas`, `pytest`, deterministic OOXML inspection, LibreOffice-based DOCX rendering.

---

## File Structure

- Create `asset_mapping/__init__.py`: package marker.
- Create `asset_mapping/models.py`: typed mapping and source-record dataclasses.
- Create `asset_mapping/docx_inventory.py`: read-only OOXML inventory of placeholders and annotations.
- Create `asset_mapping/pdf_tables.py`: PDF page/table extraction and XLSX export.
- Create `asset_mapping/workbook_catalog.py`: workbook sheet/cell/formula cataloging.
- Create `asset_mapping/field_rules.py`: deterministic context-to-field and source-category rules.
- Create `asset_mapping/mapping_engine.py`: joins Word locations, source catalogs, and field rules.
- Create `asset_mapping/report_builder.py`: produces the mapping DOCX.
- Create `asset_mapping/validation.py`: coverage and provenance checks.
- Create `scripts/build_word_mapping.py`: end-to-end CLI.
- Create `tests/`: focused unit tests and integration coverage checks.
- Create `outputs/通富审计报告_全部表格.xlsx`: derived PDF-table workbook.
- Create `outputs/资产评估Word填充映射说明.docx`: final requested document.
- Create `outputs/资产评估Word填充映射.json`: machine-readable mapping used to generate the Word document.

### Task 1: Data Model and Project Skeleton

**Files:**
- Create: `asset_mapping/__init__.py`
- Create: `asset_mapping/models.py`
- Test: `tests/test_models.py`

- [ ] **Step 1: Write the failing data-model test**

```python
from asset_mapping.models import MappingRecord, SourceKind


def test_mapping_record_serializes_source_and_status():
    record = MappingRecord(
        location_id="DOC-P0005-X01",
        context="XXX有限责任公司拟收购",
        marker="XXX",
        field_key="commissioning_party_name",
        field_name="委托人全称",
        source_kind=SourceKind.MANUAL_INPUT,
        source_file="人工基础信息",
        source_locator="委托方名称",
        rule="由用户录入并在全文复用",
        status="需人工输入",
    )
    data = record.to_dict()
    assert data["source_kind"] == "人工输入"
    assert data["location_id"] == "DOC-P0005-X01"
```

- [ ] **Step 2: Run the model test and verify failure**

Run:

```bash
/Users/ghh/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -m pytest tests/test_models.py -v
```

Expected: `ModuleNotFoundError: No module named 'asset_mapping'`.

- [ ] **Step 3: Implement the dataclasses and source categories**

`SourceKind` must contain `PDF_TABLE`, `REPORTING_XLSX`, `INCOME_XLSX`, `MANUAL_INPUT`, `USER_SELECTION`, `SYSTEM_CALCULATION`, `API`, `LLM_GENERATION`, and `MISSING`. `MappingRecord.to_dict()` must serialize enum values as Chinese labels and retain all provenance fields.

- [ ] **Step 4: Run the model test**

Run: `/Users/ghh/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -m pytest tests/test_models.py -v`

Expected: 1 test passes.

- [ ] **Step 5: Commit the data model**

```bash
git add Code/mcpify/AssetAppraisal/asset_mapping Code/mcpify/AssetAppraisal/tests/test_models.py
git commit -m "feat: add asset mapping data model"
```

### Task 2: Inventory Every Word Location

**Files:**
- Create: `asset_mapping/docx_inventory.py`
- Test: `tests/test_docx_inventory.py`

- [ ] **Step 1: Write coverage tests against the retained template**

```python
from pathlib import Path
from asset_mapping.docx_inventory import inventory_docx


TEMPLATE = Path("资产评估工作流/评估报告版式-沟通标注版.docx")


def test_inventory_covers_all_placeholders_and_highlight_only_blocks():
    inventory = inventory_docx(TEMPLATE)
    assert len(inventory.placeholders) == 127
    assert len({item.paragraph_id for item in inventory.placeholders}) == 65
    assert len(inventory.highlight_only_blocks) == 20
    assert sum(item.part.startswith("word/footer") for item in inventory.placeholders) == 2


def test_inventory_finds_no_comments_or_tracked_changes():
    inventory = inventory_docx(TEMPLATE)
    assert inventory.comment_count == 0
    assert inventory.insertion_count == 0
    assert inventory.deletion_count == 0
```

- [ ] **Step 2: Run tests and verify failure**

Run: `/Users/ghh/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -m pytest tests/test_docx_inventory.py -v`

Expected: import failure for `asset_mapping.docx_inventory`.

- [ ] **Step 3: Implement deterministic OOXML inventory**

Parse `word/document.xml`, headers, footers, text boxes, and table paragraphs. Concatenate visible text within each paragraph before matching `X{2,}` so split runs are counted once. Generate stable location IDs from part, paragraph index, and occurrence index. Record context, marker, table membership, highlight instructions, direct red font, comments, and tracked changes without writing the DOCX.

- [ ] **Step 4: Run inventory tests**

Run: `/Users/ghh/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -m pytest tests/test_docx_inventory.py -v`

Expected: both tests pass with 127 placeholders and 20 highlight-only blocks.

- [ ] **Step 5: Commit the inventory implementation**

```bash
git add Code/mcpify/AssetAppraisal/asset_mapping/docx_inventory.py Code/mcpify/AssetAppraisal/tests/test_docx_inventory.py
git commit -m "feat: inventory appraisal template locations"
```

### Task 3: Extract and Catalog Every PDF Table

**Files:**
- Create: `asset_mapping/pdf_tables.py`
- Test: `tests/test_pdf_tables.py`
- Create: `outputs/通富审计报告_全部表格.xlsx`

- [ ] **Step 1: Write PDF catalog tests**

```python
from pathlib import Path
from asset_mapping.pdf_tables import extract_pdf_table_catalog


PDF = Path("资产评估工作流/通富2025.6.30合并及母公司审计报告.pdf")


def test_pdf_catalog_records_every_page_and_table_provenance():
    catalog = extract_pdf_table_catalog(PDF)
    assert catalog.page_count > 0
    assert len(catalog.page_audit) == catalog.page_count
    assert all(table.page_number >= 1 for table in catalog.tables)
    assert all(table.table_index >= 1 for table in catalog.tables)
    assert all(table.rows for table in catalog.tables)
```

- [ ] **Step 2: Run test and verify failure**

Run: `/Users/ghh/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -m pytest tests/test_pdf_tables.py -v`

Expected: import failure for `asset_mapping.pdf_tables`.

- [ ] **Step 3: Implement two-pass table extraction**

Use `pdfplumber`. First pass uses line-based extraction for ruled financial tables. Second pass uses text-based strategies on pages where the first pass finds no table but the page contains tabular financial text. Normalize blank cells without changing numeric strings. Retain page number, table index, extraction strategy, bounding box, detected title, row count, column count, and review status. Record every page in `page_audit`, including pages with no detected table.

- [ ] **Step 4: Export the derived workbook**

Create an `目录` sheet listing every detected table and source page, a `页面审计` sheet listing every PDF page, and one sheet per extracted table. Use safe sheet names, retain original row order, and include source-page metadata above each table.

- [ ] **Step 5: Run PDF tests and inspect extraction coverage**

Run:

```bash
/Users/ghh/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -m pytest tests/test_pdf_tables.py -v
/Users/ghh/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -c "from pathlib import Path; from asset_mapping.pdf_tables import extract_pdf_table_catalog, export_catalog_xlsx; c=extract_pdf_table_catalog(Path('资产评估工作流/通富2025.6.30合并及母公司审计报告.pdf')); export_catalog_xlsx(c, Path('outputs/通富审计报告_全部表格.xlsx')); print(c.page_count, len(c.tables))"
```

Expected: tests pass; workbook exists and printed counts are nonzero. Pages with uncertain extraction are explicitly marked for review rather than silently omitted.

- [ ] **Step 6: Commit PDF extraction**

```bash
git add Code/mcpify/AssetAppraisal/asset_mapping/pdf_tables.py Code/mcpify/AssetAppraisal/tests/test_pdf_tables.py
git commit -m "feat: extract audit report tables with provenance"
```

### Task 4: Catalog the Two Excel Sources

**Files:**
- Create: `asset_mapping/workbook_catalog.py`
- Test: `tests/test_workbook_catalog.py`

- [ ] **Step 1: Write workbook catalog tests**

```python
from pathlib import Path
from asset_mapping.workbook_catalog import catalog_workbook


def test_catalog_retains_sheet_cell_formula_and_display_value():
    path = Path("资产评估工作流/通富热处理（昆山）有限公司-收益法-20250630.xlsx")
    catalog = catalog_workbook(path)
    assert catalog.sheets
    assert any(cell.formula is not None for sheet in catalog.sheets for cell in sheet.cells)
    assert all(cell.coordinate for sheet in catalog.sheets for cell in sheet.cells)


def test_reporting_workbook_has_searchable_labels():
    path = Path("资产评估工作流/上报表文件_通富昆山_已处理.xlsx")
    catalog = catalog_workbook(path)
    matches = catalog.find_labels(["净资产", "所有者权益", "评估值"])
    assert matches
```

- [ ] **Step 2: Run tests and verify failure**

Run: `/Users/ghh/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -m pytest tests/test_workbook_catalog.py -v`

Expected: import failure for `asset_mapping.workbook_catalog`.

- [ ] **Step 3: Implement workbook cataloging**

Load workbooks twice with `openpyxl`: once with formulas and once with `data_only=True`. Record sheet visibility, merged cells, non-empty cell coordinates, labels, formulas, cached values, number formats, and nearby row/column context. Provide exact-label and normalized fuzzy-label search while keeping the original text available for provenance.

- [ ] **Step 4: Run workbook tests**

Run: `/Users/ghh/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -m pytest tests/test_workbook_catalog.py -v`

Expected: both tests pass.

- [ ] **Step 5: Commit workbook cataloging**

```bash
git add Code/mcpify/AssetAppraisal/asset_mapping/workbook_catalog.py Code/mcpify/AssetAppraisal/tests/test_workbook_catalog.py
git commit -m "feat: catalog appraisal workbooks"
```

### Task 5: Semantic Field Rules and Mapping Engine

**Files:**
- Create: `asset_mapping/field_rules.py`
- Create: `asset_mapping/mapping_engine.py`
- Test: `tests/test_mapping_engine.py`
- Create: `outputs/资产评估Word填充映射.json`

- [ ] **Step 1: Write representative mapping tests**

```python
from asset_mapping.field_rules import classify_context


def test_classifies_manual_and_calculated_fields():
    assert classify_context("二、委托人：XXX有限责任公司", 0).field_key == "commissioning_party_name"
    assert classify_context("十、评估基准日：20XX年XX月XX日", 0).field_key == "valuation_date_year"
    assert classify_context("评估增值XXX万元，增值率XXX", 0).source_kind.value == "系统计算"


def test_classifies_api_and_llm_generation_blocks():
    assert classify_context("被评估单位股权结构及历史沿革（企查查API获取）", 0).source_kind.value == "API获取"
    assert classify_context("盈利模式，SWOT分析", 0).source_kind.value == "大模型生成"
```

- [ ] **Step 2: Run tests and verify failure**

Run: `/Users/ghh/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -m pytest tests/test_mapping_engine.py -v`

Expected: import failure for `asset_mapping.field_rules`.

- [ ] **Step 3: Implement ordered semantic rules**

Rules must use section heading, full paragraph context, marker index, and neighboring labels. They must distinguish commissioning party, appraised entity, transaction type, valuation subject, valuation date, report date, report number, book value, income-approach value, asset-approach value, final selected value, increment, increment rate, tax rates, financial statements, shareholder history, intellectual property, company profile, industry narrative, products, customers/suppliers, SWOT, and comparable companies.

API rules must state suggested provider class, query key, returned fields, query date, review requirement, and manual fallback. LLM rules must state evidence inputs and prohibit generated financial figures.

- [ ] **Step 4: Join rules with exact source locators**

Search the PDF table catalog and workbook catalogs for the field's candidate labels. Record exact file, PDF page/table or workbook sheet/cell when confidence is sufficient. If multiple candidates conflict, record all candidates and status `存在冲突，需复核`. If no exact locator exists, retain the correct source category and status `资料中未定位` rather than guessing.

- [ ] **Step 5: Validate exhaustive coverage and write JSON**

Run:

```bash
/Users/ghh/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -m pytest tests/test_mapping_engine.py -v
/Users/ghh/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 scripts/build_word_mapping.py --mapping-json-only
```

Expected: JSON contains 147 location records, each with a non-empty location ID, field name, source category, rule, and status; 127 are placeholder records and 20 are highlight-only blocks.

- [ ] **Step 6: Commit mapping rules**

```bash
git add Code/mcpify/AssetAppraisal/asset_mapping/field_rules.py Code/mcpify/AssetAppraisal/asset_mapping/mapping_engine.py Code/mcpify/AssetAppraisal/tests/test_mapping_engine.py
git commit -m "feat: map word fields to appraisal sources"
```

### Task 6: Build the Word Mapping Document

**Files:**
- Create: `asset_mapping/report_builder.py`
- Create: `scripts/build_word_mapping.py`
- Test: `tests/test_report_builder.py`
- Create: `outputs/资产评估Word填充映射说明.docx`

- [ ] **Step 1: Write report structure tests**

```python
from pathlib import Path
from docx import Document


def test_generated_report_contains_required_sections():
    path = Path("outputs/资产评估Word填充映射说明.docx")
    doc = Document(path)
    text = "\n".join(p.text for p in doc.paragraphs)
    assert "数据来源说明" in text
    assert "业务字段来源总表" in text
    assert "API获取" in text
    assert "附录A：全部占位符逐项映射" in text
    assert "附录B：黄色标注内容块逐项映射" in text
```

- [ ] **Step 2: Implement the report builder**

Use a restrained business-report style. Create landscape sections for wide mapping tables, repeat header rows, prevent fixed row heights, and use explicit column widths. Group the main table by unique field key. Append all 127 placeholders and 20 highlight-only locations in stable Word order. Include source legends, API notes, calculation rules, conflicts, missing data, and conditional-method notes.

- [ ] **Step 3: Generate the DOCX**

Run:

```bash
/Users/ghh/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 scripts/build_word_mapping.py
```

Expected: the PDF-table XLSX, mapping JSON, and mapping DOCX are generated under `outputs/`; no source file timestamp or checksum changes.

- [ ] **Step 4: Run report tests**

Run: `/Users/ghh/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -m pytest tests/test_report_builder.py -v`

Expected: report structure test passes.

- [ ] **Step 5: Commit report generation**

```bash
git add Code/mcpify/AssetAppraisal/asset_mapping/report_builder.py Code/mcpify/AssetAppraisal/scripts/build_word_mapping.py Code/mcpify/AssetAppraisal/tests/test_report_builder.py
git commit -m "feat: generate appraisal word mapping report"
```

### Task 7: Validate, Render, and Deliver

**Files:**
- Create: `asset_mapping/validation.py`
- Test: `tests/test_validation.py`
- Verify: `outputs/资产评估Word填充映射说明.docx`

- [ ] **Step 1: Implement structural validation**

Validation must fail when any of the 127 placeholders or 20 highlight-only blocks is absent, any location ID is duplicated, any source category is empty, a calculated field lacks a formula, an API field lacks a manual fallback, or an LLM field lacks an evidence/review rule.

- [ ] **Step 2: Run the complete test suite**

Run: `/Users/ghh/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 -m pytest tests -v`

Expected: all tests pass.

- [ ] **Step 3: Verify source immutability**

Compute SHA-256 hashes for the four source files before and after generation and assert they match. Expected: all four hashes are unchanged.

- [ ] **Step 4: Render the generated DOCX**

Run:

```bash
/Users/ghh/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 /Users/ghh/.codex/plugins/cache/openai-primary-runtime/documents/26.715.12143/skills/documents/render_docx.py outputs/资产评估Word填充映射说明.docx --output_dir /tmp/asset-mapping-render --emit_pdf
```

Expected: one PNG per page and a non-empty PDF are produced.

- [ ] **Step 5: Inspect every rendered page**

Check every page image at 100% for clipped text, table overflow, repeated headers, page breaks, Chinese glyphs, and readable source locators. Fix the DOCX builder and re-render until all pages pass.

- [ ] **Step 6: Final consistency audit**

Confirm the Word report and JSON contain identical record counts and location IDs, the PDF-table workbook index matches extracted sheets, and every unresolved field is clearly labeled as manual input, user selection, API, LLM generation, conflict, or missing evidence.

- [ ] **Step 7: Commit validation**

```bash
git add Code/mcpify/AssetAppraisal/asset_mapping/validation.py Code/mcpify/AssetAppraisal/tests/test_validation.py
git commit -m "test: validate appraisal mapping deliverables"
```
