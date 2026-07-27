# Optional Materials and Highlighted Placeholders Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Generate a review Word report from any available subset of inputs, preserve and yellow-highlight unresolved placeholders, and export a page-aware detailed generation issue list.

**Architecture:** Keep business decisions in pure `domain/` functions, isolate DOCX highlighting and XLSX issue export in adapters, and let the API/pipeline orchestrate optional sources. Known mapped omissions produce structured issues before Word generation; a final DOCX scan highlights and registers any unmapped placeholders.

**Tech Stack:** Python 3.11, Pydantic, FastAPI, lxml/OOXML, python-docx, openpyxl, Vue 3, Ant Design Vue, pytest, Vite.

---

### Task 1: Missing-value and generation-issue business rules

**Files:**
- Create: `demo/domain/generation_issues.py`
- Modify: `demo/domain/replacement.py`
- Modify: `demo/domain/financial_matching.py`
- Test: `demo/tests/test_generation_issues.py`
- Test: `demo/tests/test_word_replacement.py`
- Test: `demo/tests/test_financial_matching.py`

- [ ] **Step 1: Write failing tests for placeholder selection and issue records**

```python
def test_missing_field_uses_original_marker_and_creates_location_issue():
    location = {
        "location_id": "DOCUMENT-P0001-X01",
        "record_type": "占位符",
        "field_key": "company_name",
        "field_name": "公司名称",
        "marker": "XXX",
        "context": "公司名称：XXX",
        "source_kind": "人工输入",
        "source_file": "人工基础信息",
        "source_locator": "公司名称",
    }
    replacements = build_replacements([location], {})
    issues = issues_for_missing_locations([location], {})
    assert replacements[location["location_id"]] == "XXX"
    assert issues[0]["location_id"] == location["location_id"]
    assert issues[0]["current_text"] == "XXX"
    assert issues[0]["expected_source"] == "人工输入"
```

```python
def test_blank_configured_table_uses_highlightable_placeholders():
    matrix = blank_configured_table(
        {
            "header": ["项目", "2025年"],
            "rows": [{"label": "资产总计", "cells": ["B2"]}],
        },
        placeholder="XXX",
    )
    assert matrix == [["项目", "2025年"], ["资产总计", "XXX"]]
```

- [ ] **Step 2: Run tests and verify RED**

Run:

```bash
.venv/bin/pytest demo/tests/test_generation_issues.py demo/tests/test_word_replacement.py demo/tests/test_financial_matching.py -q
```

Expected: failure because `issues_for_missing_locations` does not exist and missing replacements still use `human_fill`.

- [ ] **Step 3: Implement pure issue and placeholder rules**

Create `demo/domain/generation_issues.py` with:

```python
from __future__ import annotations

import hashlib
from typing import Any


def missing_marker(location: dict[str, Any]) -> str:
    marker = str(location.get("marker", ""))
    return marker if "XX" in marker.upper() else "XXX"


def issues_for_missing_locations(
    locations: list[dict[str, Any]],
    fields: dict[str, Any],
) -> list[dict[str, Any]]:
    issues = []
    for location in locations:
        key = str(location["field_key"])
        if fields.get(key) not in (None, "", [], {}):
            continue
        location_id = str(location["location_id"])
        marker = missing_marker(location)
        digest = hashlib.sha1(f"{location_id}:{key}".encode()).hexdigest()[:10]
        issues.append({
            "issue_id": f"GEN-{digest}",
            "priority": "高",
            "category": "missing_field",
            "page_number": "",
            "page_basis": "unavailable",
            "location_id": location_id,
            "location_type": "段落",
            "location_description": str(location.get("context", "")),
            "field_key": key,
            "field_name": str(location.get("field_name", key)),
            "current_text": marker,
            "problem": "指定来源未匹配到可用值",
            "expected_source": str(location.get("source_kind", "")),
            "source_file": str(location.get("source_file", "")),
            "source_locator": str(location.get("source_locator", "")),
            "suggestion": "补充对应材料或人工确认后替换黄色占位符",
            "status": "待人工处理",
        })
    return issues
```

Update `build_replacements()` to use `missing_marker(item)` rather than `human_fill()`. Extend `blank_configured_table(spec, placeholder="")` so callers can explicitly request `XXX`.

- [ ] **Step 4: Run focused tests and verify GREEN**

Run:

```bash
.venv/bin/pytest demo/tests/test_generation_issues.py demo/tests/test_word_replacement.py demo/tests/test_financial_matching.py -q
```

Expected: all focused tests pass.

- [ ] **Step 5: Commit**

```bash
git add demo/domain/generation_issues.py demo/domain/replacement.py demo/domain/financial_matching.py demo/tests/test_generation_issues.py demo/tests/test_word_replacement.py demo/tests/test_financial_matching.py
git commit -m "feat: model unresolved report fields"
```

### Task 2: DOCX placeholder highlighting and final scan

**Files:**
- Modify: `demo/adapters/word.py`
- Test: `demo/tests/test_word_replacement.py`

- [ ] **Step 1: Write failing DOCX highlighting tests**

```python
def test_missing_placeholder_is_preserved_and_only_marker_is_highlighted(tmp_path):
    template = tmp_path / "template.docx"
    document = Document()
    document.add_paragraph("公司名称：XXX")
    document.save(template)
    location = inventory_template(template)[0]
    output = tmp_path / "output.docx"
    fill_template(template, output, {location["location_id"]: "XXX"})
    findings = highlight_unresolved_placeholders(output)
    paragraph = Document(output).paragraphs[0]
    assert paragraph.text == "公司名称：XXX"
    assert [run.text for run in paragraph.runs if run.font.highlight_color] == ["XXX"]
    assert findings[0]["location_id"] == location["location_id"]
```

```python
def test_missing_table_cell_is_highlighted_and_reported(tmp_path):
    template = tmp_path / "table.docx"
    document = Document()
    document.add_table(rows=2, cols=2)
    document.save(template)
    output = tmp_path / "output.docx"
    fill_template(
        template,
        output,
        {},
        table_replacements={0: [["项目", "金额"], ["资产总计", "XXX"]]},
    )
    findings = highlight_unresolved_placeholders(output)
    assert any(item["location_type"] == "表格单元格" for item in findings)
    assert Document(output).tables[0].cell(1, 1).text == "XXX"
```

- [ ] **Step 2: Run tests and verify RED**

Run:

```bash
.venv/bin/pytest demo/tests/test_word_replacement.py -q
```

Expected: import failure because `highlight_unresolved_placeholders` does not exist.

- [ ] **Step 3: Implement OOXML marker splitting, highlighting, and scanning**

Add to `demo/adapters/word.py`:

```python
def highlight_unresolved_placeholders(path: Path) -> list[dict[str, Any]]:
    """Highlight unresolved markers in all Word parts and return their locations."""
    # Read every PART_RE part, split runs around PLACEHOLDER matches while
    # deep-copying rPr, add <w:highlight w:val="yellow"> only to marker runs,
    # derive DOCUMENT-P####-X## IDs for paragraph markers, and derive
    # TABLE-T##-R##-C## IDs for table-cell markers.
```

The implementation must:

- preserve surrounding run formatting;
- highlight only marker text;
- scan body, headers, footers, footnotes and endnotes;
- return `part`, paragraph index, occurrence, context, table index, row and column when available;
- be idempotent;
- leave `unresolved_placeholders()` as a read-only diagnostic rather than a hard gate.

Update the `strip_yellow_only` path so a replacement matching `20XX|X{2,}` is inserted through `_replace_yellow_annotation()` instead of discarded.

- [ ] **Step 4: Run focused tests and verify GREEN**

Run:

```bash
.venv/bin/pytest demo/tests/test_word_replacement.py -q
```

Expected: all DOCX replacement tests pass.

- [ ] **Step 5: Commit**

```bash
git add demo/adapters/word.py demo/tests/test_word_replacement.py
git commit -m "feat: highlight unresolved Word placeholders"
```

### Task 3: Optional and resilient source readers

**Files:**
- Modify: `demo/adapters/excel.py`
- Modify: `demo/adapters/materials.py`
- Modify: `demo/run.py`
- Test: `demo/tests/test_adapters.py`
- Test: `demo/tests/test_demo_run.py`

- [ ] **Step 1: Write failing tests for absent and structurally different sources**

```python
def test_optional_excel_cell_returns_issue_instead_of_raising(tmp_path):
    workbook_path = tmp_path / "other-layout.xlsx"
    workbook = Workbook()
    workbook.active.title = "资产负债表"
    workbook.save(workbook_path)
    values, issues = try_read_cells(workbook_path, ["06N_资产负债表!L75"])
    assert values == {}
    assert "缺少工作表" in issues[0]
```

```python
def test_run_project_generates_with_only_one_manual_field(tmp_path):
    result = run_project(
        Path("demo/projects/tongfu.yaml"),
        output_dir=tmp_path,
        offline=True,
        manual_inputs_override={"target_company_name": "示例有限公司"},
        source_overrides={
            "audit_pdf": None,
            "reference_report": None,
            "audited_financials": None,
            "income_workbook": None,
            "reporting_workbook": None,
        },
    )
    assert result.report_path.exists()
    assert "XXX" in "\n".join(
        paragraph.text for paragraph in Document(result.report_path).paragraphs
    )
```

- [ ] **Step 2: Run tests and verify RED**

Run:

```bash
.venv/bin/pytest demo/tests/test_adapters.py demo/tests/test_demo_run.py -q
```

Expected: `try_read_cells` missing and `Path(None)`/missing source failures.

- [ ] **Step 3: Add safe reader APIs and optional source override semantics**

Add safe adapter functions:

```python
def try_read_cells(
    path: Path | None,
    locators: list[str],
) -> tuple[dict[str, object], list[str]]:
    if path is None:
        return {}, ["缺少来源文件"]
    try:
        return read_cells(path, locators), []
    except KeyError as exc:
        return {}, [f"缺少工作表：{exc}"]
    except (OSError, ValueError, BadZipFile) as exc:
        return {}, [f"材料无法读取：{exc}"]
```

In `run_project()` change `source_overrides` to `dict[str, Path | None]`. A key explicitly mapped to `None` removes the configured sample source. Wrap each configured financial table, financial field, long-term table and material field independently; on failure add a structured issue, assign missing evidence, and continue. Use `blank_configured_table(spec, placeholder="XXX")` when a configured table cannot be read.

Remove the hard unresolved-placeholder exception. After saving the Word, run `highlight_unresolved_placeholders()`.

- [ ] **Step 4: Run focused tests and verify GREEN**

Run:

```bash
.venv/bin/pytest demo/tests/test_adapters.py demo/tests/test_demo_run.py -q
```

Expected: all focused tests pass and the report exists with absent sources.

- [ ] **Step 5: Commit**

```bash
git add demo/adapters/excel.py demo/adapters/materials.py demo/run.py demo/tests/test_adapters.py demo/tests/test_demo_run.py
git commit -m "feat: continue when project materials are absent"
```

### Task 4: Detailed generation issue artifacts and page mapping

**Files:**
- Create: `demo/adapters/generation_issues.py`
- Modify: `demo/domain/generation_issues.py`
- Modify: `demo/domain/template_pagination.py`
- Modify: `demo/run.py`
- Modify: `demo/pipeline.py`
- Test: `demo/tests/test_generation_issues.py`
- Test: `demo/tests/test_pipeline.py`

- [ ] **Step 1: Write failing tests for page-aware issue export**

```python
def test_generation_issue_export_has_required_columns(tmp_path):
    output = export_generation_issues(
        tmp_path / "生成问题清单.xlsx",
        [{
            "issue_id": "GEN-1",
            "priority": "高",
            "page_number": 9,
            "page_basis": "generated_report",
            "location_id": "DOCUMENT-P0100-X01",
            "location_type": "段落",
            "location_description": "公司名称：XXX",
            "field_key": "company_name",
            "field_name": "公司名称",
            "current_text": "XXX",
            "problem": "指定来源未匹配到可用值",
            "expected_source": "人工输入",
            "source_file": "",
            "source_locator": "公司名称",
            "suggestion": "补充公司名称",
            "status": "待人工处理",
            "category": "missing_field",
        }],
    )
    sheet = load_workbook(output, read_only=True)["生成问题"]
    assert sheet["C2"].value == 9
    assert sheet["E2"].value == "DOCUMENT-P0100-X01"
```

```python
def test_pipeline_without_pdf_exports_report_and_issue_list(tmp_path):
    result = run_pipeline(
        project_config=Path("demo/projects/tongfu.yaml"),
        pdf_path=None,
        output_dir=tmp_path,
        ocr_adapter=None,
        source_overrides={
            "audit_pdf": None,
            "reference_report": None,
            "audited_financials": None,
            "income_workbook": None,
            "reporting_workbook": None,
        },
        manual_inputs_override={"target_company_name": "示例有限公司"},
    )
    assert result.report_path.exists()
    assert result.ocr_workbook_path is None
    assert (tmp_path / "生成问题清单.xlsx").exists()
    assert (tmp_path / "生成问题清单.json").exists()
```

- [ ] **Step 2: Run tests and verify RED**

Run:

```bash
.venv/bin/pytest demo/tests/test_generation_issues.py demo/tests/test_pipeline.py -q
```

Expected: missing exporter and `pdf_path=None` path resolution failure.

- [ ] **Step 3: Implement issue enrichment, XLSX/JSON export, and optional PDF pipeline**

Create `export_generation_issues()` with the 16 columns in the design and conventional header formatting. Add:

```python
def apply_page_locations(
    issues: list[dict[str, Any]],
    generated_pages: dict[str, int | str],
    template_pages: dict[str, int | str],
) -> list[dict[str, Any]]:
    # Prefer generated report pages, fall back to template pages, otherwise
    # retain an empty page and page_basis="unavailable".
```

Change `run_pipeline(pdf_path: Path | None, ...)` and `PipelineResult.ocr_workbook_path` to optional. If PDF is absent:

- record OCR/export nodes as skipped;
- use an empty normalized OCR contract;
- omit PDF hash and OCR workbook from outputs;
- continue source/API/LLM routing;
- generate the Word;
- highlight unresolved markers;
- map report pages when the page reader is available;
- merge mapped missing issues and fallback-scan issues;
- export `生成问题清单.xlsx` and `生成问题清单.json`;
- add `generation_validation` and both artifacts to the manifest;
- prevent final candidate creation when unresolved issues remain.

- [ ] **Step 4: Run focused tests and verify GREEN**

Run:

```bash
.venv/bin/pytest demo/tests/test_generation_issues.py demo/tests/test_pipeline.py -q
```

Expected: optional-PDF pipeline and issue export tests pass.

- [ ] **Step 5: Commit**

```bash
git add demo/adapters/generation_issues.py demo/domain/generation_issues.py demo/domain/template_pagination.py demo/run.py demo/pipeline.py demo/tests/test_generation_issues.py demo/tests/test_pipeline.py
git commit -m "feat: export page-aware generation issues"
```

### Task 5: Optional-upload HTTP API

**Files:**
- Modify: `demo/api_server.py`
- Create: `demo/tests/test_api_server.py`

- [ ] **Step 1: Write failing API tests**

```python
def test_api_accepts_manual_only_run(monkeypatch):
    monkeypatch.setattr(api_server, "_execute_run", lambda *args, **kwargs: None)
    response = TestClient(api_server.app).post(
        "/api/v1/asset-appraisal/runs",
        data={"inputs": json.dumps({"target_company_name": "示例有限公司"})},
    )
    assert response.status_code == 202
```

```python
def test_api_accepts_xlsm_income_workbook(monkeypatch, tmp_path):
    monkeypatch.setattr(api_server, "_execute_run", lambda *args, **kwargs: None)
    response = TestClient(api_server.app).post(
        "/api/v1/asset-appraisal/runs",
        data={"inputs": "{}"},
        files={"income_workbook": ("收益法.xlsm", b"placeholder", "application/vnd.ms-excel.sheet.macroEnabled.12")},
    )
    assert response.status_code == 202
```

- [ ] **Step 2: Run tests and verify RED**

Run:

```bash
.venv/bin/pytest demo/tests/test_api_server.py -q
```

Expected: 422 responses because every upload is currently required.

- [ ] **Step 3: Implement optional upload storage and validation**

Use:

```python
pdf: UploadFile | None = File(None)
reference_report: UploadFile | None = File(None)
audited_financials: UploadFile | None = File(None)
income_workbook: UploadFile | None = File(None)
reporting_workbook: UploadFile | None = File(None)
```

Validate only supplied files; accept `.xlsx` and `.xlsm` for workbook roles. Require at least one uploaded material or one non-empty manual input. Store only supplied files, pass `None` for missing source roles, derive the run ID from the first supplied filename or assessed-company name, and skip OCR cache lookup when PDF is absent. Register `生成问题清单.xlsx` and `生成问题清单.json` in `_artifact_list()`.

- [ ] **Step 4: Run focused tests and verify GREEN**

Run:

```bash
.venv/bin/pytest demo/tests/test_api_server.py -q
```

Expected: both optional-upload tests pass.

- [ ] **Step 5: Commit**

```bash
git add demo/api_server.py demo/tests/test_api_server.py
git commit -m "feat: accept partial appraisal uploads"
```

### Task 6: Frontend optional-material workflow

**Files:**
- Create: `frontend/src/domain/submission.js`
- Create: `frontend/src/domain/submission.test.js`
- Modify: `frontend/package.json`
- Modify: `frontend/src/api/asset-appraisal.js`
- Modify: `frontend/src/views/asset-appraisal/index.vue`
- Modify: `frontend/src/i18n/index.js`

- [ ] **Step 1: Write and run a failing partial-submission test**

Create a Node test for a pure `canSubmitPartial(files, inputs)` helper. Cover manual-only, one-file-only, and completely empty submissions. Add a `test:submission` package script, run it, and verify RED because the helper does not exist.

- [ ] **Step 2: Implement the pure submission rule and conditional multipart fields**

Implement `canSubmitPartial()` and use it from the Vue component so submission becomes available when any file exists or any meaningful manual input is non-empty. Change upload accepts to `.xlsx,.xlsm` and remove `required` wording from all five material captions.

Use:

```javascript
if (pdf) form.append('pdf', pdf)
if (referenceReport) form.append('reference_report', referenceReport)
if (auditedFinancials) form.append('audited_financials', auditedFinancials)
if (incomeWorkbook) form.append('income_workbook', incomeWorkbook)
if (reportingWorkbook) form.append('reporting_workbook', reportingWorkbook)
```

Update Chinese and English text to explain that missing material leaves a yellow `XXX` and that OCR cache is checked only after a PDF is selected.

- [ ] **Step 3: Run the frontend test and build**

Run:

```bash
npm run test:submission
npm run build
```

Expected: the pure behavior tests pass and the Vite build succeeds.

- [ ] **Step 4: Commit**

```bash
git add frontend/package.json frontend/src/domain/submission.js frontend/src/domain/submission.test.js frontend/src/api/asset-appraisal.js frontend/src/views/asset-appraisal/index.vue frontend/src/i18n/index.js
git commit -m "feat: allow optional appraisal materials in UI"
```

### Task 7: Contracts, fixtures, documentation, and end-to-end verification

**Files:**
- Modify: `demo/data_manifest.yaml`
- Modify: `demo/fixtures/workflow_cases.yaml`
- Modify: `demo/expected/workflow_cases.yaml`
- Modify: `demo/tests/test_contracts.py`
- Modify: `demo/README.md`
- Modify: `README.md`
- Modify: `demo/CHANGELOG.md`

- [ ] **Step 1: Register versions and representative boundary examples**

Add:

```json
"optional_sources": "optional_sources.v1",
"generation_issues": "generation_issues.v1",
"placeholder_policy": "placeholder_policy.v1"
```

Document that PDF and all business materials are optional; absence creates highlighted placeholders and generation issues, while invalid workflow/template structure remains fatal.

- [ ] **Step 2: Run the full verification suite**

Run:

```bash
.venv/bin/pytest -q
.venv/bin/python -m compileall -q demo
npm --prefix frontend run build
git diff --check
```

Expected: all Python tests pass, compilation succeeds, Vite build succeeds, and Git reports no whitespace errors.

- [ ] **Step 3: Run a real no-PDF smoke test**

Run the API/pipeline with one latest-material income workbook and one manual company name, with OCR/API/LLM disabled. Verify:

- `资产评估报告_待复核.docx` exists;
- unresolved fields remain `XXX` and are yellow;
- `生成问题清单.xlsx` and JSON exist;
- at least one issue has a Word page or explicit unavailable page basis;
- no `资产评估报告_最终候选.docx` exists;
- no customer input file is modified.

- [ ] **Step 4: Commit**

```bash
git add README.md demo/README.md demo/CHANGELOG.md demo/data_manifest.yaml demo/fixtures/workflow_cases.yaml demo/expected/workflow_cases.yaml demo/tests/test_contracts.py
git commit -m "docs: document partial-material report generation"
```

- [ ] **Step 5: Finish the branch**

Run the full verification commands again, merge the feature branch into `main`, remove the worktree, and push `main` to `origin`.
