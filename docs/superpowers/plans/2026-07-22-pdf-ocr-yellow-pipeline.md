# PDF OCR Yellow-Routed Pipeline Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a reusable Demo pipeline that converts a scanned PDF into a structured OCR workbook, routes the template's 20 yellow fields strictly by the yellow instructions, and generates a separate Word report plus auditable outputs.

**Architecture:** Keep portable Pydantic contracts and pure routing/matching rules in `demo/domain/`; keep PaddleOCR, Bailian GLM, Qichacha, filesystem, XLSX, and Word operations in injected adapters. The pipeline validates the exact 7/5/6/2 yellow route partition before doing any work and rejects cross-route fallback.

**Tech Stack:** Python 3.11, Pydantic 2, PaddleOCR 3 PP-StructureV3, OpenAI-compatible Bailian `glm-5.2`, python-docx/lxml, openpyxl for the Demo XLSX adapter, pytest; artifact-tool for final workbook inspection/render QA.

---

## File map

- Create `demo/domain/yellow_routing.py`: exact yellow route constants and pure validation/merge rules.
- Create `demo/domain/ocr_normalization.py`: SDK-neutral OCR normalization helpers.
- Create `demo/domain/financial_matching.py`: layout-independent table/label/period/unit matching.
- Create `demo/domain/field_validation.py`: required-field and evidence validation.
- Create `demo/adapters/paddle_ocr.py`: PP-StructureV3 integration only.
- Create `demo/adapters/ocr_workbook.py`: OCR XLSX export/read adapter only.
- Create `demo/adapters/bailian_glm.py`: Bailian HTTP adapter only.
- Create `demo/pipeline.py`: dependency composition and CLI only.
- Create `demo/rules/financial_aliases.v1.yaml`: configurable aliases and units.
- Create `demo/prompts/yellow_narratives.v1.txt`: GLM instruction template.
- Create `demo/prompts/yellow_narratives_output.v1.json`: GLM structured-output schema.
- Create `demo/fixtures/ocr_cases.yaml`: ten de-identified normal/boundary cases.
- Create `demo/expected/ocr_cases.yaml`: expected normalized results.
- Modify `demo/schemas.py`: node input/output and evidence contracts.
- Modify `demo/projects/tongfu.yaml`: strict `yellow_routes` and OCR/LLM settings without secrets.
- Modify `demo/workflow.yaml`: explicit end-to-end nodes and checkpoint.
- Modify `demo/run.py`: delegate new end-to-end mode without putting business rules in the CLI.
- Modify `demo/README.md`, `demo/data_manifest.yaml`, `demo/CHANGELOG.md`, `.env.example`, `pyproject.toml`.
- Add focused tests under `demo/tests/` for each new unit and an end-to-end fixture pipeline.

### Task 1: Exact yellow routing business rule

**Files:**
- Create: `demo/domain/yellow_routing.py`
- Create: `demo/tests/test_yellow_routing.py`
- Modify: `demo/projects/tongfu.yaml`

- [ ] **Step 1: Write failing route-partition tests**

```python
from pathlib import Path

import pytest

from demo.adapters.word import inventory_template
from demo.domain.yellow_routing import RouteKind, route_fields, validate_yellow_routes


def test_template_yellow_fields_are_partitioned_exactly_once():
    inventory = inventory_template(Path("资产评估工作流/评估报告版式-沟通标注版.docx"))
    yellow_ids = {row["location_id"] for row in inventory if row["record_type"] == "黄色标注内容块"}
    routes = route_fields()
    validate_yellow_routes(yellow_ids, routes)
    assert len(yellow_ids) == 20
    assert sum(route.kind is RouteKind.LLM for route in routes.values()) == 7
    assert sum(route.kind is RouteKind.QICHACHA for route in routes.values()) == 5
    assert sum(route.kind is RouteKind.OCR_XLSX for route in routes.values()) == 6
    assert sum(route.kind is RouteKind.NODE_INPUT for route in routes.values()) == 2


def test_unknown_or_duplicate_yellow_location_is_rejected():
    routes = route_fields()
    with pytest.raises(ValueError, match="黄色位置不一致"):
        validate_yellow_routes({"DOCUMENT-P9999-H01"}, routes)
```

- [ ] **Step 2: Verify RED**

Run: `UV_CACHE_DIR=/private/tmp/asset-appraisal-uv-cache uv run pytest demo/tests/test_yellow_routing.py -q`

Expected: collection fails because `demo.domain.yellow_routing` does not exist.

- [ ] **Step 3: Implement immutable routes and strict validation**

```python
class RouteKind(StrEnum):
    LLM = "llm"
    QICHACHA = "qichacha_api"
    OCR_XLSX = "pdf_ocr_xlsx"
    NODE_INPUT = "node_input"


@dataclass(frozen=True)
class YellowRoute:
    location_id: str
    field_key: str
    kind: RouteKind


def validate_yellow_routes(yellow_ids: set[str], routes: Mapping[str, YellowRoute]) -> None:
    route_ids = {route.location_id for route in routes.values()}
    if yellow_ids != route_ids or len(routes) != 20:
        raise ValueError("黄色位置不一致")
```

Declare the exact 7 GLM, 5 Qichacha, 6 OCR/XLSX, and 2 node-input fields from the approved design. Add the same route table to `tongfu.yaml` for audit visibility; runtime validation compares configuration against the immutable approved domain rule.

- [ ] **Step 4: Verify GREEN and full regression**

Run: `UV_CACHE_DIR=/private/tmp/asset-appraisal-uv-cache uv run pytest demo/tests/test_yellow_routing.py demo/tests/test_template_mapping.py -q`

Expected: both test modules pass.

- [ ] **Step 5: Commit**

```bash
git add demo/domain/yellow_routing.py demo/tests/test_yellow_routing.py demo/projects/tongfu.yaml
git commit -m "feat(asset-appraisal): enforce yellow source routes"
```

### Task 2: SDK-neutral OCR contracts and normalization

**Files:**
- Modify: `demo/schemas.py`
- Create: `demo/domain/ocr_normalization.py`
- Create: `demo/tests/test_ocr_normalization.py`

- [ ] **Step 1: Write failing normalization tests**

```python
from demo.domain.ocr_normalization import normalize_ocr_pages


def test_normalizes_text_and_table_cells_without_sdk_objects():
    pages = [{
        "page_number": 3,
        "page_count": 8,
        "blocks": [{"block_id": "p3-b1", "block_type": "text", "text": "资产总计", "confidence": 0.98, "bbox": [1, 2, 3, 4]}],
        "tables": [{"table_id": "p3-t1", "cells": [{"row": 1, "column": 2, "text": "1,234.50", "confidence": 0.97, "bbox": [5, 6, 7, 8]}]}],
    }]
    result = normalize_ocr_pages(pages)
    assert result["text_blocks"][0]["text"] == "资产总计"
    assert result["table_cells"][0]["text"] == "1,234.50"
    assert result["table_cells"][0]["evidence_id"] == "pdf:p3:t1:r1:c2"
```

- [ ] **Step 2: Verify RED**

Run: `UV_CACHE_DIR=/private/tmp/asset-appraisal-uv-cache uv run pytest demo/tests/test_ocr_normalization.py -q`

Expected: import fails for the missing module.

- [ ] **Step 3: Add Pydantic models and pure normalization**

Add `OcrBlock`, `OcrCell`, `OcrTable`, `OcrPage`, `OcrDocument`, `OcrPdfInput`, and `OcrPdfOutput` to `schemas.py`, each with Chinese descriptions, required flags, and examples. Implement `normalize_ocr_pages(pages: list[dict]) -> dict[str, list[dict]]` with no file or SDK access.

- [ ] **Step 4: Verify GREEN**

Run: `UV_CACHE_DIR=/private/tmp/asset-appraisal-uv-cache uv run pytest demo/tests/test_ocr_normalization.py demo/tests/test_contracts.py -q`

Expected: pass.

- [ ] **Step 5: Commit**

```bash
git add demo/schemas.py demo/domain/ocr_normalization.py demo/tests/test_ocr_normalization.py
git commit -m "feat(asset-appraisal): add portable OCR contracts"
```

### Task 3: PaddleOCR PP-StructureV3 adapter

**Files:**
- Create: `demo/adapters/paddle_ocr.py`
- Modify: `demo/adapters/ocr.py`
- Modify: `pyproject.toml`
- Create: `demo/tests/test_paddle_ocr_adapter.py`

- [ ] **Step 1: Write a failing adapter test using a fake PP-Structure result**

```python
from demo.adapters.paddle_ocr import PaddleStructureOcrAdapter


class FakeResult:
    json = {"res": {"page_index": 0, "page_count": 1, "overall_ocr_res": {"rec_texts": ["资产总计"], "rec_scores": [0.9], "rec_boxes": [[1, 2, 3, 4]]}, "table_res_list": []}}


class FakePipeline:
    def predict(self, input):
        return [FakeResult()]


def test_paddle_adapter_returns_serializable_pages(tmp_path):
    pdf = tmp_path / "scan.pdf"
    pdf.write_bytes(b"%PDF-1.4")
    pages, issues = PaddleStructureOcrAdapter(FakePipeline()).extract(pdf)
    assert not issues
    assert pages[0]["page_number"] == 1
    assert pages[0]["blocks"][0]["text"] == "资产总计"
```

- [ ] **Step 2: Verify RED**

Run: `UV_CACHE_DIR=/private/tmp/asset-appraisal-uv-cache uv run pytest demo/tests/test_paddle_ocr_adapter.py -q`

Expected: missing adapter module.

- [ ] **Step 3: Implement lazy PaddleOCR integration**

`PaddleStructureOcrAdapter` accepts an injected pipeline for tests. `create_local_pipeline()` lazily imports `PPStructureV3` and constructs it with table recognition enabled and formula recognition disabled. Convert Paddle result JSON into the portable page/block/cell dictionaries; do not leak NumPy or SDK result objects.

Pin the optional OCR environment to Python `<3.13` in documentation while keeping the package core on Python `>=3.11`. Keep Paddle imports out of `domain/`.

- [ ] **Step 4: Verify GREEN**

Run: `UV_CACHE_DIR=/private/tmp/asset-appraisal-uv-cache uv run pytest demo/tests/test_paddle_ocr_adapter.py demo/tests/test_adapters.py -q`

Expected: pass without installing Paddle because the test injects `FakePipeline`.

- [ ] **Step 5: Commit**

```bash
git add demo/adapters/paddle_ocr.py demo/adapters/ocr.py pyproject.toml demo/tests/test_paddle_ocr_adapter.py
git commit -m "feat(asset-appraisal): integrate PP-Structure OCR adapter"
```

### Task 4: OCR structured workbook adapter

**Files:**
- Create: `demo/adapters/ocr_workbook.py`
- Create: `demo/tests/test_ocr_workbook.py`

- [ ] **Step 1: Write a failing workbook round-trip test**

```python
from demo.adapters.ocr_workbook import export_ocr_workbook, read_ocr_workbook


def test_exports_four_auditable_sheets(tmp_path):
    normalized = {"text_blocks": [{"page_number": 1, "block_id": "b1", "block_type": "text", "text": "利润表", "confidence": 0.9, "bbox": [1, 2, 3, 4], "evidence_id": "pdf:p1:b1"}], "table_cells": [], "financial_data": [], "issues": []}
    path = export_ocr_workbook(tmp_path / "OCR结构化结果.xlsx", normalized)
    reloaded = read_ocr_workbook(path)
    assert set(reloaded) == {"OCR_文本", "OCR_表格", "标准财务数据", "识别问题"}
    assert reloaded["OCR_文本"][0]["证据编号"] == "pdf:p1:b1"
```

- [ ] **Step 2: Verify RED**

Run: `UV_CACHE_DIR=/private/tmp/asset-appraisal-uv-cache uv run pytest demo/tests/test_ocr_workbook.py -q`

Expected: missing adapter module.

- [ ] **Step 3: Implement the adapter**

Use the existing Python XLSX adapter dependency to write four sheets with typed numeric confidence/page/row/column values, frozen headers, filters, explicit column widths, wrapping, and no mutation of the input. The adapter accepts/returns plain dictionaries; c2m may replace it without changing `domain/`.

- [ ] **Step 4: Verify GREEN**

Run: `UV_CACHE_DIR=/private/tmp/asset-appraisal-uv-cache uv run pytest demo/tests/test_ocr_workbook.py -q`

Expected: pass.

- [ ] **Step 5: Commit**

```bash
git add demo/adapters/ocr_workbook.py demo/tests/test_ocr_workbook.py
git commit -m "feat(asset-appraisal): export OCR evidence workbook"
```

### Task 5: Layout-independent financial matching

**Files:**
- Create: `demo/domain/financial_matching.py`
- Create: `demo/domain/field_validation.py`
- Create: `demo/rules/financial_aliases.v1.yaml`
- Create: `demo/fixtures/ocr_cases.yaml`
- Create: `demo/expected/ocr_cases.yaml`
- Create: `demo/tests/test_financial_matching.py`

- [ ] **Step 1: Write failing alias/period/unit tests**

```python
from demo.domain.financial_matching import match_financial_table


def test_same_values_match_when_columns_and_units_change():
    cells = [
        {"row": 0, "column": 0, "text": "项目"},
        {"row": 0, "column": 1, "text": "2025年6月30日"},
        {"row": 0, "column": 2, "text": "2024年12月31日"},
        {"row": 1, "column": 0, "text": "资产合计"},
        {"row": 1, "column": 1, "text": "16371.913179"},
        {"row": 1, "column": 2, "text": "18082.424617"},
    ]
    result = match_financial_table(cells, aliases={"total_assets": ["总资产", "资产合计"]}, unit="万元")
    assert result["total_assets"]["2025-06-30"] == 163_719_131.79
    assert result["total_assets"]["2024-12-31"] == 180_824_246.17
```

Add nine more fixture cases covering row/column swap, `资产总计`, `负债合计`, period aliases, yuan/thousand-yuan units, split negative signs, OCR commas, missing required fields, and duplicate conflicting candidates.

- [ ] **Step 2: Verify RED**

Run: `UV_CACHE_DIR=/private/tmp/asset-appraisal-uv-cache uv run pytest demo/tests/test_financial_matching.py -q`

Expected: missing matcher.

- [ ] **Step 3: Implement pure matching and validation**

Implement `normalize_label`, `parse_period`, `parse_number`, `unit_multiplier`, `match_financial_table`, and `require_financial_fields`. All functions accept dictionaries/lists and return JSON-serializable dictionaries. Conflicting candidates remain conflicts; no first-match fallback.

- [ ] **Step 4: Verify GREEN for all ten cases**

Run: `UV_CACHE_DIR=/private/tmp/asset-appraisal-uv-cache uv run pytest demo/tests/test_financial_matching.py demo/tests/test_domain_examples.py -q`

Expected: pass.

- [ ] **Step 5: Commit**

```bash
git add demo/domain/financial_matching.py demo/domain/field_validation.py demo/rules/financial_aliases.v1.yaml demo/fixtures/ocr_cases.yaml demo/expected/ocr_cases.yaml demo/tests/test_financial_matching.py
git commit -m "feat(asset-appraisal): match financial data across layouts"
```

### Task 6: Strict Bailian GLM narrative adapter

**Files:**
- Create: `demo/adapters/bailian_glm.py`
- Create: `demo/prompts/yellow_narratives.v1.txt`
- Create: `demo/prompts/yellow_narratives_output.v1.json`
- Create: `demo/tests/test_bailian_glm.py`
- Modify: `.env.example`
- Modify: `pyproject.toml`

- [ ] **Step 1: Write failing whitelist and evidence tests**

```python
from demo.adapters.bailian_glm import BailianYellowNarrativeAdapter


class FakeClient:
    def post(self, *args, **kwargs):
        return FakeResponse({"choices": [{"message": {"content": '{"fields":{"company_profile_section":{"value":"示例概述","evidence_ids":["pdf:p1:b1"]},"book_net_assets":{"value":"999","evidence_ids":[]}}}'}}]})


class FakeResponse:
    def __init__(self, payload):
        self.payload = payload

    def raise_for_status(self):
        return None

    def json(self):
        return self.payload


def test_glm_rejects_fields_outside_seven_field_whitelist():
    values, issues = BailianYellowNarrativeAdapter(FakeClient(), "key").generate({"evidence": [{"evidence_id": "pdf:p1:b1", "text": "示例"}]})
    assert values == {"company_profile_section": "示例概述"}
    assert any("book_net_assets" in issue for issue in issues)
```

- [ ] **Step 2: Verify RED**

Run: `UV_CACHE_DIR=/private/tmp/asset-appraisal-uv-cache uv run pytest demo/tests/test_bailian_glm.py -q`

Expected: missing adapter.

- [ ] **Step 3: Implement Bailian HTTP adapter**

Read no environment variables inside the adapter. Constructor parameters are `client`, `api_key`, `base_url`, `model`, `prompt`, and `prompt_version`; the CLI composition root reads `DASHSCOPE_API_KEY`. Send `enable_thinking: false` and a JSON schema response format. Accept only the seven approved keys and only evidence IDs present in the request.

Do not put any real key in `.env.example`; add only `DASHSCOPE_API_KEY=`.

- [ ] **Step 4: Verify GREEN**

Run: `UV_CACHE_DIR=/private/tmp/asset-appraisal-uv-cache uv run pytest demo/tests/test_bailian_glm.py -q`

Expected: pass.

- [ ] **Step 5: Commit**

```bash
git add demo/adapters/bailian_glm.py demo/prompts/yellow_narratives.v1.txt demo/prompts/yellow_narratives_output.v1.json demo/tests/test_bailian_glm.py .env.example pyproject.toml
git commit -m "feat(asset-appraisal): add strict Bailian narratives"
```

### Task 7: End-to-end injected pipeline

**Files:**
- Create: `demo/pipeline.py`
- Modify: `demo/run.py`
- Modify: `demo/workflow.yaml`
- Create: `demo/tests/test_pipeline.py`

- [ ] **Step 1: Write a failing end-to-end test with injected adapters**

```python
class FixtureOcrAdapter:
    def __init__(self, fixture_path):
        self.fixture_path = Path(fixture_path)

    def extract(self, pdf_path):
        payload = json.loads(self.fixture_path.read_text(encoding="utf-8"))
        return payload["pages"], []


class FixtureLlmAdapter:
    def generate(self, evidence):
        return {"company_profile_section": "基于证据生成的示例概述"}, []


def test_pipeline_creates_ocr_xlsx_word_and_audit_without_cross_route_fallback(tmp_path):
    result = run_pipeline(
        project_config=Path("demo/projects/tongfu.yaml"),
        pdf_path=Path("资产评估工作流/通富2025.6.30合并及母公司审计报告.pdf"),
        output_dir=tmp_path,
        ocr_adapter=FixtureOcrAdapter("demo/fixtures/ocr_cases.yaml"),
        llm_adapter=FixtureLlmAdapter(),
        qichacha_adapter=None,
    )
    assert result.ocr_workbook_path.exists()
    assert result.report_path.exists()
    assert result.audit_path.exists()
    fields = json.loads((tmp_path / "normalized_fields.json").read_text())
    assert fields["company_profile_section"]
    assert fields["ownership_history"] == ""
```

- [ ] **Step 2: Verify RED**

Run: `UV_CACHE_DIR=/private/tmp/asset-appraisal-uv-cache uv run pytest demo/tests/test_pipeline.py -q`

Expected: `run_pipeline` is missing.

- [ ] **Step 3: Implement the composition root**

Implement sequential nodes matching `workflow.yaml`. Every adapter is passed as a function argument. The CLI accepts `--project`, `--pdf`, `--template`, `--output-dir`, `--ocr-engine paddle`, `--use-glm`, and node-input JSON. `run.py` remains backward compatible and delegates when `--pdf` is present.

Write manifest hashes for input PDF/template, route version, rule version, prompt version, and output paths. Generate Word only after route and required-financial validation.

- [ ] **Step 4: Verify GREEN and regressions**

Run: `UV_CACHE_DIR=/private/tmp/asset-appraisal-uv-cache uv run pytest demo/tests/test_pipeline.py demo/tests/test_demo_run.py demo/tests/test_financial_tables.py -q`

Expected: pass.

- [ ] **Step 5: Commit**

```bash
git add demo/pipeline.py demo/run.py demo/workflow.yaml demo/tests/test_pipeline.py
git commit -m "feat(asset-appraisal): add end-to-end OCR pipeline"
```

### Task 8: Delivery rules, docs, and actual-material verification

**Files:**
- Modify: `demo/README.md`
- Modify: `demo/data_manifest.yaml`
- Modify: `demo/CHANGELOG.md`
- Modify: `demo/tests/test_delivery_files.py`

- [ ] **Step 1: Add failing delivery assertions**

Assert the run outputs `OCR结构化结果.xlsx`, Word, audit XLSX, normalized JSON, issues, and manifest; asserts no yellow highlight, no `待人工补充`, immutable template hash, and route/prompt/data versions in the manifest.

- [ ] **Step 2: Verify RED**

Run: `UV_CACHE_DIR=/private/tmp/asset-appraisal-uv-cache uv run pytest demo/tests/test_delivery_files.py -q`

Expected: missing OCR workbook/manifest version assertions fail.

- [ ] **Step 3: Update documentation and manifests**

Document the Demo/c2m boundary, environment variables, Python 3.11 OCR environment, CLI example, all nodes, human checkpoint, failure policy, and the exact 7/5/6/2 yellow route. Update data versions and change log together.

- [ ] **Step 4: Run full automated verification**

Run: `UV_CACHE_DIR=/private/tmp/asset-appraisal-uv-cache uv run pytest -q`

Expected: all tests pass with no warnings.

- [ ] **Step 5: Run actual PaddleOCR environment and pipeline**

Run: `UV_CACHE_DIR=/private/tmp/asset-appraisal-uv-cache uv sync --python 3.11 --extra dev --extra services --extra ocr`

Run the CLI against the supplied scanned audit PDF and template. If model weights need downloading, permit Paddle's official model download. Keep the API key in `DASHSCOPE_API_KEY`; do not write it to disk or command logs.

- [ ] **Step 6: Inspect deliverables**

Use artifact-tool to inspect and render all four OCR workbook sheets, scan for formula errors, and verify evidence locators. Render the final Word to PNGs and inspect every page. Compare historical financial tables and key valuation amounts to the filled S2 reference.

- [ ] **Step 7: Commit**

```bash
git add demo/README.md demo/data_manifest.yaml demo/CHANGELOG.md demo/tests/test_delivery_files.py
git commit -m "docs(asset-appraisal): document OCR pipeline delivery"
```

### Task 9: Final branch verification

**Files:**
- Verify all changed files.

- [ ] **Step 1: Run the full test suite from a clean command**

Run: `UV_CACHE_DIR=/private/tmp/asset-appraisal-uv-cache uv run pytest -q`

Expected: zero failures.

- [ ] **Step 2: Validate source boundaries and secrets**

Run searches proving `demo/domain/` contains no imports of PaddleOCR, OpenAI/httpx, FastAPI, Streamlit, database modules, `.env`, or filesystem clients. Search tracked files for key-like strings and verify only the empty environment variable example exists.

- [ ] **Step 3: Inspect Git diff and status**

Run `git diff --check` and `git status --short`. Confirm only intended source, tests, docs, rules, prompts, fixtures, and expected files are tracked; outputs and model caches remain untracked.

- [ ] **Step 4: Finish the branch**

Use the finishing-development-branch workflow, preserve all commits, and report the branch name, commit IDs, tests, actual OCR limitations, and final deliverable paths.
