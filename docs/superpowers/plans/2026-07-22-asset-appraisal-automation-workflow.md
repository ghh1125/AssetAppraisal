# Asset Appraisal Automation Workflow Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a reusable, configuration-driven Python CLI that reads appraisal source files and service providers, replaces every placeholder and yellow instruction in a read-only Word template, and writes a new reviewable report plus audit artifacts.

**Architecture:** The workflow separates source providers, a normalized field registry, deterministic resolution rules, and an OOXML template adapter. Tasks communicate through typed JSON-serializable models and an isolated run directory; external provider failures degrade to explicit human-fill text instead of preventing Word generation.

**Tech Stack:** Python 3.11+, `uv`, `PyYAML`, `python-docx`, `lxml`, `openpyxl`, `pypdf`, `PyMuPDF`, `httpx`, optional `paddleocr`/`paddlepaddle`, `pytest`, LibreOffice rendering for Word QA.

---

## File Structure

Create the following focused units under `mcpify/AssetAppraisal/`:

- `pyproject.toml`: package metadata, base dependencies, OCR optional dependencies, and pytest configuration.
- `.env.example`: external service variable names without secrets.
- `README.md`: setup, configuration, command examples, outputs, and review responsibilities.
- `appraisal_workflow/__init__.py`: package version.
- `appraisal_workflow/__main__.py`: `python -m appraisal_workflow` entry point.
- `appraisal_workflow/cli.py`: argument parsing and command orchestration.
- `appraisal_workflow/config.py`: YAML loading, path resolution, and provider configuration.
- `appraisal_workflow/models.py`: typed fields, provenance, issues, mappings, and task results.
- `appraisal_workflow/artifacts.py`: isolated run directories, hashes, JSON output, and manifest data.
- `appraisal_workflow/pipeline.py`: task registration, dependency ordering, caching, and execution.
- `appraisal_workflow/registry.py`: candidate storage, source-priority resolution, and fallback values.
- `appraisal_workflow/calculations.py`: deterministic dates, units, increments, rates, and uppercase RMB.
- `appraisal_workflow/mapping.py`: mapping-file parsing and location coverage checks.
- `appraisal_workflow/providers/base.py`: OCR, LLM, and company-data provider protocols.
- `appraisal_workflow/providers/paddleocr_provider.py`: local PaddleOCR adapter.
- `appraisal_workflow/providers/openai_compatible.py`: OpenAI-compatible narrative adapter.
- `appraisal_workflow/providers/company_api.py`: configurable HTTP enterprise-data adapter.
- `appraisal_workflow/sources/excel.py`: workbook cell/formula/label extraction.
- `appraisal_workflow/sources/pdf.py`: PDF rendering, OCR caching, and page/table text catalog.
- `appraisal_workflow/templates/inventory.py`: stable Word location inventory.
- `appraisal_workflow/templates/filler.py`: run-preserving placeholder and yellow-block replacement.
- `appraisal_workflow/audit.py`: XLSX audit export and final run manifest.
- `appraisal_workflow/workflow.py`: concrete end-to-end task graph.
- `mappings/appraisal_report_v1.yaml`: 147 verified template locations and source rules.
- `projects/tongfu.yaml`: first project configuration.
- `projects/tongfu.manual.example.yaml`: non-secret human inputs.
- `tests/fixtures/`: small synthetic DOCX/PDF/XLSX/provider fixtures.
- `tests/`: unit, contract, template-regression, and end-to-end tests.

The work is organized into three increments. Phase I produces a working manual-data-to-Word pipeline. Phase II adds external and document providers. Phase III wires the real Tongfu project, audit files, CLI, documentation, and render verification.

---

## Phase I — Core Pipeline and Word Generation

### Task 1: Package, Configuration, and Core Models

**Files:**
- Create: `pyproject.toml`
- Create: `.env.example`
- Create: `appraisal_workflow/__init__.py`
- Create: `appraisal_workflow/models.py`
- Create: `appraisal_workflow/config.py`
- Test: `tests/test_config.py`
- Test: `tests/test_models.py`

- [ ] **Step 1: Create the package metadata**

Create `pyproject.toml`:

```toml
[build-system]
requires = ["setuptools>=75", "wheel"]
build-backend = "setuptools.build_meta"

[project]
name = "appraisal-workflow"
version = "0.1.0"
requires-python = ">=3.11"
dependencies = [
  "httpx>=0.27,<1",
  "lxml>=5,<7",
  "openpyxl>=3.1,<4",
  "Pillow>=10,<13",
  "pydantic>=2.8,<3",
  "pypdf>=5,<7",
  "PyMuPDF>=1.24,<2",
  "python-docx>=1.1,<2",
  "python-dotenv>=1,<2",
  "PyYAML>=6,<7",
]

[project.optional-dependencies]
ocr = ["paddleocr>=3,<4", "paddlepaddle>=3,<4"]
dev = ["pytest>=8,<10", "pytest-cov>=5,<8"]

[project.scripts]
appraisal-workflow = "appraisal_workflow.cli:main"

[tool.setuptools.packages.find]
include = ["appraisal_workflow*"]

[tool.pytest.ini_options]
testpaths = ["tests"]
addopts = "-q"
```

- [ ] **Step 2: Write failing model and configuration tests**

Create `tests/test_models.py` and `tests/test_config.py`:

```python
# tests/test_models.py
from appraisal_workflow.models import FieldCandidate


def test_field_candidate_serializes_provenance():
    item = FieldCandidate(
        field_key="assessed_entity.legal_name",
        value="通富热处理（昆山）有限公司",
        source_kind="income_workbook",
        source_file="收益法.xlsx",
        source_locator="项目信息!B5",
    )
    assert item.model_dump()["source_locator"] == "项目信息!B5"
```

```python
# tests/test_config.py
from pathlib import Path
from appraisal_workflow.config import load_project_config


def test_config_paths_are_resolved_relative_to_project_file(tmp_path: Path):
    config = tmp_path / "project.yaml"
    config.write_text(
        "project:\n  id: demo\n  template: input.docx\n  output_root: runs\n"
        "sources:\n  manual_inputs: manual.yaml\n"
        "template_mapping: mapping.yaml\n",
        encoding="utf-8",
    )
    loaded = load_project_config(config)
    assert loaded.project.template == tmp_path / "input.docx"
    assert loaded.sources.manual_inputs == tmp_path / "manual.yaml"
```

- [ ] **Step 3: Run the tests and verify RED**

Run:

```bash
cd /Users/ghh/Documents/Code/mcpify/AssetAppraisal
uv run --extra dev pytest tests/test_models.py tests/test_config.py -v
```

Expected: collection fails with `ModuleNotFoundError: No module named 'appraisal_workflow'`.

- [ ] **Step 4: Implement the models and configuration loader**

Create `appraisal_workflow/models.py` with Pydantic models for `FieldCandidate`, `ResolvedField`, `Issue`, `LocationMapping`, `TaskResult`, and `RunManifest`. Create `appraisal_workflow/config.py` with `ProjectConfig`, `ProjectSection`, `SourceSection`, `ProviderSection`, and:

```python
def load_project_config(path: Path) -> ProjectConfig:
    raw = yaml.safe_load(path.read_text(encoding="utf-8"))
    config = ProjectConfig.model_validate(raw)
    return config.resolve_paths(path.parent)
```

`resolve_paths()` must resolve template, output root, every non-null source file, and template mapping against the configuration file directory without resolving environment variable names.

- [ ] **Step 5: Run the tests and verify GREEN**

Run: `uv run --extra dev pytest tests/test_models.py tests/test_config.py -v`

Expected: 2 tests pass.

- [ ] **Step 6: Commit**

```bash
git add pyproject.toml .env.example appraisal_workflow tests/test_models.py tests/test_config.py
git commit -m "feat: add appraisal workflow configuration and models"
```

### Task 2: Isolated Run Artifacts and Task Pipeline

**Files:**
- Create: `appraisal_workflow/artifacts.py`
- Create: `appraisal_workflow/pipeline.py`
- Test: `tests/test_artifacts.py`
- Test: `tests/test_pipeline.py`

- [ ] **Step 1: Write failing artifact and task-order tests**

```python
# tests/test_artifacts.py
from pathlib import Path
from appraisal_workflow.artifacts import ArtifactStore


def test_artifact_store_never_reuses_template_path(tmp_path: Path):
    template = tmp_path / "template.docx"
    template.write_bytes(b"template")
    store = ArtifactStore.create(tmp_path / "runs", "demo", "run-1", template)
    assert store.report_path != template
    assert store.run_dir == tmp_path / "runs" / "demo" / "run-1"
```

```python
# tests/test_pipeline.py
from appraisal_workflow.pipeline import Pipeline, FunctionTask


def test_pipeline_runs_dependencies_before_consumers(tmp_path):
    events = []
    pipeline = Pipeline([
        FunctionTask("resolve", ("extract",), lambda ctx: events.append("resolve")),
        FunctionTask("extract", (), lambda ctx: events.append("extract")),
    ])
    pipeline.run(object())
    assert events == ["extract", "resolve"]
```

- [ ] **Step 2: Run and verify RED**

Run: `uv run --extra dev pytest tests/test_artifacts.py tests/test_pipeline.py -v`

Expected: imports fail because the modules do not exist.

- [ ] **Step 3: Implement ArtifactStore and Pipeline**

`ArtifactStore.create()` must create `runs/<project>/<run-id>/`, `logs/`, `ocr/`, and `cache/`; set paths for the report, audit workbook, normalized fields, issues, and manifest; and reject a run directory equal to or nested inside the template path. `Pipeline` must topologically sort named tasks, reject duplicate names and cycles, support `from_task` and `only_task`, and collect `TaskResult` records.

- [ ] **Step 4: Run and verify GREEN**

Run: `uv run --extra dev pytest tests/test_artifacts.py tests/test_pipeline.py -v`

Expected: all tests pass.

- [ ] **Step 5: Commit**

```bash
git add appraisal_workflow/artifacts.py appraisal_workflow/pipeline.py tests/test_artifacts.py tests/test_pipeline.py
git commit -m "feat: add isolated artifact store and task pipeline"
```

### Task 3: Word Inventory with Stable Location IDs

**Files:**
- Create: `appraisal_workflow/templates/__init__.py`
- Create: `appraisal_workflow/templates/inventory.py`
- Test: `tests/test_template_inventory.py`

- [ ] **Step 1: Write the failing template regression test**

```python
from pathlib import Path
from appraisal_workflow.templates.inventory import inventory_template


TEMPLATE = Path("资产评估工作流/评估报告版式-沟通标注版.docx")


def test_retained_template_has_expected_stable_locations():
    inventory = inventory_template(TEMPLATE)
    assert len(inventory.placeholders) == 127
    assert len(inventory.highlight_blocks) == 20
    assert sum(x.part.startswith("word/footer") for x in inventory.placeholders) == 2
    assert len({x.location_id for x in inventory.all_locations}) == 147
```

- [ ] **Step 2: Run and verify RED**

Run: `uv run --extra dev pytest tests/test_template_inventory.py -v`

Expected: import failure for `appraisal_workflow.templates.inventory`.

- [ ] **Step 3: Implement deterministic OOXML inventory**

Read `word/document.xml`, matching headers, footers, footnotes, and endnotes from the DOCX ZIP. For each `w:p`, concatenate `w:t`, `w:tab`, and break nodes before applying `X{2,}`. Generate IDs as `<PART>-P####-X##`; create `-H01` only when a paragraph contains highlighted instruction text and no placeholder. Record part, paragraph index, occurrence, marker, context, table membership, and highlighted text.

- [ ] **Step 4: Run and verify GREEN**

Run: `uv run --extra dev pytest tests/test_template_inventory.py -v`

Expected: the 147-location regression test passes.

- [ ] **Step 5: Commit**

```bash
git add appraisal_workflow/templates tests/test_template_inventory.py
git commit -m "feat: inventory stable appraisal template locations"
```

### Task 4: Field Registry, Fixed Priority, and Calculations

**Files:**
- Create: `appraisal_workflow/registry.py`
- Create: `appraisal_workflow/calculations.py`
- Test: `tests/test_registry.py`
- Test: `tests/test_calculations.py`

- [ ] **Step 1: Write failing resolution and calculation tests**

```python
from appraisal_workflow.models import FieldCandidate
from appraisal_workflow.registry import FieldRegistry


def test_registry_uses_first_nonempty_candidate_by_configured_priority():
    registry = FieldRegistry()
    registry.add(FieldCandidate(field_key="net_assets", value=100, source_kind="pdf", source_file="a.pdf", source_locator="p1"))
    registry.add(FieldCandidate(field_key="net_assets", value=110, source_kind="reporting_workbook", source_file="a.xlsx", source_locator="表1!B2"))
    result = registry.resolve("net_assets", ["reporting_workbook", "pdf"])
    assert result.value == 110
    assert len(result.candidates) == 2
```

```python
from decimal import Decimal
from appraisal_workflow.calculations import increment, increment_rate


def test_increment_and_rate_use_decimal():
    assert increment(Decimal("120"), Decimal("100")) == Decimal("20")
    assert increment_rate(Decimal("120"), Decimal("100")) == Decimal("0.2")
    assert increment_rate(Decimal("120"), Decimal("0")) is None
```

- [ ] **Step 2: Run and verify RED**

Run: `uv run --extra dev pytest tests/test_registry.py tests/test_calculations.py -v`

Expected: imports fail.

- [ ] **Step 3: Implement registry and deterministic calculations**

`FieldRegistry` must retain all candidates, sort only by the supplied source priority, return the first non-null/non-empty candidate, and return `【待人工补充：<field name>】` when no candidate exists. Implement decimal-safe amount conversion, increment, increment rate, report-validity date, date splitting, percentage formatting, and uppercase RMB conversion in `calculations.py`.

- [ ] **Step 4: Run and verify GREEN**

Run: `uv run --extra dev pytest tests/test_registry.py tests/test_calculations.py -v`

Expected: all tests pass.

- [ ] **Step 5: Commit**

```bash
git add appraisal_workflow/registry.py appraisal_workflow/calculations.py tests/test_registry.py tests/test_calculations.py
git commit -m "feat: add deterministic field resolution and calculations"
```

### Task 5: Mapping Loader and Exact Word Replacement

**Files:**
- Create: `appraisal_workflow/mapping.py`
- Create: `appraisal_workflow/templates/filler.py`
- Create: `scripts/convert_verified_mapping.py`
- Create: `mappings/appraisal_report_v1.yaml`
- Test: `tests/test_mapping.py`
- Test: `tests/test_word_filler.py`

- [ ] **Step 1: Write failing mapping coverage test**

```python
from pathlib import Path
from appraisal_workflow.mapping import load_mapping


def test_mapping_covers_every_verified_template_location():
    mapping = load_mapping(Path("mappings/appraisal_report_v1.yaml"))
    assert len(mapping.locations) == 147
    assert sum(x.record_type == "占位符" for x in mapping.locations) == 127
    assert sum(x.record_type == "黄色标注内容块" for x in mapping.locations) == 20
    assert len({x.location_id for x in mapping.locations}) == 147
```

- [ ] **Step 2: Write failing filler behavior test**

Create a synthetic DOCX fixture with two placeholders in one paragraph and one highlighted instruction paragraph. The test must call:

```python
result = fill_template(
    template=fixture,
    output=tmp_path / "filled.docx",
    replacements={
        "DOCUMENT-P0001-X01": "甲公司",
        "DOCUMENT-P0001-X02": "乙公司",
        "DOCUMENT-P0002-H01": "实际生成的业务概述",
    },
)
```

Assert that the template hash is unchanged, output is a different path, `XXX` is absent, original yellow instruction text is absent, and the replacement paragraph has no `w:highlight` property.

- [ ] **Step 3: Run and verify RED**

Run: `uv run --extra dev pytest tests/test_mapping.py tests/test_word_filler.py -v`

Expected: imports or missing mapping file cause failure.

- [ ] **Step 4: Convert the verified 147-record mapping**

Implement `scripts/convert_verified_mapping.py` to read:

```text
/private/tmp/asset-appraisal-mapping/Code/mcpify/AssetAppraisal/outputs/资产评估Word填充映射.json
```

and write `mappings/appraisal_report_v1.yaml` with `template_name`, `version`, and all record fields needed for resolution and replacement. Fail if the input does not contain exactly 147 unique location IDs. Run the converter once and commit the generated YAML; runtime code must not depend on `/private/tmp`.

- [ ] **Step 5: Implement mapping loader and run-preserving filler**

`load_mapping()` validates uniqueness and required fields. `fill_template()` copies the DOCX ZIP to a new file, patches only affected XML parts, replaces the Nth placeholder within the paragraph text while preserving surrounding runs, replaces highlight-only paragraph content with the resolved text, removes highlight formatting from replacement runs, and supports `inline`, `paragraph`, and `table` replacement types. For a table replacement, insert a Word table after the mapped paragraph using explicit column widths, then remove the instruction paragraph.

- [ ] **Step 6: Run and verify GREEN**

Run: `uv run --extra dev pytest tests/test_mapping.py tests/test_word_filler.py -v`

Expected: mapping coverage and filler tests pass.

- [ ] **Step 7: Commit**

```bash
git add appraisal_workflow/mapping.py appraisal_workflow/templates/filler.py scripts/convert_verified_mapping.py mappings/appraisal_report_v1.yaml tests/test_mapping.py tests/test_word_filler.py tests/fixtures
git commit -m "feat: add verified mappings and exact Word replacement"
```

At this point a manual field dictionary can generate a separate Word without external providers.

---

## Phase II — Source and Service Providers

### Task 6: Excel Source Adapter

**Files:**
- Create: `appraisal_workflow/sources/__init__.py`
- Create: `appraisal_workflow/sources/excel.py`
- Test: `tests/test_excel_source.py`

- [ ] **Step 1: Write failing workbook tests**

Create a small workbook fixture with a formula cell, cached-value companion fixture, merged labels, and a named field. Test that `WorkbookCatalog.get("项目信息!B5")` returns its value and `find_label("净资产")` returns sheet, coordinate, formula, value, and surrounding cells.

- [ ] **Step 2: Run and verify RED**

Run: `uv run --extra dev pytest tests/test_excel_source.py -v`

Expected: module import fails.

- [ ] **Step 3: Implement workbook cataloging**

Load each workbook twice with `data_only=False` and `data_only=True`. Record sheet visibility, merged ranges, coordinates, labels, formulas, cached values, data types, and number formats. Implement exact A1 locator access and normalized label search. The adapter must never save the source workbook.

- [ ] **Step 4: Run and verify GREEN**

Run: `uv run --extra dev pytest tests/test_excel_source.py -v`

Expected: all tests pass.

- [ ] **Step 5: Commit**

```bash
git add appraisal_workflow/sources tests/test_excel_source.py tests/fixtures
git commit -m "feat: add Excel source adapter"
```

### Task 7: PDF Rendering, OCR Provider, and Cache

**Files:**
- Create: `appraisal_workflow/providers/__init__.py`
- Create: `appraisal_workflow/providers/base.py`
- Create: `appraisal_workflow/providers/paddleocr_provider.py`
- Create: `appraisal_workflow/sources/pdf.py`
- Test: `tests/test_ocr_provider_contract.py`
- Test: `tests/test_pdf_source.py`

- [ ] **Step 1: Write failing provider-contract tests**

Define a fake OCR provider returning one page with text and a table. Test that `extract_pdf()` writes `ocr/pages/0001.json`, preserves `page_number=1`, and reuses the cached result on the second call without invoking the provider again.

- [ ] **Step 2: Run and verify RED**

Run: `uv run --extra dev pytest tests/test_ocr_provider_contract.py tests/test_pdf_source.py -v`

Expected: modules do not exist.

- [ ] **Step 3: Implement provider protocols and PDF source**

Create `OcrProvider`, `LlmProvider`, and `CompanyDataProvider` protocols. Use PyMuPDF to render each PDF page to PNG. Cache OCR JSON by SHA-256 of PDF bytes, page number, render DPI, provider name, and provider settings. Preserve raw OCR text, tables, and confidence values without applying confidence gates.

- [ ] **Step 4: Implement PaddleOCR adapter**

Import PaddleOCR lazily. If the optional dependency is missing, raise a typed `ProviderUnavailable` that the workflow converts to missing fields. Normalize PaddleOCR 3.x output into the provider protocol without discarding raw payloads.

- [ ] **Step 5: Run and verify GREEN**

Run: `uv run --extra dev pytest tests/test_ocr_provider_contract.py tests/test_pdf_source.py -v`

Expected: contract and cache tests pass without requiring PaddleOCR because they use the fake provider.

- [ ] **Step 6: Commit**

```bash
git add appraisal_workflow/providers appraisal_workflow/sources/pdf.py tests/test_ocr_provider_contract.py tests/test_pdf_source.py tests/fixtures
git commit -m "feat: add cached PDF OCR provider interface"
```

### Task 8: Configurable Company Data API

**Files:**
- Create: `appraisal_workflow/providers/company_api.py`
- Test: `tests/test_company_api.py`

- [ ] **Step 1: Write failing request and fallback tests**

Use `httpx.MockTransport` to assert that the configured URL, authentication header, query parameter, timeout, and JSON field mapping produce normalized company fields. Add a second test showing that a missing API key returns a provider-unavailable result instead of raising out of the workflow.

- [ ] **Step 2: Run and verify RED**

Run: `uv run --extra dev pytest tests/test_company_api.py -v`

Expected: import fails.

- [ ] **Step 3: Implement GenericCompanyApiProvider**

Read the API key only from the configured environment variable. Support configurable HTTP method, base URL, endpoint, auth header, query-parameter mapping, response root path, and response-field mapping. Redact auth headers and key-like query values from logged request metadata.

- [ ] **Step 4: Run and verify GREEN**

Run: `uv run --extra dev pytest tests/test_company_api.py -v`

Expected: request mapping and missing-key tests pass.

- [ ] **Step 5: Commit**

```bash
git add appraisal_workflow/providers/company_api.py tests/test_company_api.py
git commit -m "feat: add configurable company data provider"
```

### Task 9: OpenAI-Compatible Narrative Provider

**Files:**
- Create: `appraisal_workflow/providers/openai_compatible.py`
- Create: `appraisal_workflow/prompts.py`
- Test: `tests/test_llm_provider.py`

- [ ] **Step 1: Write failing structured-output tests**

Use `httpx.MockTransport` to return a valid JSON object inside `choices[0].message.content`. Assert that the provider returns only requested narrative keys. Add invalid JSON and missing-key tests that return typed provider failures for workflow degradation.

- [ ] **Step 2: Run and verify RED**

Run: `uv run --extra dev pytest tests/test_llm_provider.py -v`

Expected: import fails.

- [ ] **Step 3: Implement provider and prompts**

Call `/chat/completions` using configured base URL, model, API key, timeout, and temperature. Require JSON output with requested keys. Prompts must state that financial numbers may only be copied from supplied evidence and must not be invented. Redact the API key from errors and logs.

- [ ] **Step 4: Run and verify GREEN**

Run: `uv run --extra dev pytest tests/test_llm_provider.py -v`

Expected: valid, invalid, and missing-credential tests pass.

- [ ] **Step 5: Commit**

```bash
git add appraisal_workflow/providers/openai_compatible.py appraisal_workflow/prompts.py tests/test_llm_provider.py
git commit -m "feat: add OpenAI-compatible narrative provider"
```

---

## Phase III — End-to-End Workflow, Audit, and Delivery

### Task 10: Concrete Workflow Tasks and Field Resolution

**Files:**
- Create: `appraisal_workflow/workflow.py`
- Test: `tests/test_workflow_degradation.py`

- [ ] **Step 1: Write a failing degraded-run test**

Use a synthetic template, manual inputs, fake Excel source, and unavailable OCR/API/LLM providers. Assert that `build_workflow(config).run()` still produces replacements for every mapped location, with unavailable fields equal to `【待人工补充：字段名称】` and issues retained for audit.

- [ ] **Step 2: Run and verify RED**

Run: `uv run --extra dev pytest tests/test_workflow_degradation.py -v`

Expected: `build_workflow` is missing.

- [ ] **Step 3: Implement the task graph**

Register tasks in this dependency order:

```text
inventory
├── load_manual
├── extract_excel
├── extract_pdf
├── fetch_company
└── generate_narrative (depends on all source tasks)
    ↓
resolve_fields
    ↓
fill_word
    ↓
export_audit
```

Each source task writes candidates to `FieldRegistry`. Provider exceptions become issues and missing candidates. `resolve_fields` applies mapping priorities and calculations. `fill_word` supplies a replacement for all 147 mapping locations and always writes a new output path.

- [ ] **Step 4: Run and verify GREEN**

Run: `uv run --extra dev pytest tests/test_workflow_degradation.py -v`

Expected: degraded workflow test passes.

- [ ] **Step 5: Commit**

```bash
git add appraisal_workflow/workflow.py tests/test_workflow_degradation.py tests/fixtures
git commit -m "feat: orchestrate degraded end-to-end appraisal workflow"
```

### Task 11: Audit Workbook and Run Manifest

**Files:**
- Create: `appraisal_workflow/audit.py`
- Test: `tests/test_audit.py`

- [ ] **Step 1: Write failing audit tests**

Build two resolved fields with multiple candidates and one missing field. Export the audit workbook, reload it with openpyxl, and assert sheets `填充结果`, `候选来源`, `问题清单`, and `运行信息` exist with exact headers. Assert the run manifest includes template hash, input hashes, task results, provider types, and output paths but contains no API-key values.

- [ ] **Step 2: Run and verify RED**

Run: `uv run --extra dev pytest tests/test_audit.py -v`

Expected: import fails.

- [ ] **Step 3: Implement audit exports**

Create a readable XLSX with frozen headers, filters, wrapped text, widths capped at 60 characters, and one row per Word location. Candidate-source rows must include source file and locator. Write JSON using UTF-8 with stable indentation. Recursively redact keys matching `api_key`, `authorization`, `token`, or `secret` before serialization.

- [ ] **Step 4: Run and verify GREEN**

Run: `uv run --extra dev pytest tests/test_audit.py -v`

Expected: workbook and manifest tests pass.

- [ ] **Step 5: Commit**

```bash
git add appraisal_workflow/audit.py tests/test_audit.py
git commit -m "feat: export field audit workbook and run manifest"
```

### Task 12: CLI, Tongfu Project, and Template-Safe End-to-End Test

**Files:**
- Create: `appraisal_workflow/cli.py`
- Create: `appraisal_workflow/__main__.py`
- Create: `projects/tongfu.yaml`
- Create: `projects/tongfu.manual.example.yaml`
- Create: `.env.example`
- Test: `tests/test_tongfu_e2e.py`

- [ ] **Step 1: Write the failing end-to-end test**

The test reads `projects/tongfu.yaml`, rewrites every retained input path to an absolute path, writes that resolved configuration into a temporary directory with a temporary `output_root`, uses fake network providers, runs the CLI, and asserts:

```python
assert template_sha256_after == template_sha256_before
assert report_path.exists()
assert audit_path.exists()
assert manifest_path.exists()
assert count_placeholders(report_path) == 0
assert count_original_highlight_instructions(report_path) == 0
assert report_path.resolve() != template_path.resolve()
```

- [ ] **Step 2: Run and verify RED**

Run: `uv run --extra dev pytest tests/test_tongfu_e2e.py -v`

Expected: CLI or project configuration is missing.

- [ ] **Step 3: Implement CLI and project configuration**

Support:

```bash
python -m appraisal_workflow run projects/tongfu.yaml
python -m appraisal_workflow run projects/tongfu.yaml --from-task extract_pdf
python -m appraisal_workflow run projects/tongfu.yaml --only-task inventory
```

Return `0` when Word generation succeeds even if human-fill issues exist; return nonzero for configuration, template, or output-write failures. Print only run ID, task summary, report path, audit path, and issue count.

Populate `projects/tongfu.yaml` with the retained template and three retained source files. Populate the manual example with named but empty user-editable fields. Configure provider environment-variable names without secrets.

- [ ] **Step 4: Run and verify GREEN**

Run: `uv run --extra dev pytest tests/test_tongfu_e2e.py -v`

Expected: end-to-end test passes and the original template hash remains unchanged.

- [ ] **Step 5: Commit**

```bash
git add appraisal_workflow/cli.py appraisal_workflow/__main__.py projects .env.example tests/test_tongfu_e2e.py
git commit -m "feat: add Tongfu CLI workflow"
```

### Task 13: Documentation and Full Verification

**Files:**
- Create: `README.md`
- Modify: `projects/tongfu.manual.example.yaml`
- Test: all tests and generated artifacts

- [ ] **Step 1: Write README usage and review instructions**

Document:

1. `uv sync --extra dev --extra ocr` installation.
2. Copy `.env.example` to `.env` and fill provider variables.
3. Copy `tongfu.manual.example.yaml` to `tongfu.manual.yaml` and fill manual fields.
4. Run `uv run python -m appraisal_workflow run projects/tongfu.yaml`.
5. Locate the new Word and audit workbook under the printed run directory.
6. Confirm that the original Word remains the template.
7. Explain that OCR, API, LLM, financial values, and conclusions require manual professional review.

- [ ] **Step 2: Run the complete automated suite**

Run:

```bash
uv run --extra dev pytest -v
```

Expected: all tests pass with no warnings from project code.

- [ ] **Step 3: Run the real project in degraded mode**

Run without provider keys first:

```bash
uv run python -m appraisal_workflow run projects/tongfu.yaml
```

Expected: command exits `0`, creates a unique run directory, produces a new Word and audit workbook, replaces all 147 locations, and uses explicit human-fill text for unavailable external fields.

- [ ] **Step 4: Run OCR-enabled project**

After `uv sync --extra ocr`, run:

```bash
uv run --extra ocr python -m appraisal_workflow run projects/tongfu.yaml --from-task extract_pdf
```

Expected: 48 PDF page OCR JSON files exist under the run cache/output, and a new report is generated.

- [ ] **Step 5: Render and visually inspect the generated Word**

Run:

```bash
env TMPDIR=/private/tmp \
/Users/ghh/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 \
/Users/ghh/.codex/plugins/cache/openai-primary-runtime/documents/26.715.12143/skills/documents/render_docx.py \
runs/<project>/<run-id>/资产评估报告_待复核.docx \
--output_dir /private/tmp/appraisal-workflow-render --emit_pdf
```

Inspect every rendered page for missing Chinese glyphs, clipped tables, overlapping text, broken page headers/footers, and unremoved yellow instruction text. If LibreOffice font substitution prevents reliable Chinese inspection, open the generated Word in WPS and inspect the first, middle, and last pages.

- [ ] **Step 6: Verify the template was not changed**

Run:

```bash
shasum -a 256 资产评估工作流/评估报告版式-沟通标注版.docx
git status --short -- 资产评估工作流/评估报告版式-沟通标注版.docx
```

Expected: hash matches the manifest input hash and Git reports no modification to the template.

- [ ] **Step 7: Commit documentation**

```bash
git add README.md projects/tongfu.manual.example.yaml
git commit -m "docs: document appraisal workflow setup and review"
```

---

## Final Acceptance Checklist

- [ ] `python -m appraisal_workflow run projects/tongfu.yaml` is the primary interface.
- [ ] The code does not import or call Codex.
- [ ] The original Word template hash is unchanged.
- [ ] Every run writes a separately named Word file.
- [ ] All 127 placeholders and 20 yellow instructions are replaced.
- [ ] Missing OCR/API/LLM values become explicit human-fill text and do not stop generation.
- [ ] Excel, OCR, company API, and LLM adapters are replaceable through configuration.
- [ ] Audit XLSX and JSON manifest trace every filled location to its source or fallback.
- [ ] Secrets are absent from YAML, Word, audit, JSON, and logs.
- [ ] Automated tests and Word render/WPS checks pass.
