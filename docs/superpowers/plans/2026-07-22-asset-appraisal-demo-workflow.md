# Asset Appraisal Demo Workflow Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build an executable, c2m-ready business Demo that fills a new appraisal Word from configured sources while keeping reusable business rules, schemas, prompts, examples, and acceptance tests independent of production infrastructure.

**Architecture:** `demo/domain/` contains pure functions that accept Pydantic/plain data and return JSON-serializable results. `demo/adapters/` owns files, Excel, OCR, HTTP, environment variables, and Word I/O; `demo/run.py` composes them according to `workflow.yaml`. The Demo always creates a new Word, replaces all 127 placeholders and 20 yellow instructions, and leaves correctness review to the user.

**Tech Stack:** Python 3.11+, Pydantic 2, PyYAML, python-docx/lxml, openpyxl, PyMuPDF, httpx, optional PaddleOCR, pytest.

---

### Task 1: Demo Contract and Workflow Skeleton

**Files:**
- Create: `pyproject.toml`
- Create: `demo/__init__.py`
- Create: `demo/schemas.py`
- Create: `demo/workflow.yaml`
- Create: `demo/tests/test_contracts.py`

- [ ] Write `test_contracts.py` first. Assert every workflow node names an existing input/output Pydantic model, every model field has a Chinese description and example, and the required node order is `inventory → extract_sources → resolve_fields → generate_narrative → fill_word → export_audit`.
- [ ] Run `uv run --extra dev pytest demo/tests/test_contracts.py -v`; verify failure because files are missing.
- [ ] Implement the minimal package and schemas: `SourceEvidence`, `FieldCandidate`, `ResolvedField`, `WordLocation`, `LocationMapping`, `WorkflowInput`, `WorkflowOutput`, and per-node input/output models. All models use `ConfigDict(extra="forbid")`.
- [ ] Create `workflow.yaml` with node name, input model, output model, dependencies, and `human_checkpoint` for generated report review.
- [ ] Run the test and verify it passes.
- [ ] Commit with `feat: add appraisal demo contracts`.

### Task 2: Pure Field Rules and Ten Representative Examples

**Files:**
- Create: `demo/domain/__init__.py`
- Create: `demo/domain/registry.py`
- Create: `demo/domain/calculations.py`
- Create: `demo/fixtures/cases.yaml`
- Create: `demo/expected/cases.yaml`
- Create: `demo/tests/test_domain_examples.py`

- [ ] Write the failing parameterized test that loads exactly 10 fixture cases and compares domain outputs with `expected/cases.yaml`.
- [ ] Include normal and boundary cases for source priority, empty candidates, date splitting, amount unit conversion, increment, zero-book-value rate, report validity, company-name normalization, selected methods, and human-fill fallback.
- [ ] Run the test and verify imports fail.
- [ ] Implement pure functions only: `resolve_candidate`, `resolve_all`, `split_date`, `convert_amount`, `increment`, `increment_rate`, `validity_period`, `normalize_company_name`, `format_methods`, and `human_fill`.
- [ ] Verify `domain/` contains no imports of `os`, `pathlib`, `httpx`, `openpyxl`, `docx`, `streamlit`, `fastapi`, or c2m modules.
- [ ] Run the parameterized test and verify 10 cases pass.
- [ ] Commit with `feat: add reusable appraisal domain rules`.

### Task 3: Mapping, Template Inventory, and Pure Replacement Planning

**Files:**
- Create: `demo/domain/mapping.py`
- Create: `demo/domain/replacement.py`
- Create: `demo/adapters/word.py`
- Create: `demo/mappings/appraisal_report_v1.yaml`
- Create: `demo/tests/test_template_mapping.py`
- Create: `demo/tests/test_word_replacement.py`

- [ ] Write the mapping regression test first: the retained template must inventory 127 placeholder locations, 20 highlight-only locations, 2 footer placeholders, and 147 unique stable IDs; the mapping file must cover the same IDs.
- [ ] Write the Word replacement test first using a synthetic fixture: two placeholders in one paragraph and one highlighted instruction must become three supplied values, the template hash must remain unchanged, and the output path must differ.
- [ ] Run both tests and verify failure.
- [ ] Implement pure `build_replacement_plan(mappings, resolved_fields)` returning JSON-serializable replacement records with human-fill fallback.
- [ ] Implement the Word adapter to inventory OOXML parts and apply a replacement plan to a copied DOCX. Replace the Nth `X{2,}` occurrence, replace highlight-only paragraph content, and remove highlight formatting. Support text and table payloads.
- [ ] Generate and commit `appraisal_report_v1.yaml` from the previously verified 147-record mapping artifact; runtime must not reference temporary paths.
- [ ] Run both tests and verify they pass.
- [ ] Commit with `feat: add template mapping and Word replacement`.

### Task 4: External Adapters with Dependency Injection

**Files:**
- Create: `demo/adapters/__init__.py`
- Create: `demo/adapters/excel.py`
- Create: `demo/adapters/ocr.py`
- Create: `demo/adapters/company_api.py`
- Create: `demo/adapters/llm.py`
- Create: `demo/prompts/company_narrative.v1.txt`
- Create: `demo/prompts/narrative_output.v1.json`
- Create: `demo/tests/test_adapters.py`

- [ ] Write failing adapter contract tests with temporary Excel/PDF fixtures and injected fake HTTP/LLM transports.
- [ ] Assert adapters return Pydantic/plain data, never write Word, and missing optional dependencies or credentials return empty results plus issue text instead of raising into the workflow.
- [ ] Run tests and verify failure.
- [ ] Implement Excel exact-cell and label readers without saving source workbooks.
- [ ] Implement OCR interface plus PaddleOCR adapter loaded lazily; accept OCR output as-is without confidence gating.
- [ ] Implement configurable company-data HTTP adapter and OpenAI-compatible LLM adapter. Clients, keys, URLs, and transports are passed into constructors; adapters do not create global clients.
- [ ] Store the narrative Prompt and JSON output schema as versioned files.
- [ ] Run tests and verify they pass.
- [ ] Commit with `feat: add injectable appraisal source adapters`.

### Task 5: Orchestration, Audit, and CLI

**Files:**
- Create: `demo/run.py`
- Create: `demo/adapters/audit.py`
- Create: `demo/tests/test_demo_run.py`

- [ ] Write the failing offline end-to-end test first using a synthetic template and sources. It must produce a new Word, `字段审计清单.xlsx`, `run_manifest.json`, and `issues.json`, leaving the source template hash unchanged.
- [ ] Assert success even when OCR/API/LLM are unavailable; all unfilled positions must contain `【待人工补充：字段名】` rather than `XX`/`XXX` or original yellow instructions.
- [ ] Run the test and verify failure.
- [ ] Implement `python -m demo.run <project.yaml> [--output-dir PATH] [--offline]`. The orchestration layer loads files and environment variables, injects adapters, calls pure domain functions, writes the new Word, and exports the audit workbook/JSON.
- [ ] Keep state local to one run; do not implement user, permission, database, queue, retry, monitoring, or c2m behavior.
- [ ] Run the test and verify it passes.
- [ ] Commit with `feat: add executable appraisal demo workflow`.

### Task 6: Tongfu Configuration, Data Manifest, and Rule Change Record

**Files:**
- Create: `demo/projects/tongfu.yaml`
- Create: `demo/projects/tongfu.manual.example.yaml`
- Create: `demo/data_manifest.yaml`
- Create: `demo/CHANGELOG.md`
- Create: `demo/tests/test_tongfu_configuration.py`

- [ ] Write the failing configuration test first. Resolve project-relative paths and assert the template plus three current source files exist when the Demo is located beside `资产评估工作流/`.
- [ ] Assert `data_manifest.yaml` records source name, business purpose, version/date, update method, and missing-data behavior without storing private file contents or secrets.
- [ ] Run the test and verify failure.
- [ ] Create the Tongfu project configuration, empty/example manual inputs, manifest, and initial changelog entry describing the 147-location mapping and Prompt version.
- [ ] Run the test and verify it passes.
- [ ] Commit with `docs: add Tongfu demo configuration and manifest`.

### Task 7: Real Run, README, and Final Acceptance

**Files:**
- Create: `demo/README.md`
- Modify: any files required by failures found during verification, with a failing regression test first.

- [ ] Run all tests: `uv run --extra dev pytest demo/tests -v`.
- [ ] Run the actual project offline: `uv run python -m demo.run demo/projects/tongfu.yaml --offline`.
- [ ] Verify the original template SHA-256 is unchanged.
- [ ] Verify the generated Word ZIP is valid, contains zero `X{2,}` placeholders, contains none of the 20 original yellow instruction texts, and exists outside the source directory.
- [ ] Render the generated Word and inspect all pages; use WPS if LibreOffice substitutes Chinese fonts incorrectly.
- [ ] Write README sections for business goal, workflow nodes, input/output contracts, setup, run command, output artifacts, artificial-review responsibility, c2m handoff boundary, and how to add a provider/template/project.
- [ ] Run the full suite again and verify no project warnings or failures.
- [ ] Commit with `docs: complete appraisal demo delivery`.

---

## Acceptance Checklist

- [ ] `demo/` contains every directory and file required by the handoff specification.
- [ ] `domain/` is framework-, file-, network-, environment-, and c2m-independent.
- [ ] Every workflow node declares input and output models with Chinese field documentation and examples.
- [ ] Ten representative fixture/expected pairs pass.
- [ ] Prompts and output schema are separately versioned.
- [ ] The original Word remains unchanged and every run creates a new Word.
- [ ] All 127 placeholders and 20 yellow instructions are replaced.
- [ ] OCR output is used without business-confidence gating.
- [ ] Missing external services still produce a reviewable Word.
- [ ] Audit artifacts trace values to sources or human-fill fallback.
- [ ] `data_manifest.yaml` and `CHANGELOG.md` are complete.
- [ ] README clearly states Demo/c2m reuse boundaries.
