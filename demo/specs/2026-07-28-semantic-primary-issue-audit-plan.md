# Excel Semantic Priority and Issue Audit Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make semantic Excel extraction the primary path, reject incomplete all-zero appraisal columns, preserve complete evidence, and produce a business-readable issue workbook.

**Architecture:** Keep semantic matching pure and deterministic in `demo/adapters/semantic_excel.py`; centralize source precedence and evidence serialization in `demo/run.py` and `demo/pipeline.py`; keep presentation-only workbook formatting in `demo/adapters/generation_issues.py`. Fixed project coordinates remain a fallback and never overwrite an available semantic result.

**Tech Stack:** Python 3.11, Pydantic, openpyxl, pytest, python-docx/OOXML pipeline.

---

### Task 1: Reject incomplete all-zero appraisal columns

**Files:**
- Modify: `demo/adapters/semantic_excel.py`
- Test: `demo/tests/test_semantic_excel.py`

- [ ] **Step 1: Write failing tests**

Create one workbook whose net-asset row has nonzero book value and zero appraisal value while the full appraisal column is blank/zero. Assert `asset_approach_value` is absent and issues contain `[incomplete_appraisal_column]`. Create a second workbook with zero net assets but another nonzero appraisal amount in the same column and assert zero remains a valid extracted result.

- [ ] **Step 2: Run tests and verify RED**

Run:

```bash
PYTHONPATH=. /Users/ghh/Documents/Code/mcpify/AssetAppraisal/.venv/bin/pytest demo/tests/test_semantic_excel.py -q
```

Expected: the incomplete-column test fails because the current extractor returns `0.0`.

- [ ] **Step 3: Implement the minimum rule**

Add a pure helper that returns true only when the candidate appraisal value is zero, book value is nonzero, and the selected appraisal column contains no other nonzero numeric value. Skip `asset_approach_value` and append an `[incomplete_appraisal_column]` issue when true.

- [ ] **Step 4: Run tests and verify GREEN**

Run the focused command from Step 2. Expected: all semantic Excel tests pass.

- [ ] **Step 5: Commit**

```bash
git add demo/adapters/semantic_excel.py demo/tests/test_semantic_excel.py
git commit -m "fix: reject incomplete appraisal columns"
```

### Task 2: Make semantic extraction primary and preserve evidence

**Files:**
- Modify: `demo/run.py`
- Modify: `demo/pipeline.py`
- Modify: `demo/adapters/audit.py`
- Test: `demo/tests/test_demo_run.py`
- Test: `demo/tests/test_pipeline.py`

- [ ] **Step 1: Write failing precedence and evidence tests**

Create a workbook where a configured legacy coordinate contains a different valid number while the semantic net-assets row contains the intended number. Assert the normalized field uses semantic evidence. Run a pipeline fixture containing a semantic table and assert `normalized_evidence.json`, workflow trace, and field audit preserve the source filename and cell locator.

- [ ] **Step 2: Run tests and verify RED**

Run:

```bash
PYTHONPATH=. /Users/ghh/Documents/Code/mcpify/AssetAppraisal/.venv/bin/pytest demo/tests/test_demo_run.py demo/tests/test_pipeline.py -q
```

Expected: evidence-file assertions fail and at least one fixed-coordinate precedence assertion fails.

- [ ] **Step 3: Implement semantic-first merging**

For semantic fields, assign semantic values before accepting fixed fallback values. Historical tables use `merge_historical_tables`; scalar and structured tables use semantic values whenever present. Write `normalized_evidence.json` beside `normalized_fields.json`.

- [ ] **Step 4: Read evidence directly in the pipeline**

Load `normalized_evidence.json` from the temporary base run. Preserve evidence kinds beginning with `semantic_excel` and remove scope-table logic that overwrites semantic evidence with missing/fixed metadata.

- [ ] **Step 5: Run focused tests and verify GREEN**

Run the command from Step 2. Expected: all selected tests pass.

- [ ] **Step 6: Commit**

```bash
git add demo/run.py demo/pipeline.py demo/adapters/audit.py demo/tests/test_demo_run.py demo/tests/test_pipeline.py
git commit -m "feat: prioritize semantic excel evidence"
```

### Task 3: Redesign the generated issue workbook

**Files:**
- Modify: `demo/domain/generation_issues.py`
- Modify: `demo/adapters/generation_issues.py`
- Test: `demo/tests/test_generation_issues.py`

- [ ] **Step 1: Write failing workbook-layout tests**

Export unsorted issues with pages `10`, `2`, and blank. Assert the workbook contains `检查总览` and `问题明细`; assert detail rows are ordered page 2, page 10, blank; assert the first columns are `优先级、处理状态、Word页码、检查位置`; assert totals and priority counts are present in the overview.

- [ ] **Step 2: Run tests and verify RED**

Run:

```bash
PYTHONPATH=. /Users/ghh/Documents/Code/mcpify/AssetAppraisal/.venv/bin/pytest demo/tests/test_generation_issues.py -q
```

Expected: the overview sheet is missing and the existing column order does not match.

- [ ] **Step 3: Add business display fields and stable sorting**

Create `inspection_location` as `第N页｜位置类型｜位置描述`, or `页码待定位｜...` when unavailable. Sort by numeric page, priority rank, and location ID without mutating the input list.

- [ ] **Step 4: Implement the two-sheet workbook**

Build `检查总览` with counts and instructions. Build `问题明细` with business columns first, technical columns last, filters, frozen panes, print setup, row heights, alternating fills, and priority colors.

- [ ] **Step 5: Run tests and verify GREEN**

Run the command from Step 2. Expected: all issue-workbook tests pass.

- [ ] **Step 6: Commit**

```bash
git add demo/domain/generation_issues.py demo/adapters/generation_issues.py demo/tests/test_generation_issues.py
git commit -m "feat: make issue checklist reviewer friendly"
```

### Task 4: Cross-project regression and documentation

**Files:**
- Modify: `demo/CHANGELOG.md`
- Test: `demo/tests/`

- [ ] **Step 1: Run the full backend suite**

```bash
PYTHONPATH=. /Users/ghh/Documents/Code/mcpify/AssetAppraisal/.venv/bin/pytest demo/tests -q
```

Expected: 100% pass with no failures.

- [ ] **Step 2: Re-run the four latest-data projects without OCR**

Use the existing cached/local Excel materials for 亦盛、夏弗纳、铁投能源、浙江晶引. Assert all four create a Word report, field audit, two-sheet issue workbook, JSON issue list, normalized fields, and normalized evidence.

- [ ] **Step 3: Verify content and formatting**

Confirm Zhejiang no longer contains the invalid asset-approach zero conclusion. Compare filled core values and historical-table cells to evidence locators. Render all changed Word pages and inspect for overflow, clipping, table-width changes, or pagination regressions.

- [ ] **Step 4: Update change log**

Document semantic-first precedence, incomplete appraisal-column handling, complete evidence propagation, and the two-sheet reviewer checklist.

- [ ] **Step 5: Run final verification and commit**

```bash
git diff --check
PYTHONPATH=. /Users/ghh/Documents/Code/mcpify/AssetAppraisal/.venv/bin/pytest demo/tests -q
git add demo/CHANGELOG.md
git commit -m "docs: record semantic audit improvements"
```
