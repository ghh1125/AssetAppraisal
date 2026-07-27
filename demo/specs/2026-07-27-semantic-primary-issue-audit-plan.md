# Semantic-First Excel and Review-Friendly Issues Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make semantic Excel extraction the primary source, reject unfinished all-zero appraisal columns, preserve complete evidence, and export a business-readable paginated issue workbook.

**Architecture:** Keep deterministic extraction in `demo/adapters/semantic_excel.py`, move source-selection and issue ordering rules into pure domain functions, and let `run_project` publish normalized evidence for `run_pipeline`. Preserve fixed coordinates only as fallback data. Export the issue workbook as a summary sheet plus a sorted detail sheet without changing the JSON-safe issue contract.

**Tech Stack:** Python 3.11, Pydantic, openpyxl, python-docx, pytest.

---

### Task 1: Reject unfinished all-zero appraisal columns

**Files:**
- Modify: `demo/domain/financial_table_semantics.py`
- Modify: `demo/adapters/semantic_excel.py`
- Test: `demo/tests/test_financial_table_semantics.py`
- Test: `demo/tests/test_semantic_excel.py`

- [ ] **Step 1: Write failing pure-rule tests**

Add tests that require an all-zero appraisal column with a non-zero book value to be classified as unfinished, while a zero net result remains valid when the appraisal column contains another non-zero amount.

```python
def test_all_zero_appraisal_column_is_unfinished():
    assert appraisal_zero_is_unfinished(
        book_value=120.0,
        appraised_value=0.0,
        appraisal_column_values=[0, 0, None],
    )


def test_zero_result_is_valid_when_column_contains_appraisal_activity():
    assert not appraisal_zero_is_unfinished(
        book_value=120.0,
        appraised_value=0.0,
        appraisal_column_values=[80, -80, 0],
    )
```

- [ ] **Step 2: Run the tests and verify RED**

Run:

```bash
PYTHONPATH=. /Users/ghh/Documents/Code/mcpify/AssetAppraisal/.venv/bin/pytest \
  demo/tests/test_financial_table_semantics.py -q
```

Expected: collection or assertion failure because `appraisal_zero_is_unfinished` does not exist.

- [ ] **Step 3: Implement the pure zero-validity rule**

Add:

```python
def appraisal_zero_is_unfinished(
    *,
    book_value: float | None,
    appraised_value: float | None,
    appraisal_column_values: Sequence[object],
) -> bool:
    if book_value in (None, 0) or appraised_value != 0:
        return False
    numeric = [
        float(value)
        for value in appraisal_column_values
        if isinstance(value, (int, float)) and not isinstance(value, bool)
    ]
    return not any(value != 0 for value in numeric)
```

- [ ] **Step 4: Write and verify failing adapter tests**

Create one workbook whose appraisal column is all zero and assert:

```python
assert "asset_approach_value" not in result["fields"]
assert any("[unfinished_appraisal]" in issue for issue in result["issues"])
```

Create a second workbook with appraisal activity elsewhere and assert a zero net result remains present.

- [ ] **Step 5: Apply the rule in semantic extraction**

Collect the selected appraisal column values in `_net_asset_candidate`. When the rule returns true, set `appraised` to `None`, retain the original locator as `unfinished_appraisal_locator`, and append:

```python
issues.append(
    "asset_approach_value：[unfinished_appraisal] "
    f"{path.name} {locator} 所在评估列为空或全零，已保留 XXX"
)
```

- [ ] **Step 6: Run focused tests and commit**

Run:

```bash
PYTHONPATH=. /Users/ghh/Documents/Code/mcpify/AssetAppraisal/.venv/bin/pytest \
  demo/tests/test_financial_table_semantics.py \
  demo/tests/test_semantic_excel.py -q
```

Expected: all pass.

Commit:

```bash
git add demo/domain/financial_table_semantics.py demo/adapters/semantic_excel.py \
  demo/tests/test_financial_table_semantics.py demo/tests/test_semantic_excel.py
git commit -m "fix: leave unfinished appraisal values unresolved"
```

### Task 2: Make semantic Excel evidence the primary source

**Files:**
- Create: `demo/domain/source_precedence.py`
- Modify: `demo/run.py`
- Modify: `demo/pipeline.py`
- Test: `demo/tests/test_source_precedence.py`
- Test: `demo/tests/test_demo_run.py`
- Test: `demo/tests/test_pipeline.py`

- [ ] **Step 1: Write failing precedence tests**

Define the desired interface:

```python
def test_semantic_source_replaces_fixed_coordinate_result():
    selected = prefer_semantic_result(
        fixed_value=1,
        fixed_evidence={"kind": "reporting_workbook", "locator": "旧表!A1"},
        semantic_value=2,
        semantic_evidence={
            "kind": "semantic_excel",
            "file": "新表.xlsx",
            "locator": "汇总表!D22",
        },
    )
    assert selected["value"] == 2


def test_fixed_coordinate_is_used_when_semantic_value_is_missing():
    selected = prefer_semantic_result(
        fixed_value=1,
        fixed_evidence={"kind": "reporting_workbook", "locator": "旧表!A1"},
        semantic_value=None,
        semantic_evidence={},
    )
    assert selected["value"] == 1
```

- [ ] **Step 2: Run the tests and verify RED**

Run:

```bash
PYTHONPATH=. /Users/ghh/Documents/Code/mcpify/AssetAppraisal/.venv/bin/pytest \
  demo/tests/test_source_precedence.py -q
```

Expected: collection failure because the module does not exist.

- [ ] **Step 3: Implement the pure precedence helper**

Return semantic value and evidence whenever the semantic value is not empty. Return the fixed pair only when semantic is absent. Keep historical table merging in `merge_historical_tables`.

- [ ] **Step 4: Add failing end-to-end evidence tests**

Add a synthetic workbook where the legacy coordinate contains a misleading number but semantic headers identify the correct value. Assert:

```python
assert fields["asset_scope_summary_table"] == expected_semantic_table
assert evidence["asset_scope_summary_table"] == {
    "kind": "semantic_excel",
    "file": workbook.name,
    "locator": "汇总表!C6；...",
}
```

Also assert `normalized_evidence.json` exists and `run_pipeline` trace candidates retain the same file and locator.

- [ ] **Step 5: Publish normalized evidence from `run_project`**

After `normalized_fields.json`, write:

```python
write_json(run_dir / "normalized_evidence.json", evidence)
```

Include the path in the run manifest outputs.

- [ ] **Step 6: Apply semantic-first selection in `run_project`**

For scalar and table fields returned by `extract_workbook_facts`, use semantic values first. For historical tables, merge compatible semantic tables and select the higher-quality explicit-period table. Do not allow a configured table to overwrite a valid semantic table.

- [ ] **Step 7: Preserve semantic evidence in `run_pipeline`**

Read `normalized_evidence.json` when present and use `_legacy_evidence` only as a backward-compatible fallback. In the scope-table compatibility block, bypass fixed-coordinate rereading whenever evidence kind starts with `semantic_excel`.

- [ ] **Step 8: Run focused tests and commit**

Run:

```bash
PYTHONPATH=. /Users/ghh/Documents/Code/mcpify/AssetAppraisal/.venv/bin/pytest \
  demo/tests/test_source_precedence.py \
  demo/tests/test_demo_run.py \
  demo/tests/test_pipeline.py -q
```

Expected: all pass.

Commit:

```bash
git add demo/domain/source_precedence.py demo/run.py demo/pipeline.py \
  demo/tests/test_source_precedence.py demo/tests/test_demo_run.py demo/tests/test_pipeline.py
git commit -m "feat: make semantic excel the primary source"
```

### Task 3: Export a review-friendly paginated issue workbook

**Files:**
- Modify: `demo/domain/generation_issues.py`
- Modify: `demo/adapters/generation_issues.py`
- Modify: `demo/tests/test_generation_issues.py`
- Modify: `demo/pipeline.py`
- Modify: `demo/run.py`

- [ ] **Step 1: Write failing organization tests**

Require page-aware ordering and a human-readable location:

```python
organized = organize_generation_issues(
    [
        {"page_number": "", "priority": "高", "location_description": "未定位"},
        {"page_number": 9, "priority": "中", "location_description": "利润表"},
        {"page_number": 5, "priority": "高", "location_description": "公司名称"},
    ]
)
assert [item["page_number"] for item in organized] == [5, 9, ""]
assert organized[0]["review_location"] == "第5页｜公司名称"
```

- [ ] **Step 2: Run the tests and verify RED**

Run:

```bash
PYTHONPATH=. /Users/ghh/Documents/Code/mcpify/AssetAppraisal/.venv/bin/pytest \
  demo/tests/test_generation_issues.py -q
```

Expected: failure because `organize_generation_issues` and the summary sheet do not exist.

- [ ] **Step 3: Implement pure issue organization**

Sort by page availability, numeric page, priority rank, and location ID. Add:

```python
item["review_location"] = (
    f"第{page}页｜{description}"
    if page not in (None, "")
    else f"页码待确认｜{description}"
)
item["review_action"] = item.get("suggestion", "人工核对并更新")
```

Translate `[unfinished_appraisal]` into the business-facing problem “疑似尚未完成评估，评估列为空或全零” and the action “确认评估工作簿是否已完成；完成后重新上传，未完成则保持黄色 XXX”。

- [ ] **Step 4: Upgrade the workbook export**

Create:

- `检查总览`: title, totals by priority, affected page count, page-level issue counts, and a three-step review instruction;
- `问题明细`: business columns first, technical columns last, freeze at `A2`, filter, wrapped text, priority/status colors, alternating row fills, and sensible widths.

Keep the old machine keys in JSON and technical detail columns for compatibility.

- [ ] **Step 5: Update pipeline and CLI callers**

Call `organize_generation_issues` after applying page locations and before exporting XLSX/JSON so both artifacts share the same order and user-facing descriptions.

- [ ] **Step 6: Run focused tests and commit**

Run:

```bash
PYTHONPATH=. /Users/ghh/Documents/Code/mcpify/AssetAppraisal/.venv/bin/pytest \
  demo/tests/test_generation_issues.py \
  demo/tests/test_pipeline.py \
  demo/tests/test_demo_run.py -q
```

Expected: all pass.

Commit:

```bash
git add demo/domain/generation_issues.py demo/adapters/generation_issues.py \
  demo/tests/test_generation_issues.py demo/pipeline.py demo/run.py
git commit -m "feat: make generation issues easier to review"
```

### Task 4: Documentation and full regression

**Files:**
- Modify: `demo/CHANGELOG.md`
- Modify: `README.md`
- Modify: `demo/README.md`
- Modify: `demo/data_manifest.yaml`
- Modify: `demo/tests/test_contracts.py`

- [ ] **Step 1: Update versioned behavior documentation**

Document `generation_issues.v2`, semantic-first precedence, unfinished-appraisal handling, and complete evidence propagation. Explain that fixed coordinates are fallback only.

- [ ] **Step 2: Update contract test**

Change the expected generation issue version to `generation_issues.v2` and verify the trace and manifest report the new version.

- [ ] **Step 3: Run all backend tests**

Run:

```bash
PYTHONPATH=. /Users/ghh/Documents/Code/mcpify/AssetAppraisal/.venv/bin/pytest demo/tests -q
```

Expected: all pass.

- [ ] **Step 4: Re-run four real projects without OCR**

Use the existing latest Excel materials for 亦盛、夏弗纳、铁投能源、浙江晶引. Assert:

- 浙江晶引 no longer outputs an asset-basis result of zero;
- every filled structured table has source file and locator;
- all issue workbook page numbers are populated when the page reader is enabled;
- each issue workbook contains `检查总览` and `问题明细`.

- [ ] **Step 5: Render and inspect changed Word pages**

Render all four reports and inspect pages containing the final conclusion, historical tables, long-term assets, and highlighted unresolved values. Confirm no overflow, overlap, or formatting regression.

- [ ] **Step 6: Commit documentation**

```bash
git add README.md demo/README.md demo/CHANGELOG.md demo/data_manifest.yaml \
  demo/tests/test_contracts.py
git commit -m "docs: document semantic-first review workflow"
```

