# Soft Financial Validation Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Generate a reviewable Word report when configured financial fields are missing, while recording blanks and high-priority issues and suppressing the final-candidate report.

**Architecture:** Keep missing-field decisions in pure functions under `demo/domain/`. The CLI extraction and end-to-end pipeline consume the same policy, preserve blank evidence, and expose validation status through audit, trace, issues, and manifest outputs.

**Tech Stack:** Python 3.11, Pydantic v2, pytest, python-docx, existing JSON project configuration.

---

### Task 1: Define the pure missing-field policy

**Files:**
- Modify: `demo/domain/field_validation.py`
- Modify: `demo/domain/narrative_policy.py`
- Test: `demo/tests/test_financial_matching.py`
- Test: `demo/tests/test_narrative_policy.py`

- [ ] **Step 1: Write failing tests**

```python
def test_missing_fields_are_blank_with_high_priority_issues():
    result = apply_missing_field_policy(
        {"present": 1},
        {"present": {"kind": "excel", "file": "source.xlsx", "locator": "A1"}},
        ["present", "missing"],
        "金额及财务结果字段",
    )
    assert result["valid"] is False
    assert result["fields"]["missing"] == ""
    assert result["evidence"]["missing"]["kind"] == "blank"
    assert result["issues"] == [
        "高优先级：金额及财务结果字段未匹配到，已留空：missing"
    ]


def test_candidate_report_is_suppressed_when_financial_fields_are_incomplete():
    reviews = {"data": {"status": "completed", "findings": []}}
    assert should_create_candidate_report(
        reviews,
        financial_fields_complete=False,
    ) is False
```

- [ ] **Step 2: Run focused tests and verify RED**

```bash
UV_CACHE_DIR=/tmp/uv-cache uv run --python 3.11 python -m pytest \
  demo/tests/test_financial_matching.py \
  demo/tests/test_narrative_policy.py -q
```

Expected: missing policy import or unsupported keyword failure.

- [ ] **Step 3: Implement the policy**

```python
def apply_missing_field_policy(
    fields: dict[str, Any],
    evidence: dict[str, dict[str, Any]],
    required_fields: list[str],
    label: str,
) -> dict[str, Any]:
    updated_fields = dict(fields)
    updated_evidence = {key: dict(value) for key, value in evidence.items()}
    missing = sorted(
        key
        for key in required_fields
        if updated_fields.get(key) in (None, "", [], {})
    )
    for key in missing:
        updated_fields[key] = ""
        updated_evidence[key] = {
            "kind": "blank",
            "file": "",
            "locator": "指定来源未匹配到值",
        }
    return {
        "valid": not missing,
        "missing_fields": missing,
        "fields": updated_fields,
        "evidence": updated_evidence,
        "issues": [
            f"高优先级：{label}未匹配到，已留空：{key}"
            for key in missing
        ],
    }
```

Extend `should_create_candidate_report` with a keyword-only `financial_fields_complete: bool = True` and require it to be true.

- [ ] **Step 4: Run tests and verify GREEN**

Run the focused command from Step 2. Expected: all tests pass.

### Task 2: Remove the base extraction hard stop

**Files:**
- Modify: `demo/run.py`
- Test: `demo/tests/test_demo_run.py`

- [ ] **Step 1: Write a failing integration test**

Create a temporary project configuration with absolute material paths and append `synthetic_missing_amount` to `required_financial_fields`.

```python
result = run_project(config_path, output_dir=tmp_path / "run", offline=True)
fields = json.loads((tmp_path / "run/normalized_fields.json").read_text())
assert result.report_path.exists()
assert fields["synthetic_missing_amount"] == ""
assert any(
    "高优先级：财务材料字段未匹配到，已留空：synthetic_missing_amount" == issue
    for issue in result.issues
)
```

- [ ] **Step 2: Run the new test and verify RED**

Expected: `ValueError: 财务材料字段未能提取`.

- [ ] **Step 3: Apply the pure policy in `run_project`**

Replace the `raise ValueError` block with `apply_missing_field_policy`, assign its returned fields/evidence, and append its issues. Do not change corrupt-workbook or invalid-template exceptions.

- [ ] **Step 4: Run Demo tests and verify GREEN**

```bash
UV_CACHE_DIR=/tmp/uv-cache uv run --python 3.11 python -m pytest \
  demo/tests/test_demo_run.py demo/tests/test_financial_tables.py -q
```

Expected: all tests pass.

### Task 3: Make the end-to-end monetary gate soft

**Files:**
- Modify: `demo/pipeline.py`
- Modify: `demo/domain/financial_matching.py`
- Test: `demo/tests/test_pipeline.py`
- Test: `demo/tests/test_financial_matching.py`

- [ ] **Step 1: Write a failing pipeline test**

Wrap the real `run_project` in the test and blank `book_net_assets` in its normalized output before the outer pipeline continues. Run the three review adapters and assert:

```python
assert result.report_path.exists()
assert not (tmp_path / "资产评估报告_最终候选.docx").exists()
assert json.loads((tmp_path / "normalized_fields.json").read_text())["book_net_assets"] == ""
assert any("高优先级：金额及财务结果字段未匹配到，已留空：book_net_assets" == issue for issue in result.issues)
assert json.loads(result.manifest_path.read_text())["financial_validation"]["valid"] is False
```

Also assert the `fill_word` trace node is `completed_with_issues` and the audit row uses source kind `blank`.

- [ ] **Step 2: Run the new test and verify RED**

Expected: current monetary gate raises and no Word is returned.

- [ ] **Step 3: Replace the monetary exception with the soft policy**

Apply `apply_missing_field_policy` before replacements. Preserve the returned validation object as:

```python
financial_validation = {
    "valid": monetary_policy["valid"],
    "missing_fields": monetary_policy["missing_fields"],
    "conflicts": list(fields.get("_conflicts", [])),
}
```

Use `completed_with_issues` for `fill_word` when validation is not valid. Pass `financial_fields_complete=financial_validation["valid"]` to candidate-report policy. Add `financial_validation` to `run_manifest.json`.

- [ ] **Step 4: Ensure missing configured financial tables clear template defaults**

Add a pure helper:

```python
def blank_configured_table(spec: dict[str, Any]) -> list[list[str]]:
    matrix = [list(spec["header"])] if spec.get("include_header", True) else []
    for row in spec["rows"]:
        matrix.append([str(row["label"]), *([""] * len(row["cells"]))])
    return matrix
```

When a configured financial-table field is not a populated `{caption, rows}` value, pass this blank matrix to the Word table replacement adapter.

- [ ] **Step 5: Run pipeline tests and verify GREEN**

```bash
UV_CACHE_DIR=/tmp/uv-cache uv run --python 3.11 python -m pytest \
  demo/tests/test_pipeline.py demo/tests/test_financial_matching.py -q
```

Expected: all tests pass.

### Task 4: Update versions, examples, and documentation

**Files:**
- Modify: `demo/data_manifest.yaml`
- Modify: `demo/fixtures/workflow_cases.yaml`
- Modify: `demo/expected/workflow_cases.yaml`
- Modify: `demo/tests/test_domain_examples.py`
- Modify: `demo/README.md`
- Modify: `README.md`
- Modify: `demo/CHANGELOG.md`

- [ ] **Step 1: Add a boundary example for incomplete financial validation**

Add one deterministic example demonstrating that completed reviews plus incomplete financial fields do not produce a candidate report.

- [ ] **Step 2: Register `financial_validation.v2`**

Add `"financial_validation": "financial_validation.v2"` to `rule_versions` and record the soft-validation behavior in the changelog.

- [ ] **Step 3: Update user documentation**

State that unmatched fields remain blank, generate a high-priority issue, and allow the待复核 Word; corrupt files and contract/template failures still stop; incomplete financial data suppresses only the final-candidate Word.

- [ ] **Step 4: Run contract and example tests**

```bash
UV_CACHE_DIR=/tmp/uv-cache uv run --python 3.11 python -m pytest \
  demo/tests/test_contracts.py demo/tests/test_domain_examples.py -q
```

Expected: all tests pass.

### Task 5: Full verification and delivery

**Files:**
- Verify all modified files.

- [ ] **Step 1: Run the complete Python suite**

```bash
UV_CACHE_DIR=/tmp/uv-cache uv run --python 3.11 python -m pytest -q
```

Expected: all tests pass.

- [ ] **Step 2: Compile and build**

```bash
UV_CACHE_DIR=/tmp/uv-cache uv run --python 3.11 python -m compileall -q demo
cd frontend
npm run build
```

Expected: Python compilation and Vite build succeed.

- [ ] **Step 3: Check privacy and repository hygiene**

```bash
git diff --check
git ls-files | rg '(^|/)\.env$|runs/|资产评估工作流|docs/superpowers'
```

Expected: no secrets, customer materials, generated runs, or private superpower documents are tracked.

- [ ] **Step 4: Commit and push**

```bash
git add README.md demo
git commit -m "fix: generate review report with missing financial fields"
git push origin main
```
