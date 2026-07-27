# Semantic Excel Coverage Improvement Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Improve historical financial-statement and long-term-asset coverage across varying Excel layouts without reducing the accuracy or traceability of existing appraisal totals.

**Architecture:** Keep `openpyxl` access in `demo/adapters/semantic_excel.py`, and move reusable label, period, candidate-ranking, and safe-derivation rules into a pure `demo/domain/financial_table_semantics.py` module. The adapter emits values together with exact source locators; uncertain candidates remain unresolved. Existing pipeline, frontend, and Word-writing interfaces remain unchanged.

**Tech Stack:** Python 3.11, openpyxl, pytest, Pydantic-compatible JSON data, YAML project configuration.

---

### Task 1: Change the default Bailian model

**Files:**
- Modify: `demo/tests/test_llm_config.py`
- Modify: `demo/domain/llm_config.py`
- Modify: `demo/adapters/bailian_glm.py`
- Modify: `demo/adapters/llm_review.py`
- Modify: `demo/projects/tongfu.yaml`
- Modify: `demo/data_manifest.yaml`
- Modify: `.env.example`
- Modify: `README.md`
- Modify: `demo/README.md`
- Modify: `demo/CHANGELOG.md`

- [ ] **Step 1: Write the failing default-model test**

Add:

```python
def test_llm_models_use_current_project_default_when_no_override_exists():
    models = resolve_llm_models({}, {})
    assert set(models.values()) == {"qwen3.7-max-2026-05-17"}
```

- [ ] **Step 2: Run the test and verify RED**

Run:

```bash
.venv/bin/pytest demo/tests/test_llm_config.py::test_llm_models_use_current_project_default_when_no_override_exists -q
```

Expected: FAIL because the current fallback is `qwen3.7-flash`.

- [ ] **Step 3: Replace only default model literals**

Define in `demo/domain/llm_config.py`:

```python
DEFAULT_LLM_MODEL = "qwen3.7-max-2026-05-17"
```

Use it for the configured fallback and environment fallback. Update adapter constructor defaults, project YAML, manifest, environment example, README examples, Demo README, and changelog to the same model. Preserve all four per-task environment overrides.

- [ ] **Step 4: Run focused tests and verify GREEN**

Run:

```bash
.venv/bin/pytest demo/tests/test_llm_config.py demo/tests/test_llm_factory.py demo/tests/test_bailian_glm.py demo/tests/test_llm_review.py -q
```

Expected: all tests pass after updating explicit default-model expectations; task-specific override tests remain unchanged.

- [ ] **Step 5: Commit**

```bash
git add .env.example README.md demo/README.md demo/CHANGELOG.md demo/data_manifest.yaml demo/projects/tongfu.yaml demo/domain/llm_config.py demo/adapters/bailian_glm.py demo/adapters/llm_review.py demo/tests/test_llm_config.py demo/tests/test_llm_factory.py demo/tests/test_bailian_glm.py demo/tests/test_llm_review.py
git commit -m "feat: use qwen 3.7 max as default appraisal model"
```

### Task 2: Add pure period and label semantics

**Files:**
- Create: `demo/domain/financial_table_semantics.py`
- Create: `demo/tests/test_financial_table_semantics.py`

- [ ] **Step 1: Write failing tests for period recognition**

Create tests:

```python
from demo.domain.financial_table_semantics import (
    canonical_period,
    choose_historical_columns,
)


def test_canonical_period_accepts_dates_and_rejects_rates():
    assert canonical_period("2025年6月30日") == (2025, 6, 30, "2025年6月30日")
    assert canonical_period("2024年度") == (2024, 12, 31, "2024年度")
    assert canonical_period("期末数") == (None, None, None, "期末数")
    assert canonical_period("增长率%") is None


def test_historical_columns_prefer_actual_periods_over_growth_and_forecast():
    headers = {
        2: ["2023年度", "实际数"],
        3: ["2024年度", "审定数"],
        4: ["同比增长率"],
        5: ["2025年度", "预测数"],
    }
    assert choose_historical_columns(headers, valuation_year=2024) == [2, 3]
```

- [ ] **Step 2: Run tests and verify RED**

Run:

```bash
.venv/bin/pytest demo/tests/test_financial_table_semantics.py -q
```

Expected: collection fails because the module does not exist.

- [ ] **Step 3: Implement the pure period rules**

Implement:

```python
def canonical_period(value: object) -> tuple[int | None, int | None, int | None, str] | None:
    text = str(value or "").strip()
    if not text or any(token in text for token in ("增长", "占比", "比例", "%", "预测", "预算")):
        return None
    match = re.search(r"(20\\d{2})年(?:(\\d{1,2})月)?(?:(\\d{1,2})日)?", text)
    if match:
        year = int(match.group(1))
        month = int(match.group(2) or 12)
        day = int(match.group(3) or 31)
        return year, month, day, text
    if text in {"期初数", "期末数", "年初数", "年末数"}:
        return None, None, None, text
    return None
```

`choose_historical_columns` must reject forecast/budget/growth columns, prefer actual/audited columns, and return columns sorted by normalized period.

- [ ] **Step 4: Run tests and verify GREEN**

Run:

```bash
.venv/bin/pytest demo/tests/test_financial_table_semantics.py -q
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add demo/domain/financial_table_semantics.py demo/tests/test_financial_table_semantics.py
git commit -m "feat: add financial period semantics"
```

### Task 3: Make historical tables period-first and safely derive equity

**Files:**
- Modify: `demo/tests/test_semantic_excel.py`
- Modify: `demo/adapters/semantic_excel.py`
- Modify: `demo/domain/financial_table_semantics.py`

- [ ] **Step 1: Write failing adapter tests**

Add a workbook with a multi-row header, a growth column, and no explicit equity row:

```python
def test_history_uses_period_columns_and_derives_missing_equity(tmp_path: Path):
    path = _save_workbook(
        tmp_path / "multi-header.xlsx",
        {
            "资产负债表": [
                ["金额单位：人民币元"],
                ["项目", "2023年度", "2024年度", "2024年度", "2025年度"],
                [None, "实际数", "审定数", "增长率%", "预测数"],
                ["资产总计", 100, 200, 1.0, 300],
                ["负债合计", 40, 70, 0.75, 100],
            ]
        },
    )
    facts = extract_workbook_facts(path, "audited_financials")
    rows = facts["fields"]["historical_balance_sheet_table"]["rows"]
    assert rows[0] == ["项目\\报表日", "历史期1", "2023年度", "2024年度"]
    assert rows[1] == ["总资产", "XXX", "100.00", "200.00"]
    assert rows[3] == ["所有者权益", "XXX", "60.00", "130.00"]
    assert "derived" in facts["evidence"]["historical_balance_sheet_table"]["kind"]
```

Add a second test ensuring an amount such as `56129.203566` can never become a period header.

- [ ] **Step 2: Run the two tests and verify RED**

Run:

```bash
.venv/bin/pytest demo/tests/test_semantic_excel.py -k "period_columns or amount_as_period" -q
```

Expected: FAIL because the adapter currently takes the last three numbers to the right of a label.

- [ ] **Step 3: Implement period-first extraction**

Replace `_values_right_of` and `_period_headers` usage with:

1. reconstruct up to four header rows above each candidate table;
2. call `choose_historical_columns`;
3. read only selected amount columns for every canonical row;
4. align values by selected column instead of independently taking the last three numbers;
5. merge rows from candidate sheets by canonical period;
6. derive equity only when both assets and liabilities exist for the same period.

Evidence locators must include the label cells and selected amount-cell coordinates. Derived equity evidence must include both operand coordinates and `kind="semantic_excel_derived"`.

- [ ] **Step 4: Run historical-table tests**

Run:

```bash
.venv/bin/pytest demo/tests/test_semantic_excel.py -k "historical or balance or income_workbook_can_supply_history" -q
```

Expected: PASS, including existing dual-sided statements.

- [ ] **Step 5: Commit**

```bash
git add demo/adapters/semantic_excel.py demo/domain/financial_table_semantics.py demo/tests/test_semantic_excel.py
git commit -m "feat: align historical statements by period"
```

### Task 4: Add long-term-asset detail aggregation

**Files:**
- Modify: `demo/tests/test_semantic_excel.py`
- Modify: `demo/domain/financial_table_semantics.py`
- Modify: `demo/adapters/semantic_excel.py`

- [ ] **Step 1: Write failing tests for alias and detail aggregation**

Add:

```python
def test_long_term_assets_fall_back_to_detail_rows(tmp_path: Path):
    path = _save_workbook(
        tmp_path / "asset-detail-only.xlsx",
        {
            "固定资产明细表": [
                ["金额单位：人民币元"],
                ["资产编号", "资产名称", "资产类别", "账面原值", "累计折旧", "账面净值", "评估值"],
                ["D001", "电脑", "办公电子设备", 10_000, 2_000, 8_000, 8_500],
                ["D002", "打印机", "电子设备类", 5_000, 1_000, 4_000, 4_200],
                ["", "合计", "", 15_000, 3_000, 12_000, 12_700],
            ],
            "汇总表": [
                ["金额单位：人民币元"],
                ["项目", "账面价值", "评估价值"],
                ["无形资产", 30_000, 35_000],
                ["长期待摊费用", 10_000, 10_000],
            ],
        },
    )
    facts = extract_workbook_facts(path, "reporting_workbook")
    rows = facts["fields"]["long_term_assets_table"]["rows"]
    assert rows[1] == ["电子设备", "12,000.00", "2项", "以评估明细表为准"]
    assert "固定资产明细表!F3:F4" in facts["evidence"]["long_term_assets_table"]["locator"]
```

- [ ] **Step 2: Run the test and verify RED**

Run:

```bash
.venv/bin/pytest demo/tests/test_semantic_excel.py::test_long_term_assets_fall_back_to_detail_rows -q
```

Expected: FAIL because only exact `电子设备` rows in specially named summary sheets are currently accepted.

- [ ] **Step 3: Implement safe detail aggregation**

Add pure aliases in `financial_table_semantics.py` and adapter scanning that:

- recognizes identifier, name, category, book cost, accumulated depreciation, book net value, and appraisal-value headers;
- requires a unique book-net-value column;
- excludes total rows from item counting and summation;
- matches only canonical electronic-equipment categories;
- sums detail book-net values;
- records a compact contiguous range when possible, otherwise records individual cells;
- returns quantity only when valid detail rows are present.

Use the existing summary-table value first. Use detail aggregation only when summary extraction returns `None`.

- [ ] **Step 4: Run long-term and existing semantic tests**

Run:

```bash
.venv/bin/pytest demo/tests/test_semantic_excel.py -q
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add demo/adapters/semantic_excel.py demo/domain/financial_table_semantics.py demo/tests/test_semantic_excel.py
git commit -m "feat: aggregate long-term assets from detail sheets"
```

### Task 5: Record unresolved reasons without blocking Word generation

**Files:**
- Modify: `demo/tests/test_semantic_excel.py`
- Modify: `demo/adapters/semantic_excel.py`
- Modify: `demo/pipeline.py`
- Modify: `demo/schemas.py`
- Modify: `demo/CHANGELOG.md`

- [ ] **Step 1: Write failing unresolved-reason test**

Create a workbook with two equally plausible electronic-equipment amount columns and assert:

```python
assert facts["fields"]["long_term_assets_table"]["rows"][1][1] == "XXX"
assert any(
    issue["field_key"] == "long_term_assets_table"
    and issue["reason"] == "ambiguous_candidate"
    for issue in facts["issues"]
)
```

- [ ] **Step 2: Run the test and verify RED**

Run:

```bash
.venv/bin/pytest demo/tests/test_semantic_excel.py -k ambiguous -q
```

Expected: FAIL because semantic extraction currently always returns an empty issue list.

- [ ] **Step 3: Add JSON-serializable issue metadata**

Return issue dictionaries containing:

```python
{
    "field_key": "long_term_assets_table",
    "reason": "ambiguous_candidate",
    "source_file": path.name,
    "source_locator": "固定资产明细表",
    "message": "电子设备存在多个无法区分的账面金额列，已保留 XXX",
}
```

Allow reasons `source_absent`, `formula_unavailable`, `ambiguous_candidate`, and `unmatched_label`. Convert them into the existing human-readable pipeline issue list and final page-mapped problem list. These issues must never stop Word generation.

- [ ] **Step 4: Run pipeline and semantic tests**

Run:

```bash
.venv/bin/pytest demo/tests/test_semantic_excel.py demo/tests/test_pipeline.py -q
```

Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add demo/adapters/semantic_excel.py demo/pipeline.py demo/schemas.py demo/tests/test_semantic_excel.py demo/CHANGELOG.md
git commit -m "feat: explain unresolved excel fields"
```

### Task 6: Run full regression and four-project local acceptance

**Files:**
- Modify only if a regression exposes a defect in files already listed above.
- Do not add real customer workbooks, PDFs, generated reports, OCR output, API keys, or run directories to Git.

- [ ] **Step 1: Run the complete backend test suite**

Run:

```bash
.venv/bin/pytest demo/tests -q
```

Expected: all tests pass.

- [ ] **Step 2: Run frontend regression**

Run:

```bash
npm test -- --run
npm run build
```

from `frontend/`.

Expected: tests and build pass.

- [ ] **Step 3: Run the four local projects without OCR or external APIs**

Reuse the existing four-project validation command and cached/latest Excel materials. Do not start PaddleOCR. Generate each Word, workflow trace, audit manifest, and page-mapped issue list into a new timestamped directory under `runs/`.

- [ ] **Step 4: Compare before and after**

Report per project:

- the 12 core values and source cells;
- historical table filled/total cells;
- long-term table filled/total cells;
- wrong-period-header count;
- unresolved reasons by category;
- generated Word and problem-list paths.

Required result: core values remain 12/12 exact; wrong-period-header count becomes zero; no newly filled amount lacks source-cell evidence.

- [ ] **Step 5: Final consistency check**

Run:

```bash
git diff --check
git status --short
```

Confirm `.env`, real materials, generated reports, OCR data, and `runs/` remain ignored.

- [ ] **Step 6: Commit any regression-only fixes**

```bash
git add demo/adapters/semantic_excel.py demo/domain/financial_table_semantics.py demo/pipeline.py demo/schemas.py demo/tests/test_semantic_excel.py demo/tests/test_financial_table_semantics.py demo/CHANGELOG.md
git commit -m "fix: preserve appraisal excel regression coverage"
```
