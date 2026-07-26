# Workflow Contract Compliance Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the asset-appraisal Demo prove at runtime that every executed workflow node follows a documented input/output contract and retains rule, Prompt, model, data, artifact, issue, and evidence lineage.

**Architecture:** Keep pure validation and policy functions in `demo/domain/`; keep JSON/file writing in `demo/adapters/`; use `demo/pipeline.py` only to orchestrate and record node results. Add contract and trace schemas to `demo/schemas.py`, validate `demo/workflow.yaml` before execution, and export one `workflow_trace.json` per run.

**Tech Stack:** Python 3.11, Pydantic v2, pytest, JSON-compatible YAML files, existing FastAPI/Vue shell.

---

## File map

- Create `demo/domain/workflow_contracts.py`: pure workflow graph and schema-metadata validation.
- Create `demo/domain/narrative_policy.py`: pure module-to-field and final-candidate policy.
- Create `demo/adapters/workflow_trace.py`: in-memory trace recorder and JSON export.
- Create `demo/fixtures/workflow_cases.yaml`: normal and boundary examples for new rules.
- Create `demo/expected/workflow_cases.yaml`: deterministic expected results.
- Create `demo/tests/test_workflow_contracts.py`: contract validation tests.
- Create `demo/tests/test_narrative_policy.py`: module-selection policy tests.
- Create `demo/tests/test_workflow_trace.py`: trace recorder tests.
- Modify `demo/schemas.py`: workflow definition and trace models.
- Modify `demo/workflow.yaml`: explicit contract version.
- Modify `demo/pipeline.py`: validate the workflow and record actual nodes.
- Modify `demo/api_server.py`: expose `workflow_trace.json` as a downloadable artifact.
- Modify `demo/tests/test_pipeline.py`: end-to-end trace assertions.
- Modify `demo/tests/test_domain_examples.py`: execute the new representative examples.
- Modify `demo/data_manifest.yaml`: register contract/rule/output versions.
- Modify `demo/README.md`, `README.md`, and `demo/CHANGELOG.md`: document behavior and handoff boundary.

### Task 1: Validate workflow declarations as pure business data

**Files:**
- Create: `demo/domain/workflow_contracts.py`
- Modify: `demo/schemas.py`
- Modify: `demo/workflow.yaml`
- Test: `demo/tests/test_workflow_contracts.py`

- [ ] **Step 1: Write failing tests for valid and invalid workflow declarations**

```python
def test_validates_current_workflow_models_and_dependencies():
    definition = json.loads(Path("demo/workflow.yaml").read_text(encoding="utf-8"))
    result = validate_workflow_contract(definition, schemas)
    assert result == {"valid": True, "issues": [], "node_count": 13}


def test_rejects_unknown_model_and_dependency_cycle():
    definition = {
        "version": "test",
        "contract_version": "workflow_contract.v1",
        "nodes": [
            {
                "name": "a",
                "input_model": "MissingInput",
                "output_model": "InventoryOutput",
                "depends_on": ["b"],
                "human_checkpoint": None,
            },
            {
                "name": "b",
                "input_model": "InventoryInput",
                "output_model": "InventoryOutput",
                "depends_on": ["a"],
                "human_checkpoint": None,
            },
        ],
    }
    result = validate_workflow_contract(definition, schemas)
    assert not result["valid"]
    assert "a：输入模型不存在：MissingInput" in result["issues"]
    assert "工作流依赖存在环" in result["issues"]
```

- [ ] **Step 2: Run the focused tests and verify RED**

Run:

```bash
UV_CACHE_DIR=/tmp/uv-cache uv run --python 3.11 python -m pytest demo/tests/test_workflow_contracts.py -q
```

Expected: collection fails because `demo.domain.workflow_contracts` does not exist.

- [ ] **Step 3: Add typed workflow declaration models**

Add to `demo/schemas.py`:

```python
class WorkflowNodeDefinition(DemoModel):
    name: str = Field(description="工作流节点名称", examples=["ocr_pdf"])
    input_model: str = Field(description="节点输入模型名称", examples=["OcrPdfInput"])
    output_model: str = Field(description="节点输出模型名称", examples=["OcrPdfOutput"])
    depends_on: list[str] = Field(description="前置节点名称", examples=[["inventory"]])
    human_checkpoint: str | None = Field(
        default=None,
        description="节点完成后的人工确认要求",
        examples=["评估师审核问题清单"],
    )


class WorkflowDefinition(DemoModel):
    version: str = Field(description="工作流业务版本", examples=["1.1.0"])
    contract_version: str = Field(description="节点契约规则版本", examples=["workflow_contract.v1"])
    nodes: list[WorkflowNodeDefinition] = Field(description="有序工作流节点", examples=[[]])
```

Add `"contract_version": "workflow_contract.v1"` to `demo/workflow.yaml`.

- [ ] **Step 4: Implement pure contract validation**

Implement in `demo/domain/workflow_contracts.py`:

```python
def validate_workflow_contract(
    payload: Mapping[str, Any],
    schema_module: Any,
) -> dict[str, Any]:
    issues: list[str] = []
    try:
        definition = WorkflowDefinition.model_validate(payload)
    except ValidationError as exc:
        return {"valid": False, "issues": [str(exc)], "node_count": 0}

    names = [node.name for node in definition.nodes]
    if len(names) != len(set(names)):
        issues.append("工作流节点名称重复")
    known = set(names)
    for node in definition.nodes:
        for dependency in node.depends_on:
            if dependency not in known:
                issues.append(f"{node.name}：前置节点不存在：{dependency}")
        for direction, model_name in (
            ("输入", node.input_model),
            ("输出", node.output_model),
        ):
            model = getattr(schema_module, model_name, None)
            if model is None:
                issues.append(f"{node.name}：{direction}模型不存在：{model_name}")
                continue
            for field_name, field in model.model_fields.items():
                if not field.description or not any("\u4e00" <= char <= "\u9fff" for char in field.description):
                    issues.append(f"{model_name}.{field_name}：缺少中文业务说明")
                if not field.examples:
                    issues.append(f"{model_name}.{field_name}：缺少示例")
    if _has_cycle(definition.nodes):
        issues.append("工作流依赖存在环")
    return {"valid": not issues, "issues": issues, "node_count": len(definition.nodes)}
```

Use a local depth-first traversal in `_has_cycle`; do not import filesystem or infrastructure modules.

- [ ] **Step 5: Run tests and verify GREEN**

Run the focused tests, then:

```bash
UV_CACHE_DIR=/tmp/uv-cache uv run --python 3.11 python -m pytest demo/tests/test_contracts.py demo/tests/test_workflow_contracts.py -q
```

Expected: all tests pass.

### Task 2: Move narrative and candidate-report decisions into pure policy

**Files:**
- Create: `demo/domain/narrative_policy.py`
- Create: `demo/tests/test_narrative_policy.py`
- Modify: `demo/pipeline.py`

- [ ] **Step 1: Write failing policy tests**

```python
def test_company_profile_is_always_allowed_and_only_selected_modules_are_added():
    allowed = select_narrative_fields(
        {"company_profile_section", "industry_overview", "main_products"},
        ["main_products"],
    )
    assert allowed == {"company_profile_section", "main_products"}


def test_candidate_report_requires_at_least_one_completed_review():
    assert should_create_candidate_report({}) is False
    assert should_create_candidate_report(
        {"format": {"status": "completed", "findings": []}}
    ) is True
```

- [ ] **Step 2: Run tests and verify RED**

Run:

```bash
UV_CACHE_DIR=/tmp/uv-cache uv run --python 3.11 python -m pytest demo/tests/test_narrative_policy.py -q
```

Expected: import fails because `demo.domain.narrative_policy` does not exist.

- [ ] **Step 3: Implement the minimal pure policy**

```python
def select_narrative_fields(
    routed_fields: set[str],
    selected_modules: list[str],
) -> set[str]:
    selected = set(selected_modules)
    return {
        field_key
        for field_key in routed_fields
        if field_key == "company_profile_section" or field_key in selected
    }


def should_create_candidate_report(reviews: Mapping[str, Mapping[str, Any]]) -> bool:
    return any(review.get("status") in {"completed", "completed_with_issues"} for review in reviews.values())
```

- [ ] **Step 4: Replace the equivalent inline branches in `demo/pipeline.py`**

Use:

```python
llm_allowed = select_narrative_fields(
    fields_for_route(routes, RouteKind.BAILIAN_GLM),
    selected_modules,
)
```

and:

```python
if should_create_candidate_report(reviews):
    final_report = output_dir / "资产评估报告_最终候选.docx"
    shutil.copy2(report, final_report)
```

- [ ] **Step 5: Run focused and pipeline tests**

Run:

```bash
UV_CACHE_DIR=/tmp/uv-cache uv run --python 3.11 python -m pytest demo/tests/test_narrative_policy.py demo/tests/test_pipeline.py -q
```

Expected: all tests pass.

### Task 3: Add typed runtime trace records

**Files:**
- Modify: `demo/schemas.py`
- Create: `demo/adapters/workflow_trace.py`
- Create: `demo/tests/test_workflow_trace.py`

- [ ] **Step 1: Write failing recorder tests**

```python
def test_trace_recorder_validates_models_and_serializes_versions(tmp_path):
    recorder = WorkflowTraceRecorder(
        workflow_version="1.1.0",
        contract_version="workflow_contract.v1",
        versions={"prompt": "yellow_narratives.v1"},
    )
    recorder.record(
        node_name="inventory",
        input_model=InventoryInput,
        output_model=InventoryOutput,
        input_payload={"template_path": "template.docx"},
        output_payload={"locations": []},
        status="completed",
        evidence=[],
        issues=[],
        human_checkpoint=None,
    )
    path = recorder.export(tmp_path / "workflow_trace.json")
    payload = json.loads(path.read_text(encoding="utf-8"))
    assert payload["nodes"][0]["input_model"] == "InventoryInput"
    assert payload["versions"]["prompt"] == "yellow_narratives.v1"


def test_trace_recorder_rejects_invalid_output():
    recorder = WorkflowTraceRecorder("1.1.0", "workflow_contract.v1", {})
    with pytest.raises(ValidationError):
        recorder.record(
            node_name="inventory",
            input_model=InventoryInput,
            output_model=InventoryOutput,
            input_payload={"template_path": "template.docx"},
            output_payload={},
            status="completed",
            evidence=[],
            issues=[],
            human_checkpoint=None,
        )
```

- [ ] **Step 2: Run tests and verify RED**

Run:

```bash
UV_CACHE_DIR=/tmp/uv-cache uv run --python 3.11 python -m pytest demo/tests/test_workflow_trace.py -q
```

Expected: import fails because the trace recorder does not exist.

- [ ] **Step 3: Add trace schemas**

Add to `demo/schemas.py`:

```python
class NodeTrace(DemoModel):
    node_name: str = Field(description="实际执行的节点名称", examples=["fill_word"])
    status: str = Field(description="节点执行状态", examples=["completed"])
    started_at: str = Field(description="节点开始时间", examples=["2026-07-26T12:00:00+08:00"])
    finished_at: str = Field(description="节点结束时间", examples=["2026-07-26T12:00:01+08:00"])
    input_model: str = Field(description="节点输入模型名称", examples=["FillWordInput"])
    output_model: str = Field(description="节点输出模型名称", examples=["FillWordOutput"])
    input_data: dict[str, Any] = Field(description="已校验的节点输入摘要", examples=[{}])
    output_data: dict[str, Any] = Field(description="已校验的节点输出摘要", examples=[{}])
    evidence: list[dict[str, Any]] = Field(description="节点使用的来源证据", examples=[[]])
    issues: list[str] = Field(description="节点问题列表", examples=[[]])
    human_checkpoint: str | None = Field(
        default=None,
        description="节点人工确认要求",
        examples=["评估师审核问题清单"],
    )


class WorkflowTrace(DemoModel):
    workflow_version: str = Field(description="工作流业务版本", examples=["1.1.0"])
    contract_version: str = Field(description="节点契约规则版本", examples=["workflow_contract.v1"])
    versions: dict[str, str] = Field(description="规则、Prompt、模型和数据版本", examples=[{"prompt": "yellow_narratives.v1"}])
    nodes: list[NodeTrace] = Field(description="实际执行节点轨迹", examples=[[]])
```

- [ ] **Step 4: Implement the trace adapter**

`WorkflowTraceRecorder.record` must:

1. call `input_model.model_validate(input_payload)`;
2. call `output_model.model_validate(output_payload)`;
3. store `model_dump(mode="json")` values;
4. store timezone-aware ISO timestamps;
5. append a `NodeTrace`;
6. return the validated output model.

`export` must validate the complete `WorkflowTrace` and call the existing `write_json`.

- [ ] **Step 5: Run trace and schema tests**

Run:

```bash
UV_CACHE_DIR=/tmp/uv-cache uv run --python 3.11 python -m pytest demo/tests/test_workflow_trace.py demo/tests/test_contracts.py -q
```

Expected: all tests pass.

### Task 4: Integrate contract validation and trace export into the pipeline

**Files:**
- Modify: `demo/pipeline.py`
- Modify: `demo/api_server.py`
- Modify: `demo/tests/test_pipeline.py`

- [ ] **Step 1: Add a failing end-to-end trace assertion**

Extend the existing pipeline fixture test:

```python
trace_path = tmp_path / "workflow_trace.json"
assert trace_path.exists()
trace = json.loads(trace_path.read_text(encoding="utf-8"))
assert trace["contract_version"] == "workflow_contract.v1"
assert [node["node_name"] for node in trace["nodes"]] == [
    "inventory",
    "ocr_pdf",
    "export_ocr_workbook",
    "extract_sources",
    "resolve_fields",
    "select_narrative_modules",
    "generate_narrative",
    "fill_word",
    "llm_format_review",
    "llm_data_validation",
    "llm_semantic_review",
    "review_aggregate",
    "export_audit",
]
assert str(trace_path) in manifest["outputs"]
```

Add a separate failing test which passes an invalid workflow definition through a configurable `workflow_path` and asserts that report generation stops before OCR.

- [ ] **Step 2: Run the focused pipeline tests and verify RED**

Run:

```bash
UV_CACHE_DIR=/tmp/uv-cache uv run --python 3.11 python -m pytest demo/tests/test_pipeline.py -q
```

Expected: `workflow_trace.json` is missing.

- [ ] **Step 3: Validate the workflow before external work**

Add an optional `workflow_path: Path | None = None` argument to `run_pipeline`. Load the supplied path or `demo/workflow.yaml`, call `validate_workflow_contract`, and raise:

```python
raise ValueError("工作流契约校验失败：" + "；".join(contract_result["issues"]))
```

This must occur before `run_project`, OCR, API, LLM, or Word calls.

- [ ] **Step 4: Record the actual node outputs**

Create one recorder at pipeline start. At each existing stage, record the smallest contract-valid JSON summary. Examples:

```python
recorder.record(
    node_name="inventory",
    input_model=schemas.InventoryInput,
    output_model=schemas.InventoryOutput,
    input_payload={"template_path": str(template)},
    output_payload={"locations": template_inventory},
    status="completed",
    evidence=[],
    issues=[],
    human_checkpoint=node_map["inventory"].human_checkpoint,
)
```

For disabled review adapters, record `skipped` with an output payload using:

```python
{
    "review_type": review_name,
    "status": "skipped",
    "summary": "未启用对应 LLM 审核",
    "findings": [],
    "model": "",
    "prompt_version": review_prompt_versions[review_name],
}
```

Do not include raw PDF text or secrets in the trace. Evidence entries retain only source kind, file name, locator, and version.

- [ ] **Step 5: Export trace and expose it through the API**

Export before `run_manifest.json`, append its path to manifest outputs, and add:

```python
"workflow_trace.json": "工作流节点轨迹",
```

to `demo/api_server.py::_artifact_list`.

- [ ] **Step 6: Run pipeline tests and verify GREEN**

Run:

```bash
UV_CACHE_DIR=/tmp/uv-cache uv run --python 3.11 python -m pytest demo/tests/test_pipeline.py demo/tests/test_delivery_files.py -q
```

Expected: all tests pass and the trace artifact is downloadable.

### Task 5: Add representative new-rule fixtures

**Files:**
- Create: `demo/fixtures/workflow_cases.yaml`
- Create: `demo/expected/workflow_cases.yaml`
- Modify: `demo/tests/test_domain_examples.py`

- [ ] **Step 1: Add a failing parameterized fixture test**

```python
def test_workflow_policy_examples():
    cases = json.loads(Path("demo/fixtures/workflow_cases.yaml").read_text(encoding="utf-8"))
    expected = json.loads(Path("demo/expected/workflow_cases.yaml").read_text(encoding="utf-8"))
    assert len(cases) >= 10
    actual = {}
    for case in cases:
        if case["op"] == "methods":
            actual[case["id"]] = normalize_valuation_methods(case["value"])
        elif case["op"] == "modules":
            actual[case["id"]] = normalize_narrative_modules(case["value"])
        elif case["op"] == "narrative_fields":
            actual[case["id"]] = sorted(select_narrative_fields(set(case["routed"]), case["selected"]))
        elif case["op"] == "candidate":
            actual[case["id"]] = should_create_candidate_report(case["reviews"])
        elif case["op"] == "review_aggregate":
            actual[case["id"]] = aggregate_reviews(case["reviews"])
    assert actual == expected
```

- [ ] **Step 2: Run and verify RED**

Run:

```bash
UV_CACHE_DIR=/tmp/uv-cache uv run --python 3.11 python -m pytest demo/tests/test_domain_examples.py::test_workflow_policy_examples -q
```

Expected: failure because the fixture files do not exist.

- [ ] **Step 3: Add at least 12 deterministic cases**

Use the 12 cases specified in the design:

- valid method combinations;
- invalid/empty method handled in a dedicated `pytest.raises` test;
- all modules;
- one module;
- empty module handled in a dedicated `pytest.raises` test;
- field whitelist filtering;
- no reviews;
- completed review;
- failed review aggregate;
- model override resolution;
- QCC identity mismatch covered by its existing test;
- OCR cache behavior covered by its existing pipeline test.

The fixture file must use only invented company names and values.

- [ ] **Step 4: Run the example tests and verify GREEN**

Run:

```bash
UV_CACHE_DIR=/tmp/uv-cache uv run --python 3.11 python -m pytest demo/tests/test_domain_examples.py demo/tests/test_manual_inputs.py demo/tests/test_llm_config.py -q
```

Expected: all tests pass.

### Task 6: Register versions and document the handoff

**Files:**
- Modify: `demo/data_manifest.yaml`
- Modify: `demo/README.md`
- Modify: `README.md`
- Modify: `demo/CHANGELOG.md`
- Test: `demo/tests/test_contracts.py`

- [ ] **Step 1: Add a failing version-registration test**

```python
def test_manifest_registers_contract_trace_and_review_schema_versions():
    manifest = json.loads(Path("demo/data_manifest.yaml").read_text(encoding="utf-8"))
    versions = manifest["rule_versions"]
    assert versions["workflow_contract"] == "workflow_contract.v1"
    assert versions["workflow_trace"] == "workflow_trace.v1"
    assert versions["narrative_policy"] == "narrative_policy.v1"
    assert versions["review_output_schema"] == "review_output.v1"
```

- [ ] **Step 2: Run and verify RED**

Run:

```bash
UV_CACHE_DIR=/tmp/uv-cache uv run --python 3.11 python -m pytest demo/tests/test_contracts.py::test_manifest_registers_contract_trace_and_review_schema_versions -q
```

Expected: missing key assertion.

- [ ] **Step 3: Add manifest versions and documentation**

Register the four versions in `data_manifest.yaml`. Document:

- what `workflow_trace.json` contains;
- which failures terminate generation;
- which provider failures degrade to blank plus issue;
- that candidate Word does not silently apply LLM changes;
- that c2m remains responsible for users, permissions, async jobs, storage, retries, monitoring, and release.

Add a dated `0.6.0` changelog entry covering code, prompts, fixtures, outputs, and migration behavior.

- [ ] **Step 4: Run contract tests and verify GREEN**

Run:

```bash
UV_CACHE_DIR=/tmp/uv-cache uv run --python 3.11 python -m pytest demo/tests/test_contracts.py -q
```

Expected: all tests pass.

### Task 7: Full verification and delivery

**Files:**
- Verify all modified files.

- [ ] **Step 1: Run the complete Python suite**

```bash
UV_CACHE_DIR=/tmp/uv-cache uv run --python 3.11 python -m pytest -q
```

Expected: all tests pass.

- [ ] **Step 2: Compile Python sources**

```bash
UV_CACHE_DIR=/tmp/uv-cache uv run --python 3.11 python -m compileall -q demo
```

Expected: exit code 0.

- [ ] **Step 3: Build the frontend**

```bash
npm run build
```

Run from `frontend/`.

Expected: Vite build succeeds; the existing chunk-size advisory is acceptable.

- [ ] **Step 4: Check repository hygiene**

```bash
git diff --check
git status -sb
git ls-files | rg '(^|/)\.env$|runs/|__pycache__|docs/superpowers'
```

Expected: no whitespace errors; no credentials, generated runs, caches, or private superpower documents are tracked.

- [ ] **Step 5: Commit and push**

```bash
git add README.md demo frontend/src
git commit -m "feat: add auditable workflow contracts"
git push origin main
```

Expected: `main` matches `origin/main` and the worktree is clean.
