# Aliyun Document Mind OCR Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace local PaddleOCR as the default PDF OCR path with an Aliyun Document Mind adapter while preserving the existing OCR workbook, cache, Word, audit, and issue-list contracts.

**Architecture:** A focused adapter converts Aliyun `layouts` and table `cells` into the existing SDK-independent page contract. A provider factory lazily constructs either Aliyun, local PaddleOCR, no OCR, or an unavailable adapter; Web and CLI entry points inject the selected adapter without importing cloud or Paddle SDKs into `domain/`.

**Tech Stack:** Python 3.11, Alibaba Cloud Document Mind Python SDK, Pydantic-free adapter dictionaries, pytest, existing FastAPI/background pipeline.

---

### Task 1: Normalize Aliyun layouts into the existing OCR contract

**Files:**
- Create: `demo/adapters/aliyun_docmind_ocr.py`
- Create: `demo/tests/test_aliyun_docmind_ocr.py`

- [ ] **Step 1: Write failing conversion tests**

Create tests with a two-page payload containing a text layout and a table layout. Assert one-based page numbers, original coordinates and confidence, and table cells restored from `ysc/yec/xsc/xec`:

```python
def test_layouts_to_pages_preserves_text_table_and_evidence():
    layouts = [
        {
            "type": "text",
            "text": "资产总计",
            "pageNum": 0,
            "layoutConf": 0.98,
            "pos": [{"x": 10, "y": 20}, {"x": 90, "y": 40}],
            "uniqueId": "text-1",
        },
        {
            "type": "table",
            "pageNum": 1,
            "uniqueId": "table-1",
            "cells": [
                {
                    "ysc": 0,
                    "yec": 0,
                    "xsc": 0,
                    "xec": 1,
                    "layouts": [{"text": "项目"}],
                },
                {
                    "ysc": 1,
                    "yec": 1,
                    "xsc": 0,
                    "xec": 0,
                    "layouts": [{"text": "货币资金"}],
                },
            ],
        },
    ]
    pages = layouts_to_pages(layouts)
    assert pages[0]["page_number"] == 1
    assert pages[0]["blocks"][0]["bbox"] == [10.0, 20.0, 90.0, 40.0]
    assert pages[1]["page_number"] == 2
    assert pages[1]["tables"][0]["cells"][0]["row"] == 1
    assert pages[1]["tables"][0]["cells"][0]["column_span"] == 2
```

Add a boundary test asserting that missing confidence remains `None`, missing text falls back to cleaned `markdownContent`, and empty layouts produce an empty list.

- [ ] **Step 2: Run the conversion tests and verify RED**

Run:

```bash
/Users/ghh/Documents/Code/mcpify/AssetAppraisal/.venv/bin/pytest demo/tests/test_aliyun_docmind_ocr.py -q
```

Expected: import failure because `demo.adapters.aliyun_docmind_ocr` does not exist.

- [ ] **Step 3: Implement pure conversion helpers**

Implement:

```python
def layouts_to_pages(layouts: list[dict[str, Any]]) -> list[dict[str, Any]]:
    ...

def _bbox(pos: Any) -> list[float]:
    ...

def _table_cells(layout: dict[str, Any]) -> list[dict[str, Any]]:
    ...
```

The output must use `page_number`, `page_count`, `blocks`, and `tables`; table cells must include `row`, `column`, `row_span`, `column_span`, `text`, `confidence`, and `bbox`.

- [ ] **Step 4: Run the conversion tests and verify GREEN**

Run the same pytest command. Expected: all conversion tests pass.

- [ ] **Step 5: Commit Task 1**

```bash
git add demo/adapters/aliyun_docmind_ocr.py demo/tests/test_aliyun_docmind_ocr.py
git commit -m "feat: normalize aliyun document layouts"
```

### Task 2: Implement asynchronous submission, polling, pagination, and safe errors

**Files:**
- Modify: `demo/adapters/aliyun_docmind_ocr.py`
- Modify: `demo/tests/test_aliyun_docmind_ocr.py`

- [ ] **Step 1: Add failing orchestration tests**

Use a fake client implementing `submit`, `status`, and `result`. Assert:

```python
def test_adapter_submits_polls_and_collects_all_layout_pages(tmp_path):
    client = FakeClient(
        statuses=[{"Status": "Processing"}, {"Status": "success"}],
        result_pages=[
            {"layouts": [{"type": "text", "text": "第一页", "pageNum": 0}]},
            {"layouts": []},
        ],
    )
    adapter = AliyunDocMindOcrAdapter(
        client,
        poll_interval_seconds=0,
        timeout_seconds=60,
        sleep=lambda _: None,
    )
    pages, issues = adapter.extract(tmp_path / "audit.pdf")
    assert issues == []
    assert pages[0]["blocks"][0]["text"] == "第一页"
    assert client.result_offsets == [0, 3000]
```

Add tests for failed status, timeout, SDK exception with `code/message`, and successful empty output. Assert errors contain the stage and task ID but never contain an AccessKey value.

- [ ] **Step 2: Run targeted tests and verify RED**

Expected: `AliyunDocMindOcrAdapter` is missing.

- [ ] **Step 3: Implement the adapter protocol**

Implement:

```python
class AliyunDocMindOcrAdapter:
    def __init__(
        self,
        client: Any,
        *,
        vlm: bool = False,
        poll_interval_seconds: float = 5,
        timeout_seconds: float = 900,
        layout_step_size: int = 3000,
        sleep: Callable[[float], None] = time.sleep,
        clock: Callable[[], float] = time.monotonic,
    ): ...

    def extract(self, pdf_path: Path) -> tuple[list[dict[str, Any]], list[str]]:
        ...
```

The adapter submits once, polls until `success`/`fail`, paginates until a page returns fewer layouts than `layout_step_size`, and converts all layouts using Task 1 helpers. Catch exceptions and return a single redacted issue string instead of raising.

- [ ] **Step 4: Run targeted tests and verify GREEN**

Expected: all Aliyun adapter tests pass.

- [ ] **Step 5: Commit Task 2**

```bash
git add demo/adapters/aliyun_docmind_ocr.py demo/tests/test_aliyun_docmind_ocr.py
git commit -m "feat: add aliyun document mind ocr adapter"
```

### Task 3: Add the lazy SDK client and provider factory

**Files:**
- Create: `demo/adapters/ocr_factory.py`
- Create: `demo/tests/test_ocr_factory.py`
- Modify: `pyproject.toml`
- Modify: `uv.lock`

- [ ] **Step 1: Write failing factory tests**

Test four modes:

```python
def test_aliyun_is_default_and_does_not_load_paddle(monkeypatch):
    adapter = create_ocr_adapter(
        {
            "ALIBABA_CLOUD_ACCESS_KEY_ID": "id",
            "ALIBABA_CLOUD_ACCESS_KEY_SECRET": "secret",
        },
        aliyun_client_factory=lambda **_: FakeAliyunClient(),
        paddle_factory=lambda: pytest.fail("Paddle must not load"),
    )
    assert isinstance(adapter, AliyunDocMindOcrAdapter)

def test_missing_aliyun_credentials_returns_issue_adapter():
    pages, issues = create_ocr_adapter({}).extract(Path("audit.pdf"))
    assert pages == []
    assert "凭证" in issues[0]

def test_paddle_requires_explicit_provider():
    marker = object()
    assert create_ocr_adapter(
        {"APPRAISAL_OCR_PROVIDER": "paddle"},
        paddle_factory=lambda: marker,
    ) is marker

def test_none_provider_skips_ocr():
    assert create_ocr_adapter({"APPRAISAL_OCR_PROVIDER": "none"}) is None
```

- [ ] **Step 2: Run factory tests and verify RED**

Expected: `demo.adapters.ocr_factory` import failure.

- [ ] **Step 3: Implement the SDK wrapper and provider factory**

`AliyunDocMindSdkClient` must lazily import:

```python
from alibabacloud_docmind_api20220711.client import Client
from alibabacloud_docmind_api20220711 import models
from alibabacloud_tea_openapi import models as open_api_models
from alibabacloud_tea_util import models as util_models
```

It must initialize `open_api_models.Config` from explicit credentials, set endpoint `docmind-api.cn-hangzhou.aliyuncs.com`, and expose plain-dictionary `submit`, `status`, and `result` methods. `create_ocr_adapter` defaults to `aliyun`, accepts `paddle` and `none`, reads VLM and timeout options, and never imports Paddle modules unless `paddle` is explicitly selected.

- [ ] **Step 4: Add lightweight SDK dependencies**

Add the official Document Mind SDK, Tea OpenAPI, and Tea Util packages to the `services` extra. Regenerate `uv.lock` without adding Paddle packages to the default install.

- [ ] **Step 5: Run factory and adapter tests**

Expected: all tests pass with fake clients and no real API calls.

- [ ] **Step 6: Commit Task 3**

```bash
git add demo/adapters/ocr_factory.py demo/tests/test_ocr_factory.py pyproject.toml uv.lock
git commit -m "feat: configure remote ocr providers"
```

### Task 4: Inject the provider into Web and CLI without breaking cache behavior

**Files:**
- Modify: `demo/api_server.py`
- Modify: `demo/run.py`
- Modify: `demo/tests/test_api_server.py`
- Modify: `demo/tests/test_pipeline.py`

- [ ] **Step 1: Add failing Web/CLI integration tests**

Assert the Web executor only calls `create_ocr_adapter` when a PDF exists and no SHA-256 cache is found. Assert a cache hit leaves the factory at zero calls. Add a CLI test that an uploaded PDF with default configuration uses the factory result and never constructs local PaddleOCR directly.

- [ ] **Step 2: Run focused tests and verify RED**

Run:

```bash
/Users/ghh/Documents/Code/mcpify/AssetAppraisal/.venv/bin/pytest \
  demo/tests/test_api_server.py demo/tests/test_pipeline.py -q
```

Expected: new integration assertions fail because entry points still instantiate `PPStructureV3` directly.

- [ ] **Step 3: Replace direct Paddle creation**

In `demo/api_server.py`, after the existing cache lookup:

```python
ocr_adapter = None
if pdf_path is not None and ocr_cache is None:
    from .adapters.ocr_factory import create_ocr_adapter
    ocr_adapter = create_ocr_adapter(os.environ)
```

In `demo/run.py`, replace the direct local pipeline construction with the same factory. Keep `run_pipeline(..., ocr_workbook_path=ocr_cache)` unchanged so cached OCR remains the first path.

- [ ] **Step 4: Run focused and full backend tests**

Expected: focused tests and the complete backend suite pass without a real cloud call.

- [ ] **Step 5: Commit Task 4**

```bash
git add demo/api_server.py demo/run.py demo/tests/test_api_server.py demo/tests/test_pipeline.py
git commit -m "feat: use configured ocr provider in web and cli"
```

### Task 5: Document configuration, protect credentials, and validate the complete flow

**Files:**
- Modify: `.env.example`
- Modify: `README.md`
- Modify: `demo/README.md`
- Modify: `demo/data_manifest.yaml`
- Modify: `demo/CHANGELOG.md`
- Test: `demo/tests/test_aliyun_docmind_ocr.py`
- Test: `frontend/src/**/*.test.js`

- [ ] **Step 1: Update configuration examples and operating instructions**

Document that:

- Aliyun is the default OCR provider.
- `ALIBABA_CLOUD_ACCESS_KEY_ID` and `ALIBABA_CLOUD_ACCESS_KEY_SECRET` are RAM credentials with `AliyunDocmindFullAccess`.
- `APPRAISAL_OCR_VLM=false` uses the lower-cost base path.
- `APPRAISAL_OCR_PROVIDER=paddle` explicitly enables the local high-memory fallback.
- Cache hits do not consume API pages.
- Cloud failures keep placeholders and produce issues.

- [ ] **Step 2: Write credentials only to the ignored local `.env`**

Update the main workspace `.env`, not the worktree or any tracked fixture. Verify with `git check-ignore -v .env` and scan staged content for `LTAI` or the Secret value before every commit.

- [ ] **Step 3: Run security and regression checks**

Run:

```bash
git grep -nE 'LTAI|ALIBABA_CLOUD_ACCESS_KEY_SECRET=.{4,}' -- ':!.env.example'
/Users/ghh/Documents/Code/mcpify/AssetAppraisal/.venv/bin/pytest -q
cd frontend && npm test && npm run build
```

Expected: credential scan returns no tracked secrets; backend, frontend, and build pass.

- [ ] **Step 4: Run an optional one-page synthetic live smoke test**

Generate a one-page non-customer PDF containing a small Chinese financial table. Use the local ignored credentials to call Aliyun once, assert at least one text block or table is returned, and save the result under `/tmp`. Do not use files under `资产评估工作流/`.

- [ ] **Step 5: Verify the local process does not load Paddle**

Run the smoke test in a fresh Python process and inspect imported modules or process output. Expected: neither `paddle` nor `paddleocr` is imported in `aliyun` mode.

- [ ] **Step 6: Commit Task 5**

```bash
git add .env.example README.md demo/README.md demo/data_manifest.yaml demo/CHANGELOG.md
git commit -m "docs: explain aliyun cloud ocr setup"
```

- [ ] **Step 7: Merge, push, and report**

Fast-forward the feature branch into `main`, preserve unrelated untracked files, push `main`, and report tests plus whether the optional live smoke test was run.
