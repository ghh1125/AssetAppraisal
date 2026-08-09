# Qichacha 699 Default Endpoint and Narrative Prompt Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 默认调用企查查 699 官方接口并将其结构化证据交给六模块 LLM，以自然、专业且不虚构事实的方式生成报告内容。

**Architecture:** 在现有 `QichachaApiAdapter` 内注册 699 默认路径，针对该接口单独构造 `stockCode` 参数并解析嵌套返回；LLM 仍消费统一 evidence 契约，只调整 Prompt 的分析和写作规则。用户配置入口保持为统一 Key/SecretKey。

**Tech Stack:** Python 3.11、pytest、requests-compatible HTTP client、百炼 OpenAI-compatible API、YAML/Markdown 文档。

---

### Task 1: 699 默认接口与请求参数

**Files:**
- Modify: `demo/tests/test_adapters.py`
- Modify: `demo/adapters/company_api.py`

- [ ] **Step 1: Write the failing test**

新增测试断言 `DEFAULT_ENDPOINTS["699"] == "/IPO/GetIPODetail"`，并断言请求 Query 使用 `stockCode=600001`、不使用 `searchKey`。

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run --python 3.11 pytest demo/tests/test_adapters.py -k "699_default" -q`

Expected: FAIL because 699 is not registered by default.

- [ ] **Step 3: Write minimal implementation**

在 `DEFAULT_ENDPOINTS` 注册官方路径；在 `_get()` 中为 699 构造 `stockCode` 参数；对标候选调用时传入 915 返回的股票代码。

- [ ] **Step 4: Run test to verify it passes**

Run: `uv run --python 3.11 pytest demo/tests/test_adapters.py -k "699_default" -q`

Expected: PASS.

### Task 2: 699 嵌套响应解析

**Files:**
- Modify: `demo/tests/test_adapters.py`
- Modify: `demo/adapters/company_api.py`

- [ ] **Step 1: Write the failing test**

使用官方 `Result.BaseInfo` / `Result.IPOPublishInfo` 形状，断言提取企业名称、A 股代码、行业、证券类别、上市日期、市净率、市盈率、成立日期和发行方式，同时拒绝与 915 候选不一致的企业或股票代码。

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run --python 3.11 pytest demo/tests/test_adapters.py -k "699_nested" -q`

Expected: FAIL because the current parser reads only flat records.

- [ ] **Step 3: Write minimal implementation**

让 `_listed_company_detail_text()` 显式读取 `Result.BaseInfo` 和 `Result.IPOPublishInfo`，执行公司名称及 A 股代码校验后格式化明确返回字段。

- [ ] **Step 4: Run test to verify it passes**

Run: `uv run --python 3.11 pytest demo/tests/test_adapters.py -k "699_nested" -q`

Expected: PASS.

### Task 3: 六模块专业写作 Prompt 与配置说明

**Files:**
- Modify: `demo/tests/test_adapters.py`
- Modify: `demo/prompts/yellow_narratives.v3.txt`
- Modify: `.env.example`
- Modify: `README.md`
- Modify: `demo/README.md`
- Modify: `demo/data_manifest.yaml`
- Modify: `demo/CHANGELOG.md`

- [ ] **Step 1: Write the failing test**

断言 Prompt 包含“专业综合”“审慎分析”“事实不得虚构”等规则，并且环境示例不要求 `QICHACHA_ENDPOINT_699`。

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run --python 3.11 pytest demo/tests/test_adapters.py demo/tests/test_delivery_files.py -k "narrative or 699" -q`

Expected: FAIL because the current Prompt is mechanical and `.env.example` still exposes 699 endpoint.

- [ ] **Step 3: Write minimal implementation**

重写 v3 Prompt 为证据锚定的报告写作规则；更新说明为 699 默认启用、用户只需统一 Key/SecretKey；保留内部环境覆盖读取以便调试。

- [ ] **Step 4: Run related regression tests**

Run: `uv run --python 3.11 pytest demo/tests/test_adapters.py demo/tests/test_bailian_glm.py demo/tests/test_delivery_files.py demo/tests/test_four_node_workflow.py -q`

Expected: PASS.

- [ ] **Step 5: Verify patch quality**

Run: `git diff --check`

Expected: no output.
