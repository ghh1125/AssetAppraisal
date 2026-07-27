# 通用语义优先与问题清单可读性 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 阻止未完成评估的全零列进入报告，保证语义 Excel 证据完整，并生成按 Word 页码可直接复核的问题清单。

**Architecture:** 在 Excel 适配器中增加纯判定函数和拒绝原因；业务内核持久化完整字段证据，生产流水线直接复用；问题清单适配器只负责排序、汇总和 Excel 展示，不改变问题领域模型。

**Tech Stack:** Python 3.11、openpyxl、Pydantic、pytest、python-docx/Word OOXML。

---

### Task 1: 拒绝未完成评估的全零列

**Files:**
- Modify: `demo/adapters/semantic_excel.py`
- Test: `demo/tests/test_semantic_excel.py`

- [ ] **Step 1: 写失败测试**

创建账面净资产为正、评估列全零的工作簿，断言没有 `asset_approach_value`，并包含 `incomplete_appraisal_column` 问题；再创建评估列存在其他非零值的工作簿，断言净资产零值被保留。

- [ ] **Step 2: 验证测试按预期失败**

Run: `PYTHONPATH=. /Users/ghh/Documents/Code/mcpify/AssetAppraisal/.venv/bin/pytest demo/tests/test_semantic_excel.py -k "incomplete_appraisal or supported_zero" -q`

Expected: 第一项错误地得到 `0.0`，测试失败。

- [ ] **Step 3: 实现最小判定**

在 `_net_asset_candidate` 中检查评估列是否存在非零金额，将结果写入候选的 `appraised_incomplete`；`extract_workbook_facts` 对无效零值跳过字段并记录带来源位置的问题。

- [ ] **Step 4: 验证通过**

Run: `PYTHONPATH=. /Users/ghh/Documents/Code/mcpify/AssetAppraisal/.venv/bin/pytest demo/tests/test_semantic_excel.py -q`

Expected: 全部通过。

- [ ] **Step 5: 提交**

Commit: `fix: reject incomplete zero appraisal columns`

### Task 2: 语义结果优先并完整持久化证据

**Files:**
- Modify: `demo/run.py`
- Modify: `demo/pipeline.py`
- Test: `demo/tests/test_demo_run.py`
- Test: `demo/tests/test_pipeline.py`

- [ ] **Step 1: 写失败测试**

构造固定坐标有错误值、语义表头有正确值的工作簿，断言采用语义值；运行管线后断言历史表、资产范围表和长期资产表的来源文件及单元格不为空。

- [ ] **Step 2: 验证测试按预期失败**

Run: `PYTHONPATH=. /Users/ghh/Documents/Code/mcpify/AssetAppraisal/.venv/bin/pytest demo/tests/test_demo_run.py demo/tests/test_pipeline.py -k "semantic_primary or table_evidence" -q`

Expected: 固定表覆盖语义表，或证据显示 `unknown/missing`。

- [ ] **Step 3: 实现证据文件和优先规则**

`run_project` 输出 `normalized_evidence.json`；语义解析字段覆盖固定坐标字段，历史表继续使用兼容合并；`run_pipeline` 读取证据文件，并且只有语义表不存在时才读取固定资产范围表。

- [ ] **Step 4: 验证通过**

Run: `PYTHONPATH=. /Users/ghh/Documents/Code/mcpify/AssetAppraisal/.venv/bin/pytest demo/tests/test_demo_run.py demo/tests/test_pipeline.py -k "semantic_primary or table_evidence" -q`

Expected: 全部通过。

- [ ] **Step 5: 提交**

Commit: `feat: make semantic excel evidence primary`

### Task 3: 重做问题清单

**Files:**
- Modify: `demo/domain/generation_issues.py`
- Modify: `demo/adapters/generation_issues.py`
- Test: `demo/tests/test_generation_issues.py`

- [ ] **Step 1: 写失败测试**

断言问题按实际页码排序，无页码排最后；每项产生“第N页｜位置类型｜位置描述”；工作簿包含“检查总览”和“问题明细”，总览显示计数，明细前12列为业务检查列。

- [ ] **Step 2: 验证测试按预期失败**

Run: `PYTHONPATH=. /Users/ghh/Documents/Code/mcpify/AssetAppraisal/.venv/bin/pytest demo/tests/test_generation_issues.py -q`

Expected: 缺少新工作表和检查位置，测试失败。

- [ ] **Step 3: 实现排序和展示**

领域函数增加稳定排序、检查顺序和检查位置；Excel 适配器生成总览、明细，设置冻结窗格、筛选、列宽、边框、交替底色和优先级颜色。

- [ ] **Step 4: 验证通过并目视检查**

Run: `PYTHONPATH=. /Users/ghh/Documents/Code/mcpify/AssetAppraisal/.venv/bin/pytest demo/tests/test_generation_issues.py -q`

Expected: 全部通过。随后打开真实生成的问题清单，检查总览和明细可读性。

- [ ] **Step 5: 提交**

Commit: `feat: make generation issues reviewer friendly`

### Task 4: 真实资料回归和交付

**Files:**
- Modify: `demo/CHANGELOG.md`
- Test: `demo/tests/`

- [ ] **Step 1: 全量回归**

Run: `PYTHONPATH=. /Users/ghh/Documents/Code/mcpify/AssetAppraisal/.venv/bin/pytest demo/tests -q`

Expected: 100% 通过。

- [ ] **Step 2: 四组资料重新生成**

使用亦盛、夏弗纳、铁投能源、浙江晶引现有 Excel 和 OCR 缓存运行，不重新执行 OCR。断言浙江晶引资产基础法零值保留 `XXX`，其余已验证核心金额不变。

- [ ] **Step 3: 检查 Word 和问题清单**

渲染四份 Word，重点检查评估结论、历史表和长期资产页；打开四份问题清单，确认按页码排序、总览计数正确、来源证据可读。

- [ ] **Step 4: 更新变更记录**

在 `demo/CHANGELOG.md` 记录零值门禁、语义优先、证据持久化和问题清单新版。

- [ ] **Step 5: 提交并推送**

Commit: `feat: harden semantic excel review workflow`
