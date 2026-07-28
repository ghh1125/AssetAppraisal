# Three-Upload Frontend Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the asset-appraisal page expose only the audit PDF, asset-basis/reporting workbook, and income/market workbook.

**Architecture:** Keep the backend upload contract backward compatible, while reducing the browser-side file state and request payload to three supported inputs. Define the visible upload slots as a small pure-data module so the exact UI contract can be tested without adding a component-test dependency.

**Tech Stack:** Vue 3, Vite, Node built-in test runner, browser `FormData`.

---

### Task 1: Lock the three-upload contract with failing tests

**Files:**
- Create: `frontend/src/domain/upload-fields.js`
- Create: `frontend/src/domain/upload-fields.test.js`
- Create: `frontend/src/api/asset-appraisal.test.js`
- Modify: `frontend/package.json`

- [ ] **Step 1: Write the upload-slot test before the module exists**

```js
import assert from 'node:assert/strict'
import test from 'node:test'

import { uploadFields } from './upload-fields.js'

test('the page exposes exactly one PDF and two Excel uploads', () => {
  assert.deepEqual(uploadFields.map(({ key, accept }) => [key, accept]), [
    ['pdf', '.pdf'],
    ['reportingWorkbook', '.xlsx,.xlsm'],
    ['incomeWorkbook', '.xlsx,.xlsm'],
  ])
})
```

- [ ] **Step 2: Write the request-payload test before the builder exists**

```js
import assert from 'node:assert/strict'
import test from 'node:test'

import { buildAssetAppraisalForm } from './asset-appraisal.js'

test('the browser request sends only the three supported files', () => {
  const form = buildAssetAppraisalForm({
    pdf: new Blob(['pdf']),
    reportingWorkbook: new Blob(['asset']),
    incomeWorkbook: new Blob(['income']),
    referenceReport: new Blob(['reference']),
    auditedFinancials: new Blob(['audit']),
    inputs: {},
    useGlm: true,
    useQichacha: true,
    reuseOcr: true,
  })
  assert.deepEqual(
    [...form.keys()].filter((key) => key !== 'inputs' && !key.startsWith('use_')),
    ['pdf', 'income_workbook', 'reporting_workbook'],
  )
})
```

- [ ] **Step 3: Add both tests to the frontend test command**

```json
"test": "node --test src/**/*.test.js"
```

- [ ] **Step 4: Run the tests and confirm RED**

Run: `npm test`

Expected: FAIL because `upload-fields.js` and `buildAssetAppraisalForm` do not exist.

### Task 2: Implement the three-upload browser contract

**Files:**
- Create: `frontend/src/domain/upload-fields.js`
- Modify: `frontend/src/api/asset-appraisal.js`
- Modify: `frontend/src/views/asset-appraisal/index.vue`

- [ ] **Step 1: Define the three visible upload slots**

```js
export const uploadFields = Object.freeze([
  { key: 'pdf', accept: '.pdf', icon: 'PDF', titleKey: 'pdfTitle' },
  {
    key: 'reportingWorkbook',
    accept: '.xlsx,.xlsm',
    icon: 'XLSX',
    titleKey: 'reportingWorkbookTitle',
  },
  {
    key: 'incomeWorkbook',
    accept: '.xlsx,.xlsm',
    icon: 'XLSX',
    titleKey: 'incomeWorkbookTitle',
  },
])
```

- [ ] **Step 2: Extract and restrict request construction**

```js
export function buildAssetAppraisalForm({
  pdf,
  incomeWorkbook,
  reportingWorkbook,
  inputs,
  useGlm,
  useQichacha,
  reuseOcr,
}) {
  const form = new FormData()
  if (pdf) form.append('pdf', pdf)
  if (incomeWorkbook) form.append('income_workbook', incomeWorkbook)
  if (reportingWorkbook) form.append('reporting_workbook', reportingWorkbook)
  form.append('inputs', JSON.stringify(inputs))
  form.append('use_glm', String(useGlm))
  form.append('use_qichacha', String(useQichacha))
  form.append('reuse_ocr', String(reuseOcr ?? true))
  return form
}
```

- [ ] **Step 3: Drive the page from `uploadFields`**

Use `uploadFields` to initialize `files` and render the three drag-and-drop upload controls. Remove `referenceReport` and `auditedFinancials` from page state and from `submit()`.

- [ ] **Step 4: Run focused tests and confirm GREEN**

Run: `npm test`

Expected: all frontend unit tests pass.

### Task 3: Align user-facing copy and repository documentation

**Files:**
- Modify: `frontend/src/i18n/index.js`
- Modify: `README.md`
- Modify: `demo/README.md`
- Modify: `demo/CHANGELOG.md`

- [ ] **Step 1: Update Chinese and English upload copy**

State that the standard input is one optional audit PDF plus the available asset-basis/reporting and income/market workbooks. Remove the reference-report and audited-financial upload labels.

- [ ] **Step 2: Update README material lists**

Document exactly three browser upload roles and state that the backend template and OCR workbook are not user uploads.

- [ ] **Step 3: Record the behavior change**

Add a changelog item explaining that two unintended browser upload controls were removed while backend compatibility remains.

- [ ] **Step 4: Check for stale UI references**

Run:

```bash
rg -n "referenceReport|auditedFinancials|reference_report|audited_financials|参考评估报告 DOCX|审计财务 XLSX" frontend/src README.md demo/README.md
```

Expected: no browser-facing stale references; backend compatibility references may remain outside frontend source.

### Task 4: Verify and publish

**Files:**
- Verify all modified frontend and documentation files.

- [ ] **Step 1: Run frontend tests**

Run: `npm test`

Expected: all tests pass.

- [ ] **Step 2: Build the frontend**

Run: `npm run build`

Expected: Vite build exits successfully.

- [ ] **Step 3: Run backend API compatibility tests**

Run:

```bash
PYTHONPATH=. .venv/bin/python -m pytest demo/tests/test_api_server.py -q
```

Expected: backend still accepts legacy optional upload fields.

- [ ] **Step 4: Commit and push**

Commit the implementation and push `main` after all verification commands pass.
