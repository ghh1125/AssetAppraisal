import test from 'node:test'
import assert from 'node:assert/strict'

import { buildAssetAppraisalForm } from './asset-appraisal.js'

test('run request uploads the generic related-materials list', () => {
  const form = buildAssetAppraisalForm({
    materials: [new Blob(['pdf']), new Blob(['workbook'])],
    pdf: new Blob(['pdf']),
    reportingWorkbook: new Blob(['asset']),
    incomeWorkbook: new Blob(['income']),
    referenceReport: new Blob(['must not upload']),
    auditedFinancials: new Blob(['must not upload']),
    inputs: { target_company_name: '示例公司' },
    useGlm: true,
    useQichacha: false,
    reuseOcr: true,
  })

  assert.deepEqual([...form.keys()], [
    'materials',
    'materials',
    'pdf',
    'reporting_workbook',
    'income_workbook',
    'inputs',
    'use_glm',
    'use_qichacha',
    'reuse_ocr',
  ])
})
