import test from 'node:test'
import assert from 'node:assert/strict'

import { buildAssetAppraisalForm } from './asset-appraisal.js'

test('run request uploads confirmed audit, workbook, and source material slots', () => {
  const form = buildAssetAppraisalForm({
    auditMaterials: [new Blob(['pdf']), new Blob(['audit-doc'])],
    reportingWorkbook: new Blob(['asset']),
    incomeWorkbook: new Blob(['income']),
    registryMaterials: [new Blob(['registry'])],
    ownershipHistoryMaterials: [new Blob(['ownership'])],
    unrecordedIntangiblesMaterials: [new Blob(['intangible'])],
    companyProfileMaterials: [new Blob(['profile'])],
    inputs: { target_company_name: '示例公司' },
    useGlm: true,
    useQichacha: false,
    reuseOcr: true,
  })

  assert.deepEqual([...form.keys()], [
    'audit_materials',
    'audit_materials',
    'reporting_workbook',
    'income_workbook',
    'registry_materials',
    'ownership_history_materials',
    'unrecorded_intangibles_materials',
    'company_profile_materials',
    'inputs',
    'use_glm',
    'use_qichacha',
    'reuse_ocr',
  ])
})
