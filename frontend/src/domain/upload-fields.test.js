import test from 'node:test'
import assert from 'node:assert/strict'

import { createUploadState, uploadFields } from './upload-fields.js'

test('frontend exposes the confirmed audit and optional source material slots', () => {
  assert.deepEqual(
    uploadFields.map(({ key, accept }) => ({ key, accept })),
    [
      { key: 'auditMaterials', accept: '.pdf,.doc,.docx,.xls,.xlsx,.xlsm' },
      { key: 'reportingWorkbook', accept: '.xls,.xlsx,.xlsm' },
      { key: 'incomeWorkbook', accept: '.xls,.xlsx,.xlsm' },
      { key: 'registryMaterials', accept: '.pdf,.doc,.docx,.ppt,.pptx' },
      { key: 'ownershipHistoryMaterials', accept: '.pdf,.doc,.docx,.ppt,.pptx,.xls,.xlsx,.xlsm' },
      { key: 'unrecordedIntangiblesMaterials', accept: '.pdf,.doc,.docx,.ppt,.pptx,.xls,.xlsx,.xlsm' },
      { key: 'companyProfileMaterials', accept: '.pdf,.doc,.docx,.ppt,.pptx' },
    ],
  )
})

test('upload state keeps audit and supporting files as multiple selections', () => {
  assert.deepEqual(createUploadState(), {
    auditMaterials: [],
    reportingWorkbook: null,
    incomeWorkbook: null,
    registryMaterials: [],
    ownershipHistoryMaterials: [],
    unrecordedIntangiblesMaterials: [],
    companyProfileMaterials: [],
  })
})
