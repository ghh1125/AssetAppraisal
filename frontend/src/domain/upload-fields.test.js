import test from 'node:test'
import assert from 'node:assert/strict'

import { createUploadState, uploadFields } from './upload-fields.js'

test('frontend exposes the confirmed audit and optional source material slots', () => {
  assert.deepEqual(
    uploadFields.map(({ key, accept }) => ({ key, accept })),
    [
      { key: 'materialArchive', accept: '.rar,.zip' },
      { key: 'auditMaterials', accept: '.pdf,.doc,.docx,.xls,.xlsx,.xlsm,.png,.jpg,.jpeg' },
      { key: 'reportingWorkbook', accept: '.xls,.xlsx,.xlsm' },
      { key: 'incomeWorkbook', accept: '.xls,.xlsx,.xlsm' },
      { key: 'registryMaterials', accept: '.pdf,.doc,.docx,.ppt,.pptx,.png,.jpg,.jpeg' },
      { key: 'ownershipHistoryMaterials', accept: '.pdf,.doc,.docx,.ppt,.pptx,.xls,.xlsx,.xlsm,.png,.jpg,.jpeg' },
      { key: 'unrecordedIntangiblesMaterials', accept: '.pdf,.doc,.docx,.ppt,.pptx,.xls,.xlsx,.xlsm,.png,.jpg,.jpeg' },
      { key: 'companyProfileMaterials', accept: '.pdf,.doc,.docx,.ppt,.pptx,.xls,.xlsx,.xlsm,.png,.jpg,.jpeg' },
    ],
  )
})

test('upload state keeps audit and supporting files as multiple selections', () => {
  assert.deepEqual(createUploadState(), {
    materialArchive: null,
    auditMaterials: [],
    reportingWorkbook: null,
    incomeWorkbook: null,
    registryMaterials: [],
    ownershipHistoryMaterials: [],
    unrecordedIntangiblesMaterials: [],
    companyProfileMaterials: [],
  })
})
