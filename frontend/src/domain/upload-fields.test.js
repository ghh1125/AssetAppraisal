import test from 'node:test'
import assert from 'node:assert/strict'

import { createUploadState, uploadFields } from './upload-fields.js'

test('frontend exposes the three image-defined optional material slots', () => {
  assert.deepEqual(
    uploadFields.map(({ key, accept }) => ({ key, accept })),
    [
      { key: 'pdf', accept: '.pdf' },
      { key: 'reportingWorkbook', accept: '.xls,.xlsx,.xlsm' },
      { key: 'incomeWorkbook', accept: '.xls,.xlsx,.xlsm' },
    ],
  )
})

test('upload state contains the three optional slots', () => {
  assert.deepEqual(createUploadState(), {
    pdf: null,
    reportingWorkbook: null,
    incomeWorkbook: null,
  })
})
