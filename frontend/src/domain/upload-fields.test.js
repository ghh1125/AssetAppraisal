import test from 'node:test'
import assert from 'node:assert/strict'

import { createUploadState, uploadFields } from './upload-fields.js'

test('frontend exposes only the three agreed upload slots', () => {
  assert.deepEqual(
    uploadFields.map(({ key, accept }) => ({ key, accept })),
    [
      { key: 'pdf', accept: '.pdf' },
      { key: 'reportingWorkbook', accept: '.xlsx,.xlsm' },
      { key: 'incomeWorkbook', accept: '.xlsx,.xlsm' },
    ],
  )
})

test('upload state contains no reference report or audited-financial slot', () => {
  assert.deepEqual(createUploadState(), {
    pdf: null,
    reportingWorkbook: null,
    incomeWorkbook: null,
  })
})
