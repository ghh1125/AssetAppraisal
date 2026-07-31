import test from 'node:test'
import assert from 'node:assert/strict'

import { createUploadState, uploadFields } from './upload-fields.js'

test('frontend exposes one generic multi-file material slot', () => {
  assert.deepEqual(
    uploadFields.map(({ key, accept }) => ({ key, accept })),
    [{ key: 'materials', accept: '.pdf,.doc,.docx,.xls,.xlsx,.xlsm' }],
  )
})

test('upload state contains the related-materials list', () => {
  assert.deepEqual(createUploadState(), {
    materials: [],
  })
})
