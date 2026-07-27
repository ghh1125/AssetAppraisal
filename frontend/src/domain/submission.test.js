import assert from 'node:assert/strict'
import test from 'node:test'

import { canSubmitPartial } from './submission.js'

test('manual-only input can submit', () => {
  assert.equal(
    canSubmitPartial(
      {},
      { target_company_name: '示例有限公司' },
    ),
    true,
  )
})

test('one uploaded file can submit', () => {
  assert.equal(canSubmitPartial({ incomeWorkbook: {} }, {}), true)
})

test('completely empty input cannot submit', () => {
  assert.equal(
    canSubmitPartial(
      { pdf: null, incomeWorkbook: null },
      { target_company_name: '', narrative_modules: [] },
    ),
    false,
  )
})
