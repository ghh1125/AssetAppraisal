import assert from 'node:assert/strict'
import test from 'node:test'

import { canSubmitPartial } from './submission.js'

test('all required manual inputs can submit without files', () => {
  assert.equal(
    canSubmitPartial(
      {},
      {
        commissioning_party_name: '委托方有限公司',
        commissioning_party_short_name: '委托方',
        transaction_type: '收购',
        target_company_name: '被评估有限公司',
        target_company_short_name: '被评估',
        valuation_subject_type: '股东全部权益价值',
        selected_valuation_method: ['收益法'],
        final_valuation_method: '收益法',
        report_serial: 1,
      },
    ),
    true,
  )
})

test('files remain optional and cannot bypass required manual inputs', () => {
  assert.equal(canSubmitPartial({ incomeWorkbook: {} }, {}), false)
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
