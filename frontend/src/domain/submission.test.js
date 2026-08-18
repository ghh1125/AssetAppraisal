import assert from 'node:assert/strict'
import test from 'node:test'

import { canSubmitPartial } from './submission.js'

test('confirmed required manual inputs include valuation base date and exclude report serial', () => {
  assert.equal(
    canSubmitPartial(
      { auditMaterials: [{}] },
      {
        commissioning_party_name: '委托方有限公司',
        commissioning_party_short_name: '委托方',
        transaction_type: '收购',
        target_company_name: '被评估有限公司',
        target_company_short_name: '被评估',
        valuation_subject_type: '股东全部权益价值',
        selected_valuation_method: ['收益法'],
        final_valuation_method: '收益法',
        valuation_base_date: '2025-06-30',
      },
    ),
    true,
  )
})

test('audit materials and manual inputs are both required', () => {
  assert.equal(canSubmitPartial({ auditMaterials: [{}] }, {}), false)
  assert.equal(
    canSubmitPartial(
      { auditMaterials: [] },
      {
        commissioning_party_name: '委托方有限公司',
        commissioning_party_short_name: '委托方',
        transaction_type: '收购',
        target_company_name: '被评估有限公司',
        target_company_short_name: '被评估',
        valuation_subject_type: '股东全部权益价值',
        selected_valuation_method: ['收益法'],
        final_valuation_method: '收益法',
        valuation_base_date: '2025-06-30',
      },
    ),
    false,
  )
})

test('completely empty input cannot submit', () => {
  assert.equal(
    canSubmitPartial(
      { auditMaterials: [], incomeWorkbook: null },
      { target_company_name: '', narrative_modules: [] },
    ),
    false,
  )
})
