import assert from 'node:assert/strict'
import test from 'node:test'

import { summarizeRunIssues } from './run-issues.js'

test('summarizes an Excel fallback without exposing worksheet cell locators', () => {
  const [issue] = summarizeRunIssues([
    'historical_income_statement_table：无对应PDF文件或未识别到该字段；已采用通富审核后财报-单体1月5日.xlsx（07N_利润表!A4；07N_利润表!F4；07N_利润表!F5），未执行PDF对照',
  ])

  assert.equal(issue, '历史利润表：已使用“通富审核后财报-单体1月5日.xlsx（历史利润表）”填入，暂未能与审计 PDF 核对。')
  assert.doesNotMatch(issue, /07N_|!A4|!F4/)
})

test('explains that an uploaded audit PDF was not recognised instead of claiming it was absent', () => {
  const [issue] = summarizeRunIssues([
    'historical_income_statement_table：审计PDF已上传，但OCR未识别到该字段；已采用财报.xlsx（07N_利润表!A4；07N_利润表!F4），暂未完成PDF对照',
  ])

  assert.equal(issue, '历史利润表：审计 PDF 已上传，但未识别到对应数据；已使用“财报.xlsx（历史利润表）”填入，暂未完成核对。')
  assert.doesNotMatch(issue, /07N_|!A4|!F4/)
})

test('summarizes a source conflict with a direct human review action', () => {
  const [issue] = summarizeRunIssues([
    'asset_scope_summary_table：数据不一致；审计PDF 通富审计报告.pdf（OCR结构化结果.xlsx / asset_scope_summary_table）=XXX；其他材料 财报.xlsx（06N_资产负债表!F28）=XXX。已按审计PDF填入',
  ])

  assert.equal(issue, '资产负债范围表：审计 PDF 与“财报.xlsx（资产负债范围表）”数据不一致；报告已采用审计 PDF 数值并标红，请人工复核。')
})

test('groups raw unresolved placeholder messages into one readable summary', () => {
  const issues = summarizeRunIssues([
    'Word第7页 统一社会信用代码：XXX：生成后仍存在未解析占位符',
    'Word第8页 注册地址：XXX：生成后仍存在未解析占位符',
  ])

  assert.deepEqual(issues, ['报告仍有 2 处待补充内容，已在 Word 中保留黄色 XXX。'])
})
