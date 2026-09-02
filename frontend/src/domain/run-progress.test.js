import assert from 'node:assert/strict'
import test from 'node:test'

import { currentRunProgress } from './run-progress.js'

test('shows the active substep and its live message above the progress bar', () => {
  assert.deepEqual(
    currentRunProgress({
      status: 'running',
      message: 'LLM 正在复核 12 个字段',
      nodes: [
        {
          name: '节点 3：材料解析 / LLM 候选',
          status: 'running',
          steps: [
            { name: '整合并比对来源', status: 'completed', description: '已完成' },
            { name: 'LLM 证据复核批注', status: 'running', message: '正在复核第 3/12 个字段' },
          ],
        },
      ],
    }),
    {
      label: '节点 3：材料解析 / LLM 候选 · 正在让 LLM 复核证据并生成审核批注',
      detail: '正在复核第 3/12 个字段',
    },
  )
})

test('uses readable action wording for API and source reconciliation steps', () => {
  assert.equal(
    currentRunProgress({
      status: 'running',
      nodes: [{
        name: '节点 3：材料解析 / LLM 候选',
        steps: [{ key: 'query_qichacha', name: '企查查 API 搜索', status: 'running', message: '正在整理返回证据' }],
      }],
    }).label,
    '节点 3：材料解析 / LLM 候选 · 正在调用企查查 API 查询企业信息',
  )
})

test('falls back to the run message after all nodes complete', () => {
  assert.deepEqual(
    currentRunProgress({ status: 'completed', message: '评估报告 Word 已生成', nodes: [] }),
    { label: '已完成', detail: '评估报告 Word 已生成' },
  )
})
