import assert from 'node:assert/strict'
import test from 'node:test'

import { currentRunProgress, flattenRunSteps } from './run-progress.js'

test('shows the active substep and its live message above the progress bar', () => {
  assert.deepEqual(
    currentRunProgress({
      status: 'running',
      message: 'LLM 正在复核 12 个字段',
      nodes: [
        {
          name: '节点 2：材料解析 / LLM 候选',
          status: 'running',
          steps: [
            { name: '整合并比对来源', status: 'completed', description: '已完成' },
            { name: 'LLM 证据复核批注', status: 'running', message: '正在复核第 3/12 个字段' },
          ],
        },
      ],
    }),
    {
      label: '节点 2：材料解析 / LLM 候选 · LLM 证据复核批注',
      detail: '正在复核第 3/12 个字段',
    },
  )
})

test('falls back to the run message after all nodes complete', () => {
  assert.deepEqual(
    currentRunProgress({ status: 'completed', message: '评估报告 Word 已生成', nodes: [] }),
    { label: '已完成', detail: '评估报告 Word 已生成' },
  )
})

test('flattens node substeps for the progress strip above the bar', () => {
  const steps = flattenRunSteps({
    nodes: [
      {
        name: '节点 1：开始 / 输入',
        steps: [
          { key: 'validate_inputs', name: '校验人工输入', status: 'completed', description: '检查公司名称' },
        ],
      },
      {
        name: '节点 2：材料解析 / LLM 候选',
        steps: [
          { key: 'review_evidence', name: 'LLM 证据复核批注', status: 'running', message: '正在复核历史利润表' },
          { key: 'query_qichacha', name: '企查查 API 信息检索', status: 'pending', description: '搜索企业工商信息' },
        ],
      },
    ],
  })

  assert.deepEqual(
    steps.map(item => [item.name, item.status, item.message]),
    [
      ['校验人工输入', 'completed', '检查公司名称'],
      ['LLM 证据复核批注', 'running', '正在复核历史利润表'],
      ['企查查 API 信息检索', 'pending', '搜索企业工商信息'],
    ],
  )
})
