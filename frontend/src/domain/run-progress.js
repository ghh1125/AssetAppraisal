const ACTIVE_STEP_LABELS = Object.freeze({
  validate_inputs: '正在校验人工输入',
  store_materials: '正在整理上传材料',
  load_template: '正在加载 Word 模板',
  detect_materials: '正在识别材料类型',
  ocr_pdf: '正在解析审计 PDF',
  parse_excel: '正在解析 Excel 表格',
  reconcile_sources: '正在整合并比对 PDF 与 Excel 来源',
  review_evidence: '正在让 LLM 复核证据并生成审核批注',
  query_qichacha: '正在调用企查查 API 查询企业信息',
  generate_candidates: '正在生成 LLM 候选内容',
  wait_selection: '等待人工选择要写入的模块',
  load_selection: '正在读取人工选择结果',
  map_word: '正在匹配 Word 批注位置',
  fill_fields: '正在填写文字字段',
  fill_tables: '正在填写财务表格',
  write_comments: '正在写入审核批注',
  check_placeholders: '正在检查未填位置',
  save_word: '正在保存评估报告 Word',
  verify_output: '正在检查输出文件',
})
const ACTIVE_STEP_NAME_LABELS = Object.freeze({
  'LLM 证据复核批注': '正在让 LLM 复核证据并生成审核批注',
  '企查查 API 搜索': '正在调用企查查 API 查询企业信息',
  '生成 LLM 候选': '正在生成 LLM 候选内容',
})

function readableActiveStep(step) {
  const mapped = ACTIVE_STEP_LABELS[step?.key] || ACTIVE_STEP_NAME_LABELS[step?.name]
  if (mapped) return mapped
  const name = String(step?.name || '当前步骤')
  return /^(正在|等待|已)/.test(name) ? name : `正在${name}`
}

export function currentRunProgress(run) {
  const nodes = Array.isArray(run?.nodes) ? run.nodes : []
  for (const node of nodes) {
    const activeStep = (Array.isArray(node.steps) ? node.steps : [])
      .find(step => step.status === 'running')
    if (activeStep) {
      return {
        label: `${node.name} · ${readableActiveStep(activeStep)}`,
        detail: activeStep.message || activeStep.description || run?.message || '',
      }
    }
  }

  if (run?.status === 'completed') {
    return { label: '已完成', detail: run.message || '全部步骤已完成' }
  }
  if (run?.status === 'failed') {
    return { label: '执行失败', detail: run.error || run.message || '请查看错误信息' }
  }
  if (run?.status === 'awaiting_selection') {
    return { label: '等待人工选择', detail: run.message || '请选择要写入 Word 的候选内容' }
  }
  return { label: '准备执行', detail: run?.message || '正在等待工作流状态' }
}
