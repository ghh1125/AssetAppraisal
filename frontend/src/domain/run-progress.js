export function currentRunProgress(run) {
  const nodes = Array.isArray(run?.nodes) ? run.nodes : []
  for (const node of nodes) {
    const activeStep = (Array.isArray(node.steps) ? node.steps : [])
      .find(step => step.status === 'running')
    if (activeStep) {
      return {
        label: `${node.name} · ${activeStep.name}`,
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

export function flattenRunSteps(run) {
  const nodes = Array.isArray(run?.nodes) ? run.nodes : []
  return nodes.flatMap((node) => (
    (Array.isArray(node.steps) ? node.steps : []).map((step) => ({
      key: `${node.key || node.name || 'node'}:${step.key || step.name}`,
      name: step.name || '',
      status: step.status || 'pending',
      message: step.status === 'running'
        ? (step.message || step.description || node.message || '')
        : (step.message || step.description || ''),
    }))
  ))
}
