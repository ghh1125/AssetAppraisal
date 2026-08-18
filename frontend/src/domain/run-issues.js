const fieldNames = {
  asset_scope_summary_table: '资产负债范围表',
  historical_balance_sheet_table: '历史资产负债表',
  historical_income_statement_table: '历史利润表',
  long_term_assets_table: '主要长期资产表',
  major_long_term_assets: '主要长期资产',
  book_net_assets: '账面净资产',
  tax_rates: '税率',
  valuation_scope: '评估范围',
}

function fieldName(raw) {
  const key = raw.split('：', 1)[0]
  return fieldNames[key] || key || '该字段'
}

function workbookName(raw) {
  const fallback = raw.match(/已采用([^（；]+?\.xlsx)/)
  if (fallback) return fallback[1]
  const conflict = raw.match(/其他材料\s+([^（；]+?\.xlsx)/)
  return conflict?.[1] || ''
}

function sourceDetail(raw, source) {
  if (!source) return ''
  const key = raw.split('：', 1)[0]
  const tableNames = {
    asset_scope_summary_table: '资产负债范围表',
    historical_balance_sheet_table: '历史资产负债表',
    historical_income_statement_table: '历史利润表',
    long_term_assets_table: '主要长期资产账面记录表',
  }
  if (tableNames[key] && raw.includes(`${key}`)) return `${source}（${tableNames[key]}）`
  const match = raw.match(new RegExp(`${source.replace(/[.*+?^${}()|[\\]\\\\]/g, '\\$&')}（([^）]+)!`))
  if (match) return `${source}（${match[1]}）`
  return source
}

export function summarizeRunIssues(rawIssues = []) {
  const summaries = []
  let unresolvedCount = 0

  for (const rawValue of rawIssues) {
    const raw = String(rawValue || '')
    if (!raw) continue
    if (raw.includes('生成后仍存在未解析占位符')) {
      unresolvedCount += 1
      continue
    }

    const name = fieldName(raw)
    const source = workbookName(raw)
    const sourceLabel = sourceDetail(raw, source)
    if (raw.includes('数据不一致')) {
      summaries.push(`${name}：审计 PDF 与“${sourceLabel || '其他材料'}”数据不一致；报告已采用审计 PDF 数值并标红，请人工复核。`)
      continue
    }
    if (raw.includes('审计PDF已上传，但OCR未识别到该字段')) {
      summaries.push(`${name}：审计 PDF 已上传，但未识别到对应数据；已使用“${sourceLabel || '上传表格'}”填入，暂未完成核对。`)
      continue
    }
    if (raw.includes('未执行PDF对照') || raw.includes('暂未完成PDF对照')) {
      summaries.push(`${name}：已使用“${sourceLabel || '上传表格'}”填入，暂未能与审计 PDF 核对。`)
      continue
    }
    if (raw.includes('无对应PDF文件，无法获取审计数据') || raw.includes('未上传审计PDF，无法获取审计数据')) {
      summaries.push(`${name}：未取得可用的审计 PDF 数据，Word 中已保留黄色 XXX。`)
      continue
    }
    if (raw.startsWith('LLM取数复核：')) {
      summaries.push('取数复核：存在需要人工确认的证据或口径差异，已在对应 Word 数字处添加批注。')
    }
  }

  if (unresolvedCount) summaries.push(`报告仍有 ${unresolvedCount} 处待补充内容，已在 Word 中保留黄色 XXX。`)
  return [...new Set(summaries)]
}
