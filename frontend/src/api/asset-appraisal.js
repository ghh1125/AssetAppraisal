import { getJson, postForm } from './request'

export function checkAssetAppraisalOcrCache(file) {
  const form = new FormData()
  form.append('pdf', file)
  return postForm('/asset-appraisal/ocr-cache/check', form)
}

export function createAssetAppraisalRun() {
  return {
    mutationFn: async ({ pdf, referenceReport, auditedFinancials, incomeWorkbook, reportingWorkbook, inputs, useGlm, useQichacha, reuseOcr }) => {
      const form = new FormData()
      if (pdf) form.append('pdf', pdf)
      if (referenceReport) form.append('reference_report', referenceReport)
      if (auditedFinancials) form.append('audited_financials', auditedFinancials)
      if (incomeWorkbook) form.append('income_workbook', incomeWorkbook)
      if (reportingWorkbook) form.append('reporting_workbook', reportingWorkbook)
      form.append('inputs', JSON.stringify(inputs))
      form.append('use_glm', String(useGlm))
      form.append('use_qichacha', String(useQichacha))
      form.append('reuse_ocr', String(reuseOcr ?? true))
      return postForm('/asset-appraisal/runs', form)
    },
  }
}

export const getAssetAppraisalRun = (runId) => getJson(`/asset-appraisal/runs/${encodeURIComponent(runId)}`)
