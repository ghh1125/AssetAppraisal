import { getJson, postForm } from './request.js'

export function checkAssetAppraisalOcrCache(file) {
  const form = new FormData()
  form.append('pdf', file)
  return postForm('/asset-appraisal/ocr-cache/check', form)
}

export function buildAssetAppraisalForm({
  pdf,
  reportingWorkbook,
  incomeWorkbook,
  inputs,
  useGlm,
  useQichacha,
  reuseOcr,
}) {
  const form = new FormData()
  if (pdf) form.append('pdf', pdf)
  if (reportingWorkbook) form.append('reporting_workbook', reportingWorkbook)
  if (incomeWorkbook) form.append('income_workbook', incomeWorkbook)
  form.append('inputs', JSON.stringify(inputs))
  form.append('use_glm', String(useGlm))
  form.append('use_qichacha', String(useQichacha))
  form.append('reuse_ocr', String(reuseOcr ?? true))
  return form
}

export function createAssetAppraisalRun() {
  return {
    mutationFn: async (payload) => (
      postForm('/asset-appraisal/runs', buildAssetAppraisalForm(payload))
    ),
  }
}

export const getAssetAppraisalRun = (runId) => getJson(`/asset-appraisal/runs/${encodeURIComponent(runId)}`)
