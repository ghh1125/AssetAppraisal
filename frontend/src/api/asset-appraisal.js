import { getJson, postForm } from './request'

export function checkAssetAppraisalOcrCache(file) {
  const form = new FormData()
  form.append('pdf', file)
  return postForm('/asset-appraisal/ocr-cache/check', form)
}

export function createAssetAppraisalRun() {
  return {
    mutationFn: async ({ pdf, template, inputs, useGlm, useQichacha, reuseOcr }) => {
      const form = new FormData()
      form.append('pdf', pdf)
      form.append('template', template)
      form.append('inputs', JSON.stringify(inputs))
      form.append('use_glm', String(useGlm))
      form.append('use_qichacha', String(useQichacha))
      form.append('reuse_ocr', String(reuseOcr ?? true))
      return postForm('/asset-appraisal/runs', form)
    },
  }
}

export const getAssetAppraisalRun = (runId) => getJson(`/asset-appraisal/runs/${encodeURIComponent(runId)}`)
