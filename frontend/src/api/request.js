const API_BASE = import.meta.env?.VITE_API_BASE_URL || ''
const API_PATH = import.meta.env?.VITE_API_PATH || '/api/v1'

function apiUrl(path) {
  return `${API_BASE}${API_PATH}${path}`
}

async function parseResponse(response) {
  const contentType = response.headers.get('content-type') || ''
  const body = contentType.includes('application/json') ? await response.json() : await response.text()
  if (!response.ok) {
    const detail = typeof body === 'object' ? (body.detail || body.message) : body
    throw new Error(detail || `请求失败（${response.status}）`)
  }
  return body
}

export async function postForm(path, formData) {
  const response = await fetch(apiUrl(path), { method: 'POST', body: formData })
  return parseResponse(response)
}

export async function getJson(path) {
  const response = await fetch(apiUrl(path), { headers: { Accept: 'application/json' } })
  return parseResponse(response)
}

export function artifactUrl(runId, name) {
  return apiUrl(`/asset-appraisal/runs/${encodeURIComponent(runId)}/artifacts/${encodeURIComponent(name)}`)
}
