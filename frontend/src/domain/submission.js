function meaningful(value) {
  if (Array.isArray(value)) return value.length > 0
  if (value && typeof value === 'object') {
    return Object.values(value).some(meaningful)
  }
  return String(value ?? '').trim().length > 0
}

export function canSubmitPartial(files, inputs) {
  return Object.values(files || {}).some(Boolean)
    || Object.values(inputs || {}).some(meaningful)
}
