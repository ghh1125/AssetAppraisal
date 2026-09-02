function meaningful(value) {
  if (Array.isArray(value)) return value.length > 0
  if (value && typeof value === 'object') {
    return Object.values(value).some(meaningful)
  }
  return String(value ?? '').trim().length > 0
}

export function canSubmitPartial(files, inputs, workbookIntake = null) {
  const required = [
    inputs?.commissioning_party_name,
    inputs?.commissioning_party_short_name,
    inputs?.transaction_type,
    inputs?.target_company_name,
    inputs?.target_company_short_name,
    inputs?.valuation_subject_type,
    inputs?.selected_valuation_method,
    inputs?.final_valuation_method,
    inputs?.valuation_base_date,
  ]
  const hasAuditMaterials = Array.isArray(files?.auditMaterials) && files.auditMaterials.length > 0
  const hasDirectWorkbook = Boolean(files?.reportingWorkbook || files?.incomeWorkbook)
  const hasGeneratedWorkbooks = workbookIntake?.status === 'completed'
  return required.every(meaningful) && (hasAuditMaterials || hasDirectWorkbook || hasGeneratedWorkbooks)
}
