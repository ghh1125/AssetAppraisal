export const uploadFields = Object.freeze([
  {
    key: 'pdf',
    accept: '.pdf',
    icon: 'PDF',
    titleKey: 'auditPdfTitle',
    hintKey: 'auditPdfHint',
    multiple: false,
  },
  {
    key: 'reportingWorkbook',
    accept: '.xls,.xlsx,.xlsm',
    icon: 'Excel',
    titleKey: 'reportingWorkbookTitle',
    hintKey: 'reportingWorkbookHint',
    multiple: false,
  },
  {
    key: 'incomeWorkbook',
    accept: '.xls,.xlsx,.xlsm',
    icon: 'Excel',
    titleKey: 'incomeWorkbookTitle',
    hintKey: 'incomeWorkbookHint',
    multiple: false,
  },
])

export function createUploadState() {
  return Object.fromEntries(uploadFields.map(({ key, multiple }) => [key, multiple ? [] : null]))
}
