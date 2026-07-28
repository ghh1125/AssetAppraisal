export const uploadFields = Object.freeze([
  {
    key: 'pdf',
    accept: '.pdf',
    icon: 'PDF',
    titleKey: 'pdfTitle',
    hintKey: 'pdfHint',
  },
  {
    key: 'reportingWorkbook',
    accept: '.xlsx,.xlsm',
    icon: 'XLSX',
    titleKey: 'reportingWorkbookTitle',
    hintKey: 'reportingWorkbookHint',
  },
  {
    key: 'incomeWorkbook',
    accept: '.xlsx,.xlsm',
    icon: 'XLSX',
    titleKey: 'incomeWorkbookTitle',
    hintKey: 'incomeWorkbookHint',
  },
])

export function createUploadState() {
  return Object.fromEntries(uploadFields.map(({ key }) => [key, null]))
}
