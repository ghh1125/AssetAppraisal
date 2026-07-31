export const uploadFields = Object.freeze([
  {
    key: 'materials',
    accept: '.pdf,.doc,.docx,.xls,.xlsx,.xlsm',
    icon: 'PDF · Excel · Word',
    titleKey: 'materialsTitle',
    hintKey: 'materialsHint',
    multiple: true,
  },
])

export function createUploadState() {
  return Object.fromEntries(uploadFields.map(({ key, multiple }) => [key, multiple ? [] : null]))
}
