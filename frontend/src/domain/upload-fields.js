export const uploadFields = Object.freeze([
  {
    key: 'auditMaterials',
    accept: '.pdf,.doc,.docx,.xls,.xlsx,.xlsm',
    icon: '审计',
    titleKey: 'auditMaterialsTitle',
    hintKey: 'auditMaterialsHint',
    multiple: true,
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
  {
    key: 'registryMaterials',
    accept: '.pdf,.doc,.docx,.ppt,.pptx',
    icon: '工商',
    titleKey: 'registryMaterialsTitle',
    hintKey: 'registryMaterialsHint',
    multiple: true,
    sourceStrategy: 'registry_info_strategy',
  },
  {
    key: 'ownershipHistoryMaterials',
    accept: '.pdf,.doc,.docx,.ppt,.pptx,.xls,.xlsx,.xlsm',
    icon: '股权',
    titleKey: 'ownershipHistoryMaterialsTitle',
    hintKey: 'ownershipHistoryMaterialsHint',
    multiple: true,
    sourceStrategy: 'ownership_history_strategy',
  },
  {
    key: 'unrecordedIntangiblesMaterials',
    accept: '.pdf,.doc,.docx,.ppt,.pptx,.xls,.xlsx,.xlsm',
    icon: '无形',
    titleKey: 'unrecordedIntangiblesMaterialsTitle',
    hintKey: 'unrecordedIntangiblesMaterialsHint',
    multiple: true,
    sourceStrategy: 'unrecorded_intangibles_strategy',
  },
  {
    key: 'companyProfileMaterials',
    accept: '.pdf,.doc,.docx,.ppt,.pptx',
    icon: '介绍',
    titleKey: 'companyProfileMaterialsTitle',
    hintKey: 'companyProfileMaterialsHint',
    multiple: true,
    sourceStrategy: 'company_profile_strategy',
  },
])

export function createUploadState() {
  return Object.fromEntries(uploadFields.map(({ key, multiple }) => [key, multiple ? [] : null]))
}
