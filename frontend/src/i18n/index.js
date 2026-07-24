import { createI18n, useI18n } from 'vue-i18n'

const messages = {
  'zh-CN': {
    asset: {
      eyebrow: 'MCPIFY · ASSET APPRAISAL', title: '资产评估报告工作台', subtitle: '上传材料，填写黄色提示对应的人工信息，自动生成独立的 Word、OCR Excel 和字段审计清单。', readonly: '模板只读 · 输出不覆盖原文档',
      uploadTitle: '1. 上传材料', uploadInfo: 'PDF 会进入现有 OCR → XLSX → Word 流程；Word 模板由后端固定提供，不需要上传。上传后会立即检查是否已有同一 PDF 的 OCR 结果。', pdfTitle: '上传审计报告 PDF', uploadHint: '拖放文件或点击选择', templateSource: 'Word 模板由后端提供：评估报告版式-沟通标注版.docx（模板只读）', pdfRequired: '请上传 PDF 审计报告', ocrCacheChecking: '正在检查已有 OCR 结果……', ocrCacheHit: '已命中 OCR 缓存：{source}，生成时将跳过 OCR', ocrCacheMiss: '未命中 OCR 缓存，生成时将执行 OCR', ocrCacheCheckFailed: 'OCR 缓存检查失败',
      inputTitle: '2. 黄色提示中的用户输入', commissioningName: '委托方全称', commissioningShortName: '委托方简称', reportSerial: '评估报告编号流水号', targetShortName: '被评估单位简称', targetName: '被评估单位全称（企查查核验）', targetNamePlaceholder: '可选；留空时使用材料中识别到的企业全称', purpose: '评估目的', method: '评估方法', subject: '评估对象', transaction: '交易类型', finalMethod: '最终采用方法', purposePlaceholder: '请填写本次评估的真实业务目的', requiredHint: '请上传 PDF、Word 模板，并填写委托方全称、简称、报告流水号和评估目的',
      useGlm: '启用百炼 GLM', useQichacha: '启用企查查', reuseOcr: '复用已有 OCR 结果', generate: '3. 生成报告', generateHint: '其他字段由企查查、GLM、PDF OCR/XLSX 规则自动处理，无法可靠取得的内容保持空白。', start: '开始生成',
      result: '运行结果', queued: '排队中', running: '处理中', completed: '已完成', failed: '失败', task: '任务', taskFailed: '任务失败，请查看后端日志', completeMessage: '评估报告已生成', cannotStatus: '无法获取任务状态', createFailed: '创建评估任务失败', artifacts: { report: '评估报告 Word', ocr: 'OCR 结构化 Excel', audit: '字段审计清单', manifest: '运行清单', issues: '复核事项', fields: '标准字段结果' },
      qcc: '收购', methodOptions: { both: '收益法、资产基础法', asset: '资产基础法', income: '收益法' },
    },
  },
  'en-US': {
    asset: {
      eyebrow: 'MCPIFY · ASSET APPRAISAL', title: 'Asset Appraisal Workbench', subtitle: 'Upload materials, complete the yellow-note inputs, and generate Word, OCR Excel, and an audit checklist.', readonly: 'Read-only template · outputs never overwrite the source', uploadTitle: '1. Materials', uploadInfo: 'The PDF runs through OCR → XLSX → Word; the DOCX template is supplied by the backend. After upload, the app checks for an OCR result for the same PDF.', pdfTitle: 'Upload audit PDF', uploadHint: 'Drop a file or click to choose', templateSource: 'Word template supplied by backend: 评估报告版式-沟通标注版.docx (read-only)', pdfRequired: 'Please upload the audit PDF', ocrCacheChecking: 'Checking for an existing OCR result…', ocrCacheHit: 'OCR cache hit: {source}; OCR will be skipped', ocrCacheMiss: 'No OCR cache hit; OCR will run during generation', ocrCacheCheckFailed: 'Unable to check the OCR cache',
      inputTitle: '2. User inputs from yellow notes', commissioningName: 'Commissioning party', commissioningShortName: 'Short name', reportSerial: 'Report serial number', targetShortName: 'Target company short name', targetName: 'Target company full name (QCC check)', targetNamePlaceholder: 'Optional; leave blank to use the detected company name', purpose: 'Valuation purpose', method: 'Valuation methods', subject: 'Valuation subject', transaction: 'Transaction type', finalMethod: 'Final method', purposePlaceholder: 'Enter the actual business purpose', requiredHint: 'Upload a PDF, then complete the commissioning party, short name, serial number, and purpose',
      useGlm: 'Use Bailian GLM', useQichacha: 'Use Qichacha', reuseOcr: 'Reuse existing OCR result', generate: '3. Generate report', generateHint: 'Qichacha, GLM, and PDF OCR/XLSX rules fill the remaining fields; unavailable content stays blank.', start: 'Generate', result: 'Run result', queued: 'Queued', running: 'Running', completed: 'Completed', failed: 'Failed', task: 'Task', taskFailed: 'The task failed; see the backend log', completeMessage: 'The report is ready', cannotStatus: 'Unable to read task status', createFailed: 'Unable to create task', artifacts: { report: 'Word report', ocr: 'OCR Excel', audit: 'Audit checklist', manifest: 'Run manifest', issues: 'Review items', fields: 'Normalized fields' }, qcc: 'Acquisition', methodOptions: { both: 'Income and asset-based', asset: 'Asset-based', income: 'Income' },
    },
  },
}

export const i18n = createI18n({ legacy: false, locale: 'zh-CN', fallbackLocale: 'zh-CN', messages })
export { useI18n }
