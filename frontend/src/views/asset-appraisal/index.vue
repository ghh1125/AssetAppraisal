<script setup>
import { computed, onBeforeUnmount, reactive, ref } from 'vue'
import { message } from 'ant-design-vue'
import { useI18n } from 'vue-i18n'
import { artifactUrl } from '../../api/request'
import { checkAssetAppraisalOcrCache, createAssetAppraisalRun, getAssetAppraisalRun } from '../../api/asset-appraisal'

const { t } = useI18n()

const form = reactive({
  commissioning_party_name: '',
  commissioning_party_short_name: '',
  target_company_name: '',
  report_serial: '',
  valuation_purpose_inputs: '',
  selected_valuation_method: ['收益法', '资产基础法'],
  valuation_subject_type: '股东全部权益价值',
  transaction_type: '收购',
  final_valuation_method: '收益法',
  target_company_short_name: '',
  narrative_modules: ['industry_overview', 'business_and_segments', 'main_products', 'customers_suppliers', 'profit_model_swot', 'comparable_list'],
})
const files = reactive({ pdf: null, referenceReport: null, auditedFinancials: null, incomeWorkbook: null, reportingWorkbook: null })
const useGlm = ref(true)
const useQichacha = ref(true)
const reuseOcr = ref(true)
const ocrCache = ref({ checking: false, hit: false, source: '' })
const submitting = ref(false)
const run = ref(null)
let pollTimer = null

const canSubmit = computed(() => Boolean(
  files.pdf
  && files.referenceReport
  && files.auditedFinancials
  && files.incomeWorkbook
  && files.reportingWorkbook
  && form.commissioning_party_name
  && form.commissioning_party_short_name
  && form.target_company_name
  && form.target_company_short_name
  && form.report_serial
  && form.valuation_purpose_inputs
  && form.selected_valuation_method.length
  && form.narrative_modules.length
))
const statusText = computed(() => t(`asset.${run.value?.status || 'queued'}`))

function setFile(type, event) {
  files[type] = event.fileList?.[0]?.originFileObj || null
  if (type === 'pdf' && files.pdf) checkOcrCache(files.pdf)
  if (type === 'pdf' && !files.pdf) ocrCache.value = { checking: false, hit: false, source: '' }
}

async function checkOcrCache(file) {
  ocrCache.value = { checking: true, hit: false, source: '' }
  try {
    const result = await checkAssetAppraisalOcrCache(file)
    ocrCache.value = result
  } catch (error) {
    ocrCache.value = { checking: false, hit: false, source: '' }
    message.warning(error.message || t('asset.ocrCacheCheckFailed'))
  }
}

function clearPoll() {
  if (pollTimer) window.clearTimeout(pollTimer)
  pollTimer = null
}

async function refreshRun(runId) {
  try {
    run.value = await getAssetAppraisalRun(runId)
    if (['queued', 'running'].includes(run.value.status)) {
      pollTimer = window.setTimeout(() => refreshRun(runId), 1500)
    } else if (run.value.status === 'completed') {
    message.success(t('asset.completeMessage'))
    }
  } catch (error) {
    clearPoll()
    message.error(error.message || t('asset.cannotStatus'))
  }
}

async function submit() {
  if (!canSubmit.value) {
    message.warning(t('asset.requiredHint'))
    return
  }
  clearPoll()
  submitting.value = true
  run.value = null
  try {
    const result = await createAssetAppraisalRun().mutationFn({
      pdf: files.pdf,
      referenceReport: files.referenceReport,
      auditedFinancials: files.auditedFinancials,
      incomeWorkbook: files.incomeWorkbook,
      reportingWorkbook: files.reportingWorkbook,
      inputs: { ...form },
      useGlm: useGlm.value,
      useQichacha: useQichacha.value,
      reuseOcr: reuseOcr.value,
    })
    run.value = result
    await refreshRun(result.run_id)
  } catch (error) {
    message.error(error.message || t('asset.createFailed'))
  } finally {
    submitting.value = false
  }
}

onBeforeUnmount(clearPoll)
</script>

<template>
  <main class="app-shell">
    <header class="topbar">
      <div>
        <div class="eyebrow">{{ t('asset.eyebrow') }}</div>
        <h1>{{ t('asset.title') }}</h1>
        <p>{{ t('asset.subtitle') }}</p>
      </div>
      <a-tag color="blue">{{ t('asset.readonly') }}</a-tag>
    </header>

    <section class="workspace-grid">
      <a-card class="panel" :title="t('asset.uploadTitle')" :bordered="false">
        <a-alert :message="t('asset.uploadInfo')" type="info" show-icon />
        <div class="upload-grid">
          <a-upload-dragger :max-count="1" accept=".pdf" :before-upload="() => false" @change="setFile('pdf', $event)">
            <p class="upload-icon">PDF</p>
            <p class="upload-title">{{ t('asset.pdfTitle') }}</p>
            <p class="upload-hint">{{ t('asset.uploadHint') }}</p>
          </a-upload-dragger>
          <a-upload-dragger :max-count="1" accept=".docx" :before-upload="() => false" @change="setFile('referenceReport', $event)">
            <p class="upload-icon docx">DOCX</p>
            <p class="upload-title">{{ t('asset.referenceReportTitle') }}</p>
            <p class="upload-hint">{{ t('asset.referenceReportHint') }}</p>
          </a-upload-dragger>
          <a-upload-dragger :max-count="1" accept=".xlsx" :before-upload="() => false" @change="setFile('auditedFinancials', $event)">
            <p class="upload-icon xlsx">XLSX</p>
            <p class="upload-title">{{ t('asset.auditedFinancialsTitle') }}</p>
            <p class="upload-hint">{{ t('asset.auditedFinancialsHint') }}</p>
          </a-upload-dragger>
          <a-upload-dragger :max-count="1" accept=".xlsx" :before-upload="() => false" @change="setFile('incomeWorkbook', $event)">
            <p class="upload-icon xlsx">XLSX</p>
            <p class="upload-title">{{ t('asset.incomeWorkbookTitle') }}</p>
            <p class="upload-hint">{{ t('asset.incomeWorkbookHint') }}</p>
          </a-upload-dragger>
          <a-upload-dragger :max-count="1" accept=".xlsx" :before-upload="() => false" @change="setFile('reportingWorkbook', $event)">
            <p class="upload-icon xlsx">XLSX</p>
            <p class="upload-title">{{ t('asset.reportingWorkbookTitle') }}</p>
            <p class="upload-hint">{{ t('asset.reportingWorkbookHint') }}</p>
          </a-upload-dragger>
        </div>
        <a-alert class="template-source" :message="t('asset.templateSource')" type="success" show-icon />
        <a-alert v-if="ocrCache.checking" class="ocr-cache-status" :message="t('asset.ocrCacheChecking')" type="info" show-icon />
        <a-alert v-else-if="ocrCache.hit" class="ocr-cache-status" :message="t('asset.ocrCacheHit', { source: ocrCache.source })" type="success" show-icon />
        <a-alert v-else-if="files.pdf" class="ocr-cache-status" :message="t('asset.ocrCacheMiss')" type="warning" show-icon />
      </a-card>

      <a-card class="panel" :title="t('asset.inputTitle')" :bordered="false">
        <a-form layout="vertical">
          <div class="form-row">
            <a-form-item :label="t('asset.commissioningName')" required><a-input v-model:value="form.commissioning_party_name" /></a-form-item>
            <a-form-item :label="t('asset.targetName')" required><a-input v-model:value="form.target_company_name" :placeholder="t('asset.targetNamePlaceholder')" /></a-form-item>
          </div>
          <div class="form-row">
            <a-form-item :label="t('asset.transaction')" required><a-select v-model:value="form.transaction_type"><a-select-option value="转让">{{ t('asset.transactionOptions.transfer') }}</a-select-option><a-select-option value="收购">{{ t('asset.transactionOptions.acquisition') }}</a-select-option><a-select-option value="增资">{{ t('asset.transactionOptions.capitalIncrease') }}</a-select-option><a-select-option value="减资">{{ t('asset.transactionOptions.capitalDecrease') }}</a-select-option></a-select></a-form-item>
            <a-form-item :label="t('asset.commissioningShortName')" required><a-input v-model:value="form.commissioning_party_short_name" /></a-form-item>
          </div>
          <a-form-item :label="t('asset.purpose')" required><a-textarea v-model:value="form.valuation_purpose_inputs" :rows="3" :placeholder="t('asset.purposePlaceholder')" /></a-form-item>
          <div class="form-row three">
            <a-form-item :label="t('asset.method')" required><a-select v-model:value="form.selected_valuation_method" mode="multiple" :max-tag-count="2"><a-select-option value="资产基础法">{{ t('asset.methodOptions.asset') }}</a-select-option><a-select-option value="收益法">{{ t('asset.methodOptions.income') }}</a-select-option><a-select-option value="市场法">{{ t('asset.methodOptions.market') }}</a-select-option></a-select></a-form-item>
            <a-form-item :label="t('asset.subject')"><a-select v-model:value="form.valuation_subject_type"><a-select-option value="股东全部权益价值">股东全部权益价值</a-select-option><a-select-option value="股东部分权益价值">股东部分权益价值</a-select-option><a-select-option value="企业整体价值">企业整体价值</a-select-option><a-select-option value="资产组价值">资产组价值</a-select-option></a-select></a-form-item>
            <a-form-item :label="t('asset.finalMethod')" required><a-select v-model:value="form.final_valuation_method"><a-select-option value="资产基础法">{{ t('asset.methodOptions.asset') }}</a-select-option><a-select-option value="收益法">{{ t('asset.methodOptions.income') }}</a-select-option><a-select-option value="市场法">{{ t('asset.methodOptions.market') }}</a-select-option></a-select></a-form-item>
          </div>
          <div class="form-row">
            <a-form-item :label="t('asset.reportSerial')" required><a-input v-model:value="form.report_serial" /></a-form-item>
            <a-form-item :label="t('asset.targetShortName')"><a-input v-model:value="form.target_company_short_name" /></a-form-item>
          </div>
          <a-form-item :label="t('asset.moduleTitle')" required><a-checkbox-group v-model:value="form.narrative_modules" :options="[
            { label: t('asset.modules.industry'), value: 'industry_overview' },
            { label: t('asset.modules.business'), value: 'business_and_segments' },
            { label: t('asset.modules.products'), value: 'main_products' },
            { label: t('asset.modules.customers'), value: 'customers_suppliers' },
            { label: t('asset.modules.swot'), value: 'profit_model_swot' },
            { label: t('asset.modules.comparable'), value: 'comparable_list' },
          ]" /> </a-form-item>
          <div class="switches"><a-checkbox v-model:checked="useGlm">{{ t('asset.useGlm') }}</a-checkbox><a-checkbox v-model:checked="useQichacha">{{ t('asset.useQichacha') }}</a-checkbox><a-checkbox v-model:checked="reuseOcr">{{ t('asset.reuseOcr') }}</a-checkbox></div>
        </a-form>
      </a-card>
    </section>

    <section class="run-bar">
      <div><strong>{{ t('asset.generate') }}</strong><span>{{ t('asset.generateHint') }}</span></div>
      <a-button type="primary" size="large" :loading="submitting" :disabled="!canSubmit" @click="submit">{{ t('asset.start') }}</a-button>
    </section>

    <a-card v-if="run" class="panel result-panel" :title="t('asset.result')" :bordered="false">
      <div class="result-head"><div><span class="run-id">{{ t('asset.task') }} {{ run.run_id }}</span><a-tag :color="run.status === 'failed' ? 'red' : run.status === 'completed' ? 'green' : 'blue'">{{ statusText }}</a-tag></div><span v-if="run.message">{{ run.message }}</span></div>
      <a-progress v-if="['queued', 'running'].includes(run.status)" :percent="run.progress || 0" status="active" />
      <a-alert v-if="run.status === 'failed'" :message="run.error || t('asset.taskFailed')" type="error" show-icon />
      <div v-if="run.status === 'completed'" class="artifact-list">
        <a :href="artifactUrl(run.run_id, artifact.name)" target="_blank" v-for="artifact in run.artifacts" :key="artifact.name">{{ artifact.label || artifact.name }}</a>
      </div>
    </a-card>
  </main>
</template>

<style scoped>
.app-shell { max-width: 1320px; margin: 0 auto; padding: 42px 28px 72px; }
.topbar { display:flex; justify-content:space-between; align-items:flex-start; gap:24px; margin-bottom:28px; }
.eyebrow { color:var(--c2m-color-primary); font-size:12px; letter-spacing:.16em; font-weight:700; }
h1 { margin:8px 0 8px; font-size:34px; color:var(--c2m-text-primary); }
.topbar p { margin:0; color:var(--c2m-text-secondary); }
.workspace-grid { display:grid; grid-template-columns: .92fr 1.4fr; gap:20px; }
.panel { border-radius:18px; box-shadow:0 8px 30px rgba(31,53,81,.07); }
.upload-grid { display:grid; grid-template-columns:1fr 1fr; gap:14px; margin-top:18px; }
.ocr-cache-status { margin-top:14px; }
.template-source { margin-top:14px; }
.upload-icon { margin:6px 0 12px; color:var(--c2m-color-primary); font-weight:800; letter-spacing:.1em; }
.upload-icon.docx { color:#7b61ff; }
.upload-title { font-weight:650; color:var(--c2m-text-primary); }
.upload-hint { color:var(--c2m-text-secondary); font-size:13px; }
.form-row { display:grid; grid-template-columns:1fr 1fr; gap:14px; }
.form-row.three { grid-template-columns:1fr 1fr 1fr; }
.switches { display:flex; align-items:center; gap:18px; padding:5px 0 24px; }
.run-bar { margin:20px 0; padding:18px 22px; border-radius:16px; background:var(--c2m-bg-card); display:flex; justify-content:space-between; align-items:center; gap:18px; box-shadow:0 8px 30px rgba(31,53,81,.06); }
.run-bar span { display:block; color:var(--c2m-text-secondary); font-size:13px; margin-top:5px; }
.result-head { display:flex; justify-content:space-between; gap:12px; margin-bottom:18px; color:var(--c2m-text-secondary); }
.run-id { margin-right:12px; font-family:monospace; }
.artifact-list { display:flex; flex-wrap:wrap; gap:12px; }
.artifact-list a { padding:10px 14px; border:1px solid var(--c2m-border-light); border-radius:10px; color:var(--c2m-color-primary); background:#f8fbff; }
@media (max-width: 900px) { .workspace-grid { grid-template-columns:1fr; } .topbar, .run-bar { flex-direction:column; } .form-row, .form-row.three, .upload-grid { grid-template-columns:1fr; } }
</style>
