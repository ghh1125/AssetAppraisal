<script setup>
import { computed, onBeforeUnmount, reactive, ref } from 'vue'
import { message } from 'ant-design-vue'
import { useI18n } from 'vue-i18n'
import { artifactUrl } from '../../api/request'
import { checkAssetAppraisalOcrCache, createAssetAppraisalRun, getAssetAppraisalRun, selectAssetAppraisalCandidates } from '../../api/asset-appraisal'
import { canSubmitPartial } from '../../domain/submission'
import { createUploadState, uploadFields } from '../../domain/upload-fields'

const { t } = useI18n()

const form = reactive({
  commissioning_party_name: '',
  target_company_name: '',
  selected_valuation_method: ['收益法', '资产基础法'],
  transaction_type: '收购',
  final_valuation_method: '收益法',
})
const files = reactive(createUploadState())
const useGlm = ref(true)
const useQichacha = ref(true)
const reuseOcr = ref(true)
const ocrCache = ref({ checking: false, hit: false, source: '' })
const submitting = ref(false)
const run = ref(null)
const selectedCandidateKeys = ref([])
let pollTimer = null

const canSubmit = computed(() => canSubmitPartial(files, form))
const publicArtifacts = computed(() => (
  (run.value?.artifacts || []).filter(item => item.name === '资产评估报告_待复核.docx')
))
const statusText = computed(() => t(`asset.${run.value?.status || 'queued'}`))
const nodeStatusText = (status) => t(`asset.nodeStatus.${status || 'pending'}`)

function setFile(type, event) {
  if (type === 'materials') {
    files.materials = (event.fileList || [])
      .map(item => item.originFileObj)
      .filter(Boolean)
    const pdf = files.materials.find(item => item.name?.toLowerCase().endsWith('.pdf'))
    if (pdf) checkOcrCache(pdf)
    if (!pdf) ocrCache.value = { checking: false, hit: false, source: '' }
    return
  }
  files[type] = event.fileList?.[0]?.originFileObj || null
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
    } else if (run.value.status === 'awaiting_selection') {
      selectedCandidateKeys.value = (run.value.candidates || []).map(item => item.field_key)
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
  selectedCandidateKeys.value = []
  try {
    const result = await createAssetAppraisalRun().mutationFn({
      materials: files.materials,
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

async function confirmCandidates() {
  if (!run.value?.run_id) return
  const candidateMap = Object.fromEntries(
    (run.value.candidates || [])
      .filter(item => selectedCandidateKeys.value.includes(item.field_key))
      .map(item => [item.field_key, item.value]),
  )
  submitting.value = true
  try {
    run.value = await selectAssetAppraisalCandidates(run.value.run_id, candidateMap)
    await refreshRun(run.value.run_id)
  } catch (error) {
    message.error(error.message || t('asset.selectFailed'))
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
          <a-upload-dragger
            v-for="field in uploadFields"
            :key="field.key"
            :multiple="field.multiple"
            :max-count="field.multiple ? 20 : 1"
            :accept="field.accept"
            :before-upload="() => false"
            @change="setFile(field.key, $event)"
          >
            <p :class="['upload-icon', field.icon.toLowerCase()]">{{ field.icon }}</p>
            <p class="upload-title">{{ t(`asset.${field.titleKey}`) }}</p>
            <p class="upload-hint">{{ t(`asset.${field.hintKey}`) }}</p>
          </a-upload-dragger>
        </div>
        <a-alert class="template-source" :message="t('asset.templateSource')" type="success" show-icon />
        <a-alert v-if="ocrCache.checking" class="ocr-cache-status" :message="t('asset.ocrCacheChecking')" type="info" show-icon />
        <a-alert v-else-if="ocrCache.hit" class="ocr-cache-status" :message="t('asset.ocrCacheHit', { source: ocrCache.source })" type="success" show-icon />
        <a-alert v-else-if="files.materials?.some(file => file?.name?.toLowerCase().endsWith('.pdf'))" class="ocr-cache-status" :message="t('asset.ocrCacheMiss')" type="warning" show-icon />
      </a-card>

      <a-card class="panel" :title="t('asset.inputTitle')" :bordered="false">
        <a-form layout="vertical">
          <div class="form-row">
            <a-form-item :label="t('asset.commissioningName')"><a-input v-model:value="form.commissioning_party_name" /></a-form-item>
            <a-form-item :label="t('asset.targetName')"><a-input v-model:value="form.target_company_name" :placeholder="t('asset.targetNamePlaceholder')" /></a-form-item>
          </div>
          <a-form-item :label="t('asset.transaction')"><a-select v-model:value="form.transaction_type"><a-select-option value="转让">{{ t('asset.transactionOptions.transfer') }}</a-select-option><a-select-option value="收购">{{ t('asset.transactionOptions.acquisition') }}</a-select-option><a-select-option value="增资">{{ t('asset.transactionOptions.capitalIncrease') }}</a-select-option><a-select-option value="减资">{{ t('asset.transactionOptions.capitalDecrease') }}</a-select-option></a-select></a-form-item>
          <a-form-item :label="t('asset.method')"><a-select v-model:value="form.selected_valuation_method" mode="multiple" :max-tag-count="3"><a-select-option value="资产基础法">{{ t('asset.methodOptions.asset') }}</a-select-option><a-select-option value="收益法">{{ t('asset.methodOptions.income') }}</a-select-option><a-select-option value="市场法">{{ t('asset.methodOptions.market') }}</a-select-option></a-select></a-form-item>
          <a-form-item :label="t('asset.finalMethod')"><a-select v-model:value="form.final_valuation_method"><a-select-option value="资产基础法">{{ t('asset.methodOptions.asset') }}</a-select-option><a-select-option value="收益法">{{ t('asset.methodOptions.income') }}</a-select-option><a-select-option value="市场法">{{ t('asset.methodOptions.market') }}</a-select-option></a-select></a-form-item>
        </a-form>
      </a-card>
    </section>

    <section class="run-bar">
      <div><strong>{{ t('asset.generate') }}</strong><span>{{ t('asset.generateHint') }}</span></div>
      <a-button type="primary" size="large" :loading="submitting" :disabled="!canSubmit" @click="submit">{{ t('asset.start') }}</a-button>
    </section>

    <a-card v-if="run" class="panel result-panel" :title="t('asset.result')" :bordered="false">
      <div class="result-head"><div><span class="run-id">{{ t('asset.task') }} {{ run.run_id }}</span><a-tag :color="run.status === 'failed' ? 'red' : run.status === 'completed' ? 'green' : 'blue'">{{ statusText }}</a-tag></div><span v-if="run.message">{{ run.message }}</span></div>
      <div v-if="run.status === 'completed' && publicArtifacts.length" class="artifact-list result-artifact">
        <a :href="artifactUrl(run.run_id, artifact.name)" target="_blank" v-for="artifact in publicArtifacts" :key="artifact.name">{{ artifact.label || artifact.name }}</a>
      </div>
      <div v-if="run.nodes?.length" class="node-progress" aria-label="workflow nodes">
        <div v-for="(node, index) in run.nodes" :key="node.key" :class="['node-step', `node-${node.status}`]">
          <div class="node-marker">{{ index + 1 }}</div>
          <div class="node-copy">
            <div class="node-title"><strong>{{ node.name }}</strong><a-tag :color="node.status === 'failed' ? 'red' : node.status === 'completed' ? 'green' : node.status === 'awaiting_selection' ? 'orange' : node.status === 'running' ? 'blue' : 'default'">{{ nodeStatusText(node.status) }}</a-tag></div>
            <div class="node-description">{{ node.description }}</div>
            <div v-if="node.message" class="node-message">{{ node.message }}</div>
          </div>
        </div>
      </div>
      <a-progress v-if="['queued', 'running'].includes(run.status)" :percent="run.progress || 0" status="active" />
      <a-alert v-if="run.status === 'failed'" :message="run.error || t('asset.taskFailed')" type="error" show-icon />
      <div v-if="run.status === 'awaiting_selection'" class="candidate-panel">
        <a-alert :message="t('asset.candidateHint')" type="info" show-icon />
        <a-alert v-if="!run.candidates?.length" :message="t('asset.candidateEmpty')" type="warning" show-icon />
        <div v-if="!run.candidates?.length && run.issues?.length" class="candidate-issues">
          <strong>{{ t('asset.candidateIssues') }}</strong>
          <div v-for="issue in run.issues" :key="issue">{{ issue }}</div>
        </div>
        <a-checkbox-group v-model:value="selectedCandidateKeys" class="candidate-list">
          <div v-for="candidate in run.candidates" :key="candidate.field_key" class="candidate-item">
            <a-checkbox :value="candidate.field_key">{{ candidate.field_name || candidate.field_key }}</a-checkbox>
            <span v-if="candidate.location_ids?.length" class="candidate-location">{{ candidate.location_ids.join('、') }}</span>
            <div class="candidate-value">{{ candidate.value }}</div>
          </div>
        </a-checkbox-group>
        <a-button type="primary" :loading="submitting" @click="confirmCandidates">{{ t('asset.confirmCandidates') }}</a-button>
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
.upload-grid { display:grid; grid-template-columns:1fr; gap:14px; margin-top:18px; }
.ocr-cache-status { margin-top:14px; }
.template-source { margin-top:14px; }
.upload-icon { margin:6px 0 12px; color:var(--c2m-color-primary); font-weight:800; letter-spacing:.1em; }
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
.result-artifact { margin-bottom:20px; }
.candidate-panel { margin-top:16px; display:grid; gap:14px; }
.candidate-list { display:grid; gap:12px; }
.candidate-item { padding:12px; border:1px solid var(--c2m-border-light); border-radius:10px; background:#fbfdff; }
.candidate-value { margin:8px 0 0 24px; color:var(--c2m-text-secondary); white-space:pre-wrap; line-height:1.65; }
.candidate-location { margin-left:10px; color:var(--c2m-text-secondary); font-size:12px; }
.candidate-issues { padding:12px; border-radius:10px; background:#fffbe6; color:#8c6d1f; font-size:12px; line-height:1.7; }
.node-progress { display:grid; gap:0; margin:6px 0 20px; }
.node-step { display:flex; gap:12px; position:relative; padding:0 0 18px; }
.node-step:not(:last-child)::after { content:''; position:absolute; left:14px; top:30px; bottom:0; width:2px; background:var(--c2m-border-light); }
.node-marker { z-index:1; width:30px; height:30px; border-radius:50%; display:grid; place-items:center; background:#eef2f7; color:var(--c2m-text-secondary); font-size:12px; font-weight:700; flex:none; }
.node-completed .node-marker { background:#e6f7ee; color:#16834b; }
.node-running .node-marker { background:#e6f4ff; color:#1677ff; }
.node-awaiting_selection .node-marker { background:#fff4df; color:#d46b08; }
.node-failed .node-marker { background:#fff1f0; color:#cf1322; }
.node-copy { min-width:0; flex:1; }
.node-title { display:flex; align-items:center; gap:8px; color:var(--c2m-text-primary); }
.node-description, .node-message { color:var(--c2m-text-secondary); font-size:12px; margin-top:4px; }
@media (max-width: 900px) { .workspace-grid { grid-template-columns:1fr; } .topbar, .run-bar { flex-direction:column; } .form-row, .form-row.three, .upload-grid { grid-template-columns:1fr; } }
</style>
