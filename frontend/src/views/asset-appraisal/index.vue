<script setup>
import { computed, onBeforeUnmount, reactive, ref } from 'vue'
import { message } from 'ant-design-vue'
import { useI18n } from 'vue-i18n'
import { artifactUrl } from '../../api/request'
import { checkAssetAppraisalOcrCache, createAssetAppraisalRun, getAssetAppraisalRun, regenerateAssetAppraisalCandidate, selectAssetAppraisalCandidates, updateAssetAppraisalCandidate } from '../../api/asset-appraisal'
import { canSubmitPartial } from '../../domain/submission'
import { summarizeRunIssues } from '../../domain/run-issues'
import { currentRunProgress } from '../../domain/run-progress'
import { createUploadState, uploadFields } from '../../domain/upload-fields'

const { t } = useI18n()

const form = reactive({
  commissioning_party_name: '',
  commissioning_party_short_name: '',
  transaction_type: '',
  target_company_name: '',
  target_company_short_name: '',
  valuation_subject_type: '',
  selected_valuation_method: [],
  final_valuation_method: '',
  valuation_base_date: '',
  registry_info_strategy: 'file',
  ownership_history_strategy: 'file',
  unrecorded_intangibles_strategy: 'file',
  company_profile_strategy: 'file',
})
const files = reactive(createUploadState())
const useGlm = ref(true)
const useQichacha = ref(true)
const reuseOcr = ref(true)
const ocrCache = ref({ checking: false, hit: false, source: '' })
const submitting = ref(false)
const run = ref(null)
const selectedCandidateKeys = ref([])
const candidateSelectionInitialized = ref(false)
const candidateModalOpen = ref(false)
const activeCandidate = ref(null)
const candidateDraft = ref('')
const candidateFeedback = ref('')
const candidateSaving = ref(false)
const candidateRegenerating = ref(false)
const manualModalOpen = ref(false)
const materialsModalOpen = ref(false)
const manualDraft = ref(null)
const RUN_STATUS_REFRESH_MS = 800
let pollTimer = null

const DEFAULT_MANUAL_INPUTS = Object.freeze({
  commissioning_party_name: '上海上大热处理有限公司',
  commissioning_party_short_name: '上海上大热处理',
  transaction_type: '收购',
  target_company_name: '通富热处理（昆山）有限公司',
  target_company_short_name: '通富昆山',
  valuation_subject_type: '股东全部权益价值',
  selected_valuation_method: ['收益法', '资产基础法'],
  final_valuation_method: '收益法',
  valuation_base_date: '2025-06-30',
})

const canSubmit = computed(() => canSubmitPartial(files, form))
const publicArtifacts = computed(() => (
  (run.value?.artifacts || []).filter(item => item.name === '资产评估报告_待复核.docx')
))
const readableIssues = computed(() => summarizeRunIssues(run.value?.issues || []))
const progressSummary = computed(() => currentRunProgress(run.value))
const statusText = computed(() => t(`asset.${run.value?.status || 'queued'}`))
const nodeStatusText = (status) => t(`asset.nodeStatus.${status || 'pending'}`)
const stepStatusText = (status) => t(`asset.nodeStatus.${status || 'pending'}`)
const manualFieldCount = computed(() => [
  form.commissioning_party_name,
  form.commissioning_party_short_name,
  form.transaction_type,
  form.target_company_name,
  form.target_company_short_name,
  form.valuation_subject_type,
  form.selected_valuation_method?.length,
  form.final_valuation_method,
  form.valuation_base_date,
].filter(Boolean).length)
const uploadedFileCount = computed(() => Object.values(files).reduce((count, value) => (
  count + (Array.isArray(value) ? value.length : value ? 1 : 0)
), 0))

function setFile(type, event) {
  const field = uploadFields.find(item => item.key === type)
  files[type] = field?.multiple
    ? (event.fileList || []).map(item => item.originFileObj).filter(Boolean)
    : event.fileList?.[0]?.originFileObj || null
  if (type === 'auditMaterials') {
    const pdf = files.auditMaterials.find(file => file?.name?.toLowerCase().endsWith('.pdf'))
    if (pdf) checkOcrCache(pdf)
    else ocrCache.value = { checking: false, hit: false, source: '' }
  }
}

function showUploadField(field) {
  return !field.sourceStrategy || form[field.sourceStrategy] === 'file'
}

function openNode1Section(section) {
  if (section === 'manual') {
    manualDraft.value = snapshotManualInputs()
    manualModalOpen.value = true
  }
  if (section === 'materials') materialsModalOpen.value = true
}

function snapshotManualInputs() {
  return {
    ...form,
    selected_valuation_method: [...(form.selected_valuation_method || [])],
  }
}

function loadManualDefaults() {
  Object.assign(form, {
    ...DEFAULT_MANUAL_INPUTS,
    selected_valuation_method: [...DEFAULT_MANUAL_INPUTS.selected_valuation_method],
  })
  message.success(t('asset.manualDefaultsLoaded'))
}

function saveManualSection() {
  manualModalOpen.value = false
  manualDraft.value = null
}

function cancelManualSection() {
  if (manualDraft.value) {
    Object.assign(form, {
      ...manualDraft.value,
      selected_valuation_method: [...(manualDraft.value.selected_valuation_method || [])],
    })
  }
  manualDraft.value = null
  manualModalOpen.value = false
}

function saveMaterialsSection() {
  materialsModalOpen.value = false
}

function openCandidate(candidate) {
  activeCandidate.value = candidate
  candidateDraft.value = String(candidate?.value || '')
  candidateFeedback.value = ''
  candidateModalOpen.value = true
}

function toggleCandidate(fieldKey) {
  selectedCandidateKeys.value = selectedCandidateKeys.value.includes(fieldKey)
    ? selectedCandidateKeys.value.filter(key => key !== fieldKey)
    : [...selectedCandidateKeys.value, fieldKey]
}

function patchCandidateInRun(updated) {
  if (!updated?.candidates || !run.value) return
  run.value = { ...run.value, candidates: updated.candidates }
  if (activeCandidate.value?.field_key) {
    activeCandidate.value = updated.candidates.find(
      candidate => candidate.field_key === activeCandidate.value.field_key,
    ) || activeCandidate.value
    candidateDraft.value = String(activeCandidate.value.value || '')
  }
}

async function saveCandidateEdit() {
  if (!run.value?.run_id || !activeCandidate.value?.field_key || !candidateDraft.value.trim()) {
    message.warning('候选内容不能为空')
    return
  }
  candidateSaving.value = true
  try {
    const updated = await updateAssetAppraisalCandidate(
      run.value.run_id,
      activeCandidate.value.field_key,
      candidateDraft.value.trim(),
    )
    patchCandidateInRun(updated)
    message.success('候选内容已保存')
  } catch (error) {
    message.error(error.message || '候选内容保存失败')
    throw error
  } finally {
    candidateSaving.value = false
  }
}

async function regenerateCandidate() {
  if (!run.value?.run_id || !activeCandidate.value?.field_key || !candidateFeedback.value.trim()) {
    message.warning('请先填写重新生成的反馈')
    return
  }
  candidateRegenerating.value = true
  try {
    run.value = await regenerateAssetAppraisalCandidate(
      run.value.run_id,
      activeCandidate.value.field_key,
      candidateFeedback.value.trim(),
    )
    candidateFeedback.value = ''
    await refreshRun(run.value.run_id)
  } catch (error) {
    message.error(error.message || '候选内容重新生成失败')
  } finally {
    candidateRegenerating.value = false
  }
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
    if (activeCandidate.value?.field_key) {
      const refreshedCandidate = (run.value.candidates || []).find(
        candidate => candidate.field_key === activeCandidate.value.field_key,
      )
      if (refreshedCandidate) {
        activeCandidate.value = refreshedCandidate
        candidateDraft.value = String(refreshedCandidate.value || '')
      }
    }
    if (['queued', 'running'].includes(run.value.status)) {
      pollTimer = window.setTimeout(() => refreshRun(runId), RUN_STATUS_REFRESH_MS)
    } else if (run.value.status === 'awaiting_selection') {
      const availableKeys = (run.value.candidates || [])
        .filter(item => item.available !== false)
        .map(item => item.field_key)
      if (!candidateSelectionInitialized.value) {
        selectedCandidateKeys.value = availableKeys
        candidateSelectionInitialized.value = true
      } else {
        selectedCandidateKeys.value = selectedCandidateKeys.value.filter(key => availableKeys.includes(key))
      }
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
  candidateSelectionInitialized.value = false
  try {
    const result = await createAssetAppraisalRun().mutationFn({
      auditMaterials: files.auditMaterials,
      reportingWorkbook: files.reportingWorkbook,
      incomeWorkbook: files.incomeWorkbook,
      registryMaterials: files.registryMaterials,
      ownershipHistoryMaterials: files.ownershipHistoryMaterials,
      unrecordedIntangiblesMaterials: files.unrecordedIntangiblesMaterials,
      companyProfileMaterials: files.companyProfileMaterials,
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
      <a-card class="panel node1-card" :title="t('asset.node1Title')" :bordered="false">
        <div class="node1-section-switch">
          <span class="node1-section-switch-label">{{ t('asset.node1SwitchHint') }}</span>
          <div class="node1-entry-grid">
            <button type="button" class="node1-entry" @click="openNode1Section('manual')">
              <span class="node1-entry-icon">✓</span>
              <span class="node1-entry-copy"><strong>{{ t('asset.manualSection') }}</strong><small>{{ manualFieldCount }}/9 项已填写</small></span>
              <span class="node1-entry-arrow">›</span>
            </button>
            <button type="button" class="node1-entry" @click="openNode1Section('materials')">
              <span class="node1-entry-icon">✓</span>
              <span class="node1-entry-copy"><strong>{{ t('asset.materialSection') }}</strong><small>{{ uploadedFileCount }} 个文件已选择</small></span>
              <span class="node1-entry-arrow">›</span>
            </button>
          </div>
        </div>

        <a-modal v-model:open="manualModalOpen" :title="t('asset.manualSection')" :ok-text="t('asset.saveSection')" :cancel-text="t('asset.closeSection')" :width="760" @ok="saveManualSection" @cancel="cancelManualSection">
          <div class="manual-default-bar">
            <span>{{ t('asset.manualDefaultsHint') }}</span>
            <a-button size="small" @click="loadManualDefaults">{{ t('asset.manualDefaultsButton') }}</a-button>
          </div>
          <a-form layout="vertical">
          <div class="form-row">
            <a-form-item :label="t('asset.commissioningName')"><a-input v-model:value="form.commissioning_party_name" :maxlength="50" /></a-form-item>
            <a-form-item :label="t('asset.commissioningShortName')"><a-input v-model:value="form.commissioning_party_short_name" :maxlength="20" /></a-form-item>
          </div>
          <a-form-item :label="t('asset.transaction')"><a-select v-model:value="form.transaction_type"><a-select-option value="转让">{{ t('asset.transactionOptions.transfer') }}</a-select-option><a-select-option value="收购">{{ t('asset.transactionOptions.acquisition') }}</a-select-option><a-select-option value="增资">{{ t('asset.transactionOptions.capitalIncrease') }}</a-select-option><a-select-option value="减资">{{ t('asset.transactionOptions.capitalDecrease') }}</a-select-option></a-select></a-form-item>
          <div class="form-row">
            <a-form-item :label="t('asset.targetName')"><a-input v-model:value="form.target_company_name" :maxlength="50" :placeholder="t('asset.targetNamePlaceholder')" /></a-form-item>
            <a-form-item :label="t('asset.targetShortName')"><a-input v-model:value="form.target_company_short_name" :maxlength="20" /></a-form-item>
          </div>
          <a-form-item :label="t('asset.subjectType')"><a-select v-model:value="form.valuation_subject_type"><a-select-option value="股东全部权益价值">股东全部权益价值</a-select-option><a-select-option value="股东部分权益价值">股东部分权益价值</a-select-option><a-select-option value="企业整体价值">企业整体价值</a-select-option><a-select-option value="资产组价值">资产组价值</a-select-option></a-select></a-form-item>
          <a-form-item :label="t('asset.method')"><a-select v-model:value="form.selected_valuation_method" mode="multiple" :max-tag-count="3"><a-select-option value="资产基础法">{{ t('asset.methodOptions.asset') }}</a-select-option><a-select-option value="收益法">{{ t('asset.methodOptions.income') }}</a-select-option><a-select-option value="市场法">{{ t('asset.methodOptions.market') }}</a-select-option></a-select></a-form-item>
          <a-form-item :label="t('asset.finalMethod')"><a-select v-model:value="form.final_valuation_method"><a-select-option value="资产基础法">{{ t('asset.methodOptions.asset') }}</a-select-option><a-select-option value="收益法">{{ t('asset.methodOptions.income') }}</a-select-option><a-select-option value="市场法">{{ t('asset.methodOptions.market') }}</a-select-option></a-select></a-form-item>
          <a-form-item :label="t('asset.valuationBaseDate')"><a-date-picker v-model:value="form.valuation_base_date" value-format="YYYY-MM-DD" style="width: 100%" /></a-form-item>
          </a-form>
        </a-modal>

        <a-modal v-model:open="materialsModalOpen" :title="t('asset.materialSection')" :ok-text="t('asset.saveSection')" :cancel-text="t('asset.closeSection')" :width="980" @ok="saveMaterialsSection">
          <a-form layout="vertical">
            <a-alert :message="t('asset.uploadInfo')" type="info" show-icon />
            <div class="source-strategy-grid">
              <a-form-item :label="t('asset.registryStrategy')"><a-radio-group v-model:value="form.registry_info_strategy"><a-radio value="file">{{ t('asset.sourceFile') }}</a-radio><a-radio value="qichacha">{{ t('asset.sourceQichacha') }}</a-radio></a-radio-group></a-form-item>
              <a-form-item :label="t('asset.ownershipStrategy')"><a-radio-group v-model:value="form.ownership_history_strategy"><a-radio value="file">{{ t('asset.sourceFile') }}</a-radio><a-radio value="qichacha">{{ t('asset.sourceQichacha') }}</a-radio></a-radio-group></a-form-item>
              <a-form-item :label="t('asset.intangiblesStrategy')"><a-radio-group v-model:value="form.unrecorded_intangibles_strategy"><a-radio value="file">{{ t('asset.sourceFile') }}</a-radio><a-radio value="qichacha">{{ t('asset.sourceQichacha') }}</a-radio></a-radio-group></a-form-item>
              <a-form-item :label="t('asset.profileStrategy')"><a-radio-group v-model:value="form.company_profile_strategy"><a-radio value="file">{{ t('asset.sourceFile') }}</a-radio><a-radio value="qichacha">{{ t('asset.sourceQichacha') }}</a-radio></a-radio-group></a-form-item>
            </div>
            <div class="upload-grid">
              <template v-for="field in uploadFields" :key="field.key">
                <a-upload-dragger
                  v-if="showUploadField(field)"
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
              </template>
            </div>
            <a-alert class="template-source" :message="t('asset.templateSource')" type="success" show-icon />
            <a-alert v-if="ocrCache.checking" class="ocr-cache-status" :message="t('asset.ocrCacheChecking')" type="info" show-icon />
            <a-alert v-else-if="ocrCache.hit" class="ocr-cache-status" :message="t('asset.ocrCacheHit', { source: ocrCache.source })" type="success" show-icon />
            <a-alert v-else-if="files.auditMaterials?.some(file => file?.name?.toLowerCase().endsWith('.pdf'))" class="ocr-cache-status" :message="t('asset.ocrCacheMiss')" type="warning" show-icon />
          </a-form>
        </a-modal>
      </a-card>
    </section>

    <section class="run-bar">
      <div><strong>{{ t('asset.generate') }}</strong><span>{{ t('asset.generateHint') }}</span></div>
      <a-button type="primary" size="large" :loading="submitting" :disabled="!canSubmit" @click="submit">{{ t('asset.start') }}</a-button>
    </section>

    <a-card v-if="run" class="panel result-panel" :title="t('asset.result')" :bordered="false">
      <div class="result-head"><div><span class="run-id">{{ t('asset.task') }} {{ run.run_id }}</span><a-tag :color="run.status === 'failed' ? 'red' : run.status === 'completed' ? 'green' : 'blue'">{{ statusText }}</a-tag></div><span v-if="run.message">{{ run.message }}</span></div>
      <div class="progress-overview">
        <div class="progress-overview-head"><span>当前工作流进度</span><strong>{{ run.progress || 0 }}%</strong></div>
        <div class="progress-overview-label">{{ progressSummary.label }}</div>
        <div class="progress-overview-detail">{{ progressSummary.detail }}</div>
      </div>
      <a-progress v-if="['queued', 'running'].includes(run.status)" :percent="run.progress || 0" status="active" />
      <div v-if="run.status === 'completed' && publicArtifacts.length" class="artifact-list result-artifact">
        <a :href="artifactUrl(run.run_id, artifact.name)" target="_blank" v-for="artifact in publicArtifacts" :key="artifact.name">{{ artifact.label || artifact.name }}</a>
      </div>
      <a-alert v-if="run.status === 'completed' && readableIssues.length" class="result-issues" :message="t('asset.reconciliationHint')" type="warning" show-icon>
        <template #description><div v-for="issue in readableIssues" :key="issue">{{ issue }}</div></template>
      </a-alert>
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
      <a-alert v-if="run.status === 'failed'" :message="run.error || t('asset.taskFailed')" type="error" show-icon />
      <div v-if="run.status === 'awaiting_selection'" class="candidate-panel">
        <a-alert :message="t('asset.candidateHint')" type="info" show-icon />
        <a-alert v-if="!run.candidates?.length" :message="t('asset.candidateEmpty')" type="warning" show-icon />
        <div v-if="!run.candidates?.length && readableIssues.length" class="candidate-issues">
          <strong>{{ t('asset.candidateIssues') }}</strong>
          <div v-for="issue in readableIssues" :key="issue">{{ issue }}</div>
        </div>
        <div class="candidate-grid">
          <article
            v-for="candidate in run.candidates"
            :key="candidate.field_key"
            :class="['candidate-card', { 'candidate-card-selected': selectedCandidateKeys.includes(candidate.field_key) }]"
          >
            <div class="candidate-card-head">
              <button type="button" class="candidate-card-title" @click="openCandidate(candidate)">
                {{ candidate.field_name || candidate.field_key }}
              </button>
              <a-button
                size="small"
                :disabled="candidate.available === false"
                :type="selectedCandidateKeys.includes(candidate.field_key) ? 'primary' : 'default'"
                @click="toggleCandidate(candidate.field_key)"
              >
                {{ selectedCandidateKeys.includes(candidate.field_key) ? '已选择' : '选择' }}
              </a-button>
            </div>
            <div v-if="candidate.location_ids?.length" class="candidate-location">{{ candidate.location_ids.join('、') }}</div>
            <div class="candidate-value">{{ candidate.available === false ? '暂无可用证据；点击标题查看详情。' : candidate.value }}</div>
            <button type="button" class="candidate-edit-link" @click="openCandidate(candidate)">查看、编辑或反馈重生成</button>
          </article>
        </div>
        <a-modal
          v-model:open="candidateModalOpen"
          :title="activeCandidate?.field_name || '候选内容'"
          :ok-text="'保存修改'"
          cancel-text="关闭"
          :width="820"
          :confirm-loading="candidateSaving"
          @ok="saveCandidateEdit"
        >
          <a-form layout="vertical">
            <a-form-item label="候选内容">
              <a-textarea v-model:value="candidateDraft" :rows="12" placeholder="可以直接修改候选内容" />
            </a-form-item>
            <a-form-item label="给 LLM 的反馈（可选）">
              <a-textarea v-model:value="candidateFeedback" :rows="4" placeholder="例如：补充产品应用场景，按行业、客户和竞争格局分段" />
            </a-form-item>
            <a-button type="dashed" :loading="candidateRegenerating" @click="regenerateCandidate">根据反馈重新生成当前模块</a-button>
          </a-form>
        </a-modal>
        <div class="candidate-selection-actions">
          <span>已选择 {{ selectedCandidateKeys.length }} / {{ run.candidates?.length || 0 }} 个模块</span>
          <a-button type="primary" :loading="submitting" @click="confirmCandidates">{{ t('asset.confirmCandidates') }}</a-button>
        </div>
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
.node1-card { grid-column:1 / -1; }
.panel { border-radius:18px; box-shadow:0 8px 30px rgba(31,53,81,.07); }
.node1-section-switch { display:grid; gap:12px; margin-bottom:4px; padding:12px 14px; border:1px solid #e5edf7; border-radius:12px; background:#f8fbff; }
.node1-section-switch-label { color:var(--c2m-text-secondary); font-size:13px; }
.node1-entry-grid { display:grid; grid-template-columns:repeat(2, minmax(0, 1fr)); gap:12px; }
.node1-entry { display:flex; align-items:center; gap:10px; width:100%; padding:14px 16px; border:1px solid #dbe7f4; border-radius:12px; background:#fff; color:var(--c2m-text-primary); text-align:left; cursor:pointer; transition:border-color .2s, box-shadow .2s, transform .2s; }
.node1-entry:hover { border-color:#91caff; box-shadow:0 5px 16px rgba(22,119,255,.12); transform:translateY(-1px); }
.node1-entry-icon { width:22px; height:22px; border-radius:6px; display:grid; place-items:center; flex:none; background:#e6f4ff; color:#1677ff; font-size:13px; font-weight:800; }
.node1-entry-copy { min-width:0; flex:1; }
.node1-entry-copy strong, .node1-entry-copy small { display:block; }
.node1-entry-copy small { margin-top:4px; color:var(--c2m-text-secondary); font-size:12px; }
.node1-entry-arrow { color:#8b98a8; font-size:24px; line-height:1; }
.node1-section { min-width:0; }
.section-heading { margin-bottom:14px; color:var(--c2m-text-primary); font-size:16px; font-weight:700; }
.manual-default-bar { display:flex; align-items:center; justify-content:space-between; gap:14px; margin-bottom:18px; padding:10px 12px; border-radius:10px; background:#f8fbff; color:var(--c2m-text-secondary); font-size:12px; }
.source-strategy-grid { display:grid; grid-template-columns:repeat(2, minmax(0, 1fr)); column-gap:18px; }
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
.progress-overview { margin:0 0 16px; padding:12px 14px; border:1px solid #e5edf7; border-radius:12px; background:#f8fbff; }
.progress-overview-head { display:flex; justify-content:space-between; align-items:center; color:var(--c2m-text-secondary); font-size:12px; }
.progress-overview-head strong { color:var(--c2m-color-primary); font-size:14px; }
.progress-overview-label { margin-top:6px; color:var(--c2m-text-primary); font-weight:650; }
.progress-overview-detail { margin-top:3px; color:var(--c2m-text-secondary); font-size:12px; white-space:pre-wrap; }
.artifact-list { display:flex; flex-wrap:wrap; gap:12px; }
.artifact-list a { padding:10px 14px; border:1px solid var(--c2m-border-light); border-radius:10px; color:var(--c2m-color-primary); background:#f8fbff; }
.result-artifact { margin-bottom:20px; }
.result-issues { margin-bottom:20px; white-space:pre-wrap; }
.candidate-panel { margin-top:16px; display:grid; gap:14px; }
.candidate-grid { display:grid; grid-template-columns:repeat(2, minmax(0, 1fr)); gap:12px; }
.candidate-card { min-width:0; padding:13px 14px; border:1px solid var(--c2m-border-light); border-radius:12px; background:#fbfdff; transition:border-color .2s, box-shadow .2s; }
.candidate-card-selected { border-color:#91caff; background:#f0f7ff; box-shadow:0 4px 14px rgba(22,119,255,.1); }
.candidate-card-head { display:flex; align-items:flex-start; justify-content:space-between; gap:10px; }
.candidate-card-title { padding:0; border:0; background:transparent; color:var(--c2m-text-primary); font-weight:700; line-height:1.45; text-align:left; cursor:pointer; }
.candidate-card-title:hover, .candidate-edit-link:hover { color:var(--c2m-color-primary); }
.candidate-value { max-height:96px; margin:9px 0 0; overflow:hidden; color:var(--c2m-text-secondary); white-space:pre-wrap; line-height:1.6; font-size:12px; }
.candidate-location { margin-top:6px; color:var(--c2m-text-secondary); font-size:11px; }
.candidate-edit-link { margin-top:9px; padding:0; border:0; background:transparent; color:var(--c2m-color-primary); font-size:12px; cursor:pointer; }
.candidate-selection-actions { display:flex; align-items:center; justify-content:space-between; gap:12px; color:var(--c2m-text-secondary); font-size:12px; }
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
.node-substeps { margin-top:10px; padding:10px 12px; border:1px solid #edf1f6; border-radius:12px; background:#fbfcfe; display:grid; gap:8px; }
.node-substep { display:flex; gap:8px; align-items:flex-start; color:var(--c2m-text-secondary); }
.substep-icon { width:18px; height:18px; border-radius:50%; display:grid; place-items:center; flex:none; font-size:11px; font-weight:700; background:#edf1f6; color:#8b98a8; }
.substep-copy { min-width:0; flex:1; }
.substep-title { display:flex; align-items:center; gap:8px; font-size:12px; color:var(--c2m-text-primary); }
.substep-status { color:#9aa6b2; font-size:11px; }
.substep-description { margin-top:2px; font-size:11px; line-height:1.45; color:#8b98a8; }
.substep-completed .substep-icon { background:#e6f7ee; color:#16834b; }
.substep-completed .substep-status { color:#16834b; }
.substep-running .substep-icon { background:#e6f4ff; color:#1677ff; animation:substep-pulse 1.2s infinite; }
.substep-running .substep-status { color:#1677ff; }
.substep-failed .substep-icon { background:#fff1f0; color:#cf1322; }
.substep-failed .substep-status { color:#cf1322; }
@keyframes substep-pulse { 50% { opacity:.45; transform:scale(.85); } }
@media (max-width: 900px) { .workspace-grid { grid-template-columns:1fr; } .topbar, .run-bar { flex-direction:column; } .form-row, .form-row.three, .source-strategy-grid, .upload-grid, .node1-entry-grid, .candidate-grid { grid-template-columns:1fr; } }
</style>
