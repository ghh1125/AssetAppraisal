import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import test from 'node:test'

const source = readFileSync(resolve(dirname(fileURLToPath(import.meta.url)), 'index.vue'), 'utf8')

test('upload controls keep v-for scope available to the conditional renderer', () => {
  assert.match(source, /<template v-for="field in uploadFields"[^>]*>/)
  assert.match(source, /<a-upload-dragger\s+v-if="showUploadField\(field\)"/)
  assert.doesNotMatch(source, /<a-upload-dragger[\s\S]*v-for="field in uploadFields"[\s\S]*v-if="showUploadField\(field\)"/)
})

test('all Ant Design controls used by the template are registered', () => {
  const imports = source.match(/from 'ant-design-vue'/)?.[0] || ''
  const main = readFileSync(resolve(dirname(fileURLToPath(import.meta.url)), '../../main.js'), 'utf8')
  for (const component of ['DatePicker', 'Divider', 'Modal', 'Radio']) {
    assert.match(main, new RegExp(`\\b${component}\\b`), `${component} must be imported and registered`)
  }
})

test('shows one current progress summary instead of duplicating the active step below it', () => {
  assert.match(source, /class="progress-overview-label"/)
  assert.doesNotMatch(source, /progress-strip/)
})

test('opens node one input sections in compact dialogs and hides file-only sources when API is selected', () => {
  assert.match(source, /node1-card/)
  assert.match(source, /v-model:open="manualModalOpen"/)
  assert.match(source, /v-model:open="materialsModalOpen"/)
  assert.match(source, /openNode1Section\('manual'\)/)
  assert.match(source, /openNode1Section\('materials'\)/)
  assert.match(source, /loadManualDefaults/)
  assert.match(source, /manualDefaultsButton/)
  assert.match(source, /上海上大热处理有限公司/)
  assert.match(source, /通富热处理（昆山）有限公司/)
  assert.doesNotMatch(source, /v-model:value="node1Sections"/)
  assert.match(source, /v-if="showUploadField\(field\)"/)
  assert.match(source, /form\.registry_info_strategy/)
})

test('renders compact candidate cards with persisted edit and feedback regeneration actions', () => {
  assert.match(source, /candidate-grid/)
  assert.match(source, /openCandidate\(candidate\)/)
  assert.match(source, /toggleCandidate\(candidate\.field_key\)/)
  assert.match(source, /saveCandidateEdit/)
  assert.match(source, /regenerateCandidate/)
  assert.match(source, /candidateFeedback/)
  assert.match(source, /candidate-card-title/)
})
