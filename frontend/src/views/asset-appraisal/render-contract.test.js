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
  for (const component of ['DatePicker', 'Divider', 'Radio']) {
    assert.match(main, new RegExp(`\\b${component}\\b`), `${component} must be imported and registered`)
  }
})

test('shows one current progress summary instead of duplicating the active step below it', () => {
  assert.match(source, /class="progress-overview-label"/)
  assert.doesNotMatch(source, /progress-strip/)
})
