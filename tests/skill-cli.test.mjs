import assert from 'node:assert/strict'
import { spawn } from 'node:child_process'
import { mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { test } from 'node:test'

import { importXlsx } from '@norio-office/office-excel/xlsx'

const skill = resolve(dirname(fileURLToPath(import.meta.url)), '../norio-office-excel')
const script = join(skill, 'scripts/excel.mjs')

async function run(args, stdin) {
  const child = spawn(process.execPath, [script, ...args], { cwd: skill, stdio: ['pipe', 'pipe', 'pipe'] })
  let stdout = ''
  let stderr = ''
  child.stdout.setEncoding('utf8').on('data', (chunk) => { stdout += chunk })
  child.stderr.setEncoding('utf8').on('data', (chunk) => { stderr += chunk })
  child.stdin.end(stdin)
  const code = await new Promise((resolveCode) => child.on('close', resolveCode))
  return { code, stdout, stderr }
}

test('CLI reports public capabilities and validates fixtures', async () => {
  const capabilities = await run(['capabilities'])
  assert.equal(capabilities.code, 0, capabilities.stderr)
  const value = JSON.parse(capabilities.stdout)
  assert.equal(value.protocolVersion, 1)
  assert.ok(value.tools.includes('excel_apply_patch'))
  assert.equal((await run(['validate-workbook', '--input', 'examples/basic-workbook.json'])).code, 0)
  assert.equal((await run(['validate-patch', '--input', 'examples/header-style-patch.json'])).code, 0)
  assert.equal((await run(['validate-stream', '--input', 'examples/streamed-workbook.ndjson'])).code, 0)
  assert.equal((await run(['validate-stream', '--mode', 'edit-range', '--input', 'examples/streamed-range-edit.ndjson'])).code, 0)
})

test('CLI creates genuine XLSX and refuses to overwrite it', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'norio-excel-skill-'))
  const output = join(directory, 'result.xlsx')
  try {
    const created = await run(['create', '--input', 'examples/formulas-and-styles.json', '--output', output])
    assert.equal(created.code, 0, created.stderr)
    const first = await readFile(output)
    assert.equal(first[0], 0x50)
    assert.equal(first[1], 0x4b)
    const imported = await importXlsx(first, { fileName: 'result.xlsx' })
    assert.equal(imported.workbook.sheets[0].cells['3:2'].value, '=B2*(1-0.06)')
    assert.equal((await run(['validate-workbook', '--input', output])).code, 0)
    const duplicate = await run(['create', '--input', 'examples/basic-workbook.json', '--output', output])
    assert.notEqual(duplicate.code, 0)
    assert.deepEqual(await readFile(output), first)
  } finally {
    await rm(directory, { recursive: true, force: true })
  }
})

test('CLI creates from stdin stream and imports XLSX to a snapshot', async () => {
  const directory = await mkdtemp(join(tmpdir(), 'norio-excel-stream-'))
  const xlsx = join(directory, 'stream.xlsx')
  const json = join(directory, 'snapshot.json')
  try {
    const stream = await readFile(join(skill, 'examples/streamed-workbook.ndjson'))
    const created = await run(['create', '--stream', '--input', '-', '--output', xlsx], stream)
    assert.equal(created.code, 0, created.stderr)
    const imported = await run(['import-xlsx', '--input', xlsx, '--output', json])
    assert.equal(imported.code, 0, imported.stderr)
    const snapshot = JSON.parse(await readFile(json, 'utf8'))
    assert.equal(snapshot.title, 'stream')
    assert.equal(snapshot.sheets[0].name, '销售')
    assert.equal(snapshot.sheets[0].cells['2:2'].value, '12500')
  } finally {
    await rm(directory, { recursive: true, force: true })
  }
})

test('publish command refuses accidental execution without --skills', async () => {
  const root = resolve(skill, '../../..')
  const child = spawn(process.execPath, [join(root, 'tools/push-skills.mjs')], { cwd: root, stdio: ['ignore', 'pipe', 'pipe'] })
  let stderr = ''
  child.stderr.setEncoding('utf8').on('data', (chunk) => { stderr += chunk })
  const code = await new Promise((resolveCode) => child.on('close', resolveCode))
  assert.notEqual(code, 0)
  assert.match(stderr, /npm run push -- --skills/u)
})
