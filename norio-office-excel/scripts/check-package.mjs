import { access, readFile } from 'node:fs/promises'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const skillRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const manifest = JSON.parse(await readFile(resolve(skillRoot, 'runtime-manifest.json'), 'utf8'))
const [protocol, ai, xlsx] = await Promise.all([
  import('@norio-office/office-excel/ai-protocol'),
  import('@norio-office/office-excel/ai'),
  import('@norio-office/office-excel/xlsx'),
])

for (const relativePath of manifest.files) await access(resolve(skillRoot, relativePath))
const filterExamples = JSON.parse(await readFile(resolve(skillRoot, 'examples/filter-operations.json'), 'utf8'))
const invalidFilterExamples = filterExamples.examples
  .map((example) => ({ name: example.name, result: protocol.validateExcelPatch(example.patch) }))
  .filter((example) => !example.result.valid)
const checks = {
  protocolVersion: protocol.OFFICE_EXCEL_AI_PROTOCOL_VERSION === manifest.protocolVersion,
  streamVersion: protocol.OFFICE_EXCEL_AI_STREAM_VERSION === manifest.streamVersion,
  workbookSchemaVersion: protocol.OFFICE_EXCEL_WORKBOOK_SCHEMA_VERSION === manifest.workbookSchemaVersion,
  bridge: typeof ai.createExcelAiBridge === 'function',
  handlers: typeof ai.createExcelToolHandlers === 'function',
  importXlsx: typeof xlsx.importXlsx === 'function',
  exportXlsx: typeof xlsx.exportXlsx === 'function',
  captureTargetTool: protocol.getExcelToolDefinitions().some((tool) => tool.name === 'excel_capture_target'),
  filterExamples: filterExamples.examples.length >= 10 && invalidFilterExamples.length === 0,
}
if (Object.values(checks).some((value) => !value)) {
  throw new Error(`Public package compatibility check failed: ${JSON.stringify({ checks, invalidFilterExamples })}`)
}
process.stdout.write(`${JSON.stringify({ ok: true, checks, fileCount: manifest.files.length })}\n`)
