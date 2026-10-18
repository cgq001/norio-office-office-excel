import { readFile, rename, rm, writeFile } from 'node:fs/promises'
import { basename, resolve } from 'node:path'

import {
  createExcelStreamBuilder,
  createExcelStreamDecoder,
  getExcelToolDefinitions,
  validateExcelPatch,
  validateExcelWorkbook,
} from '@norio-office/office-excel/ai-protocol'
import { exportXlsx, importXlsx } from '@norio-office/office-excel/xlsx'

function parseArgs(argv) {
  const [command, ...rest] = argv
  const options = {}
  for (let index = 0; index < rest.length; index += 1) {
    const value = rest[index]
    if (!value.startsWith('--')) throw new Error(`Unexpected argument: ${value}`)
    const name = value.slice(2)
    if (name === 'stream' || name === 'force') options[name] = true
    else options[name] = rest[++index]
  }
  return { command, options }
}

async function readInput(input) {
  if (!input) throw new Error('--input is required')
  if (input === '-') {
    const chunks = []
    for await (const chunk of process.stdin) chunks.push(chunk)
    return Buffer.concat(chunks)
  }
  return readFile(resolve(input))
}

async function writeAtomic(output, bytes, force = false) {
  if (!output) throw new Error('--output is required')
  const destination = resolve(output)
  if (!force) {
    try {
      await readFile(destination)
      throw new Error(`Output already exists: ${destination}`)
    } catch (error) {
      if (error?.code !== 'ENOENT') throw error
    }
  }
  const temporary = `${destination}.${process.pid}.tmp`
  try {
    await writeFile(temporary, bytes)
    if (force) await rm(destination, { force: true })
    await rename(temporary, destination)
  } catch (error) {
    await rm(temporary, { force: true })
    throw error
  }
  return destination
}

function decodeStream(bytes) {
  const decoder = createExcelStreamDecoder()
  const events = [...decoder.append(bytes), ...decoder.end()]
  if (!events.length) throw new Error('Stream has no events')
  const creation = events.some((event) => event.type === 'sheet.open' || event.type === 'workbook.meta')
  const first = events[0]
  const builder = createExcelStreamBuilder({
    mode: creation ? 'create-workbook' : 'edit-range',
    streamId: first.streamId,
    documentId: 'offline-validation',
    requestId: 'offline-validation',
    contextId: 'offline-validation',
    baseRevision: 'offline-validation',
  })
  events.forEach((event) => builder.append(event))
  return builder.finalize()
}

async function readWorkbook(input) {
  const bytes = await readInput(input)
  if (/\.xlsx$/iu.test(input)) return (await importXlsx(bytes, { fileName: basename(input) })).workbook
  const workbook = JSON.parse(bytes.toString('utf8'))
  const result = validateExcelWorkbook(workbook)
  if (!result.valid) throw Object.assign(new Error('Workbook validation failed'), { issues: result.issues })
  return workbook
}

async function main() {
  const { command, options } = parseArgs(process.argv.slice(2))
  if (command === 'capabilities') {
    return { protocolVersion: 1, streamVersion: 1, workbookSchemaVersion: 1, fileFormats: { read: ['xlsx'], write: ['xlsx'] }, tools: getExcelToolDefinitions().map((tool) => tool.name) }
  }
  if (command === 'validate-workbook') {
    const workbook = await readWorkbook(options.input)
    return { valid: true, title: workbook.title, sheets: workbook.sheets.length }
  }
  if (command === 'validate-patch') {
    const result = validateExcelPatch(JSON.parse((await readInput(options.input)).toString('utf8')))
    if (!result.valid) throw Object.assign(new Error('Patch validation failed'), { issues: result.issues })
    return result
  }
  if (command === 'validate-stream') {
    const snapshot = decodeStream(await readInput(options.input))
    return { valid: true, mode: snapshot.mode, eventCount: snapshot.lastSeq, ended: snapshot.ended }
  }
  if (command === 'import-xlsx') {
    const input = await readInput(options.input)
    const imported = await importXlsx(input, { fileName: basename(options.input) })
    const output = await writeAtomic(options.output, `${JSON.stringify(imported.workbook, null, 2)}\n`, options.force)
    return { ok: true, output, compatibility: imported.compatibility }
  }
  if (command === 'create' || command === 'export-xlsx') {
    const workbook = options.stream
      ? decodeStream(await readInput(options.input)).workbook
      : await readWorkbook(options.input)
    if (!workbook) throw new Error('Stream did not create a workbook')
    const exported = await exportXlsx(workbook, { fileName: options.output })
    const output = await writeAtomic(options.output, exported.bytes, options.force)
    return { ok: true, output, size: exported.size, mimeType: exported.mimeType, compatibility: exported.compatibility }
  }
  throw new Error('Usage: excel.mjs <capabilities|validate-workbook|validate-patch|validate-stream|create|export-xlsx|import-xlsx> [--input path|-] [--output path] [--stream] [--force]')
}

try {
  const result = await main()
  process.stdout.write(`${JSON.stringify(result)}\n`)
} catch (error) {
  process.stderr.write(`${JSON.stringify({ ok: false, code: error?.code || 'CLI_ERROR', message: error?.message || String(error), issues: error?.issues })}\n`)
  process.exitCode = 1
}
