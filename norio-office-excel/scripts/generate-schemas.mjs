import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

import {
  excelPatchJsonSchema,
  excelStreamEventJsonSchema,
  excelWorkbookJsonSchema,
  getExcelToolDefinitions,
} from '@norio-office/office-excel/ai-protocol'

const skillRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..')
const schemaRoot = resolve(skillRoot, 'schemas')
const resourceMapSchema = {
  $schema: 'https://json-schema.org/draft/2020-12/schema',
  title: 'Office Excel authorized resource map',
  type: 'object',
  additionalProperties: false,
  required: ['resources'],
  properties: {
    resources: {
      type: 'array',
      items: {
        type: 'object',
        additionalProperties: false,
        required: ['assetId', 'url'],
        properties: {
          assetId: { type: 'string', minLength: 1 },
          url: { type: 'string', minLength: 1 },
          name: { type: 'string' },
          mimeType: { type: 'string' },
          width: { type: 'number', minimum: 0 },
          height: { type: 'number', minimum: 0 },
        },
      },
    },
  },
}
const toolInputsSchema = {
  $schema: 'https://json-schema.org/draft/2020-12/schema',
  title: 'Office Excel AI tool inputs',
  type: 'object',
  additionalProperties: false,
  properties: Object.fromEntries(getExcelToolDefinitions().map((tool) => [tool.name, tool.inputSchema])),
}

const schemas = {
  'workbook.schema.json': excelWorkbookJsonSchema,
  'resource-map.schema.json': resourceMapSchema,
  'patch.schema.json': excelPatchJsonSchema,
  'stream-event.schema.json': excelStreamEventJsonSchema,
  'tool-inputs.schema.json': toolInputsSchema,
}
const write = process.argv.includes('--write')
if (!write && !process.argv.includes('--check')) throw new Error('Use --write or --check')
if (write) await mkdir(schemaRoot, { recursive: true })
for (const [fileName, schema] of Object.entries(schemas)) {
  const path = resolve(schemaRoot, fileName)
  const content = `${JSON.stringify(schema, null, 2)}\n`
  if (write) await writeFile(path, content, 'utf8')
  else if (await readFile(path, 'utf8') !== content) throw new Error(`${fileName} is out of sync; run npm run skill:sync`)
}
process.stdout.write(`${write ? 'Generated' : 'Checked'} ${Object.keys(schemas).length} schemas\n`)
