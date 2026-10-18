import { mkdir, writeFile } from 'node:fs/promises'
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

await mkdir(schemaRoot, { recursive: true })
const schemas = {
  'workbook.schema.json': excelWorkbookJsonSchema,
  'resource-map.schema.json': resourceMapSchema,
  'patch.schema.json': excelPatchJsonSchema,
  'stream-event.schema.json': excelStreamEventJsonSchema,
  'tool-inputs.schema.json': toolInputsSchema,
}
for (const [fileName, schema] of Object.entries(schemas)) {
  await writeFile(resolve(schemaRoot, fileName), `${JSON.stringify(schema, null, 2)}\n`, 'utf8')
}
process.stdout.write(`${JSON.stringify({ ok: true, files: Object.keys(schemas) })}\n`)
