#!/usr/bin/env bun

import { readFile } from 'node:fs/promises'
import { decodeAgentFrame, encodeAgentFrame, type AgentFrame, type AgentResponse } from '../packages/agent-api/src/index.ts'
import type {
  CellDateStyle,
  CellNumberFormatInput,
  CellNumberFormatKind,
  CellNumberNegativeStyle,
  CellNumberZeroStyle,
  CellRangeRef,
  CellStyleField,
  CellStylePatch,
  LiteralInput,
  WorkbookPivotValueSnapshot,
} from '../packages/protocol/src/index.ts'

const [, , command, ...argv] = process.argv

type CommandName =
  | 'read-range'
  | 'write-cell'
  | 'write-range'
  | 'set-formula'
  | 'set-formulas'
  | 'set-range-style'
  | 'clear-range-style'
  | 'set-range-number-format'
  | 'clear-range-number-format'
  | 'clear-range'
  | 'create-pivot'
  | 'batch'
  | 'get-metrics'
  | 'export-snapshot'

type CliOptions = Record<string, string>
type SuccessAgentResponse = Exclude<AgentResponse, { kind: 'error' }>

function printUsage(): void {
  console.log(`Usage:
  bun scripts/spreadsheet-agent.ts read-range --range Sheet1!A1:B2 [--server URL] [--document ID] [--replica ID]
  bun scripts/spreadsheet-agent.ts write-cell --sheet Sheet1 --addr A1 --value 42 [--server URL] [--document ID] [--replica ID]
  bun scripts/spreadsheet-agent.ts write-range --range Sheet1!A1:B2 --values '[[1,2],[3,4]]'
  bun scripts/spreadsheet-agent.ts set-formula --sheet Sheet1 --addr B1 --formula 'SUM(A1:A10)'
  bun scripts/spreadsheet-agent.ts set-formulas --range Sheet1!B1:B2 --formulas '[["A1*2"],["A2*2"]]'
  bun scripts/spreadsheet-agent.ts set-range-style --range Sheet1!A1:C3 --patch '{"fill":{"backgroundColor":"#fff59d"},"font":{"family":"Georgia"}}'
  bun scripts/spreadsheet-agent.ts clear-range-style --range Sheet1!A1:C3 [--fields '["backgroundColor"]']
  bun scripts/spreadsheet-agent.ts set-range-number-format --range Sheet1!B2:B10 --format '{"kind":"accounting","currency":"USD","decimals":2}'
  bun scripts/spreadsheet-agent.ts clear-range-number-format --range Sheet1!B2:B10
  bun scripts/spreadsheet-agent.ts clear-range --range Sheet1!A1:B2
  bun scripts/spreadsheet-agent.ts create-pivot --name MyPivot --sheet Sheet1 --addr D1 --source Sheet2!A1:C100 --group '["Category"]' --values '[{"sourceColumn":"Amount","summarizeBy":"sum"}]'
  bun scripts/spreadsheet-agent.ts batch --requests @ops.json [--server URL] [--document ID] [--replica ID]
  bun scripts/spreadsheet-agent.ts get-metrics
  bun scripts/spreadsheet-agent.ts export-snapshot

JSON-heavy flags such as --values, --formulas, --group, --requests, and --value accept:
  inline JSON     '[["a","b"]]'
  @file.json      load JSON from a file
  @-              read JSON from stdin
`)
}

function parseArgs(args: readonly string[]): CliOptions {
  const options: CliOptions = {}
  for (let index = 0; index < args.length; index += 1) {
    const token = args[index]
    if (!token?.startsWith('--')) {
      throw new Error(`Unexpected argument: ${token ?? ''}`)
    }
    const key = token.slice(2)
    const value = args[index + 1]
    if (!value || value.startsWith('--')) {
      throw new Error(`Missing value for --${key}`)
    }
    options[key] = value
    index += 1
  }
  return options
}

function normalizeBaseUrl(value: string): string {
  return value.endsWith('/') ? value.slice(0, -1) : value
}

function parseJson(value: string, label: string): unknown {
  try {
    return JSON.parse(value) as unknown
  } catch (error) {
    throw new Error(`Invalid JSON for ${label}: ${error instanceof Error ? error.message : String(error)}`, {
      cause: error,
    })
  }
}

let stdinTextPromise: Promise<string> | null = null

async function readStdinText(): Promise<string> {
  if (stdinTextPromise) {
    return stdinTextPromise
  }
  stdinTextPromise = (async () => {
    const chunks: Buffer[] = []
    for await (const chunk of process.stdin) {
      chunks.push(typeof chunk === 'string' ? Buffer.from(chunk) : Buffer.from(chunk))
    }
    return Buffer.concat(chunks).toString('utf8')
  })()
  return stdinTextPromise
}

async function loadJsonArgument(value: string, label: string): Promise<unknown> {
  if (!value.startsWith('@')) {
    return parseJson(value, label)
  }
  const source = value.slice(1)
  if (!source) {
    throw new Error(`Missing JSON source for ${label}`)
  }
  const text = source === '-' ? await readStdinText() : await readFile(source, 'utf8')
  return parseJson(text, `${label} (${source === '-' ? 'stdin' : source})`)
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isLiteralInput(value: unknown): value is LiteralInput {
  return value === null || typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean'
}

function parseLiteralInput(value: unknown, label: string): LiteralInput {
  if (!isLiteralInput(value)) {
    throw new Error(`${label} must be a JSON literal (string, number, boolean, or null)`)
  }
  return value
}

function parseLiteralMatrix(value: unknown, label: string): LiteralInput[][] {
  if (!Array.isArray(value)) {
    throw new Error(`${label} must be a JSON matrix`)
  }
  return value.map((row, rowIndex) => {
    if (!Array.isArray(row)) {
      throw new Error(`${label} row ${rowIndex + 1} must be an array`)
    }
    return row.map((cell, cellIndex) => parseLiteralInput(cell, `${label}[${rowIndex}][${cellIndex}]`))
  })
}

function parseFormulaMatrix(value: unknown, label: string): string[][] {
  if (!Array.isArray(value)) {
    throw new Error(`${label} must be a JSON matrix`)
  }
  return value.map((row, rowIndex) => {
    if (!Array.isArray(row)) {
      throw new Error(`${label} row ${rowIndex + 1} must be an array`)
    }
    return row.map((cell, cellIndex) => {
      if (typeof cell !== 'string') {
        throw new Error(`${label}[${rowIndex}][${cellIndex}] must be a string`)
      }
      return cell
    })
  })
}

function parseStringArray(value: unknown, label: string): string[] {
  if (!Array.isArray(value)) {
    throw new Error(`${label} must be a JSON array`)
  }
  return value.map((entry, index) => {
    if (typeof entry !== 'string') {
      throw new Error(`${label}[${index}] must be a string`)
    }
    return entry
  })
}

function parsePivotValues(value: unknown, label: string): WorkbookPivotValueSnapshot[] {
  if (!Array.isArray(value)) {
    throw new Error(`${label} must be a JSON array`)
  }
  return value.map((entry, index) => {
    if (!isRecord(entry) || typeof entry.sourceColumn !== 'string' || typeof entry.summarizeBy !== 'string') {
      throw new Error(`${label}[${index}] must include string sourceColumn and summarizeBy fields`)
    }
    return {
      sourceColumn: entry.sourceColumn,
      summarizeBy: entry.summarizeBy,
    } satisfies WorkbookPivotValueSnapshot
  })
}

function parseStylePatch(value: unknown, label: string): CellStylePatch {
  if (!isRecord(value)) {
    throw new Error(`${label} must be an object`)
  }
  const patch: CellStylePatch = {}
  if ('fill' in value) {
    const fill = value.fill
    if (!isRecord(fill)) {
      throw new Error(`${label}.fill must be an object`)
    }
    if ('backgroundColor' in fill && typeof fill.backgroundColor !== 'string') {
      throw new Error(`${label}.fill.backgroundColor must be a string`)
    }
    patch.fill = {}
    if (typeof fill.backgroundColor === 'string') {
      patch.fill.backgroundColor = fill.backgroundColor
    }
  }
  if ('font' in value) {
    const font = value.font
    if (!isRecord(font)) {
      throw new Error(`${label}.font must be an object`)
    }
    patch.font = {}
    if (typeof font.family === 'string' || font.family === null) {
      patch.font.family = font.family
    }
    if (typeof font.size === 'number' || font.size === null) {
      patch.font.size = font.size
    }
    if (typeof font.bold === 'boolean' || font.bold === null) {
      patch.font.bold = font.bold
    }
    if (typeof font.italic === 'boolean' || font.italic === null) {
      patch.font.italic = font.italic
    }
    if (typeof font.underline === 'boolean' || font.underline === null) {
      patch.font.underline = font.underline
    }
    if (typeof font.color === 'string' || font.color === null) {
      patch.font.color = font.color
    }
  }
  if ('alignment' in value) {
    const alignment = value.alignment
    if (!isRecord(alignment)) {
      throw new Error(`${label}.alignment must be an object`)
    }
    patch.alignment = {}
    if (
      alignment.horizontal === null ||
      alignment.horizontal === 'general' ||
      alignment.horizontal === 'left' ||
      alignment.horizontal === 'center' ||
      alignment.horizontal === 'right'
    ) {
      patch.alignment.horizontal = alignment.horizontal
    }
    if (alignment.vertical === null || alignment.vertical === 'top' || alignment.vertical === 'middle' || alignment.vertical === 'bottom') {
      patch.alignment.vertical = alignment.vertical
    }
    if (typeof alignment.wrap === 'boolean' || alignment.wrap === null) {
      patch.alignment.wrap = alignment.wrap
    }
    if (typeof alignment.indent === 'number' || alignment.indent === null) {
      patch.alignment.indent = alignment.indent
    }
  }
  if ('borders' in value) {
    const borders = value.borders
    if (!isRecord(borders)) {
      throw new Error(`${label}.borders must be an object`)
    }
    patch.borders = {}
    for (const side of ['top', 'right', 'bottom', 'left'] as const) {
      const border = borders[side]
      if (border === null) {
        patch.borders[side] = null
        continue
      }
      if (!isRecord(border)) {
        continue
      }
      patch.borders[side] = {}
      if (
        border.style === null ||
        border.style === 'solid' ||
        border.style === 'dashed' ||
        border.style === 'dotted' ||
        border.style === 'double'
      ) {
        patch.borders[side].style = border.style
      }
      if (border.weight === null || border.weight === 'thin' || border.weight === 'medium' || border.weight === 'thick') {
        patch.borders[side].weight = border.weight
      }
      if (typeof border.color === 'string' || border.color === null) {
        patch.borders[side].color = border.color
      }
    }
  }
  return patch
}

function parseStyleFields(value: unknown, label: string): CellStyleField[] {
  if (!Array.isArray(value)) {
    throw new Error(`${label} must be a JSON array`)
  }
  return value.map((entry, index) => {
    if (
      entry !== 'backgroundColor' &&
      entry !== 'fontFamily' &&
      entry !== 'fontSize' &&
      entry !== 'fontBold' &&
      entry !== 'fontItalic' &&
      entry !== 'fontUnderline' &&
      entry !== 'fontColor' &&
      entry !== 'alignmentHorizontal' &&
      entry !== 'alignmentVertical' &&
      entry !== 'alignmentWrap' &&
      entry !== 'alignmentIndent' &&
      entry !== 'borderTop' &&
      entry !== 'borderRight' &&
      entry !== 'borderBottom' &&
      entry !== 'borderLeft'
    ) {
      throw new Error(`${label}[${index}] is not a supported style field`)
    }
    return entry
  })
}

function parseNumberFormatInput(value: unknown, label: string): CellNumberFormatInput {
  if (typeof value === 'string') {
    return value
  }
  if (!isRecord(value) || typeof value.kind !== 'string') {
    throw new Error(`${label} must be a string or number-format preset object`)
  }

  const preset = {
    kind: parseNumberFormatKind(value.kind, `${label}.kind`),
    ...(typeof value.currency === 'string' ? { currency: value.currency } : {}),
    ...(typeof value.decimals === 'number' ? { decimals: value.decimals } : {}),
    ...(typeof value.useGrouping === 'boolean' ? { useGrouping: value.useGrouping } : {}),
    ...(value.negativeStyle !== undefined ? { negativeStyle: parseNegativeStyle(value.negativeStyle, `${label}.negativeStyle`) } : {}),
    ...(value.zeroStyle !== undefined ? { zeroStyle: parseZeroStyle(value.zeroStyle, `${label}.zeroStyle`) } : {}),
    ...(value.dateStyle !== undefined ? { dateStyle: parseDateStyle(value.dateStyle, `${label}.dateStyle`) } : {}),
  }
  return preset
}

function parseNumberFormatKind(value: unknown, label: string): CellNumberFormatKind {
  switch (value) {
    case 'general':
    case 'number':
    case 'currency':
    case 'accounting':
    case 'percent':
    case 'date':
    case 'time':
    case 'datetime':
    case 'text':
      return value
    default:
      throw new Error(`${label} is not a supported number format kind`)
  }
}

function parseNegativeStyle(value: unknown, label: string): CellNumberNegativeStyle {
  switch (value) {
    case 'minus':
    case 'parentheses':
      return value
    default:
      throw new Error(`${label} is not a supported negative style`)
  }
}

function parseZeroStyle(value: unknown, label: string): CellNumberZeroStyle {
  switch (value) {
    case 'zero':
    case 'dash':
      return value
    default:
      throw new Error(`${label} is not a supported zero style`)
  }
}

function parseDateStyle(value: unknown, label: string): CellDateStyle {
  switch (value) {
    case 'short':
    case 'iso':
      return value
    default:
      throw new Error(`${label} is not a supported date style`)
  }
}

function parseRange(value: string, fallbackSheet?: string): CellRangeRef {
  const [sheetAndStart, endAddress] = value.includes(':') ? value.split(':') : [value, value]
  const bangIndex = sheetAndStart.indexOf('!')
  if (bangIndex >= 0) {
    return {
      sheetName: sheetAndStart.slice(0, bangIndex),
      startAddress: sheetAndStart.slice(bangIndex + 1),
      endAddress,
    }
  }
  if (!fallbackSheet) {
    throw new Error('Range must include a sheet name or use --sheet')
  }
  return {
    sheetName: fallbackSheet,
    startAddress: sheetAndStart,
    endAddress,
  }
}

function parseRangeLike(value: string | CellRangeRef | undefined, fallbackSheet?: string): CellRangeRef {
  if (typeof value === 'string') {
    return parseRange(value, fallbackSheet)
  }
  if (!isRecord(value)) {
    throw new Error('Range must be a string like Sheet1!A1:B2 or an object with sheetName/startAddress/endAddress')
  }
  const sheetName = typeof value.sheetName === 'string' ? value.sheetName : fallbackSheet
  const startAddress = typeof value.startAddress === 'string' ? value.startAddress : null
  const endAddress = typeof value.endAddress === 'string' ? value.endAddress : startAddress
  if (!sheetName || !startAddress || !endAddress) {
    throw new Error('Range object requires sheetName/startAddress/endAddress')
  }
  return { sheetName, startAddress, endAddress }
}

function parseRangeInput(value: unknown): string | CellRangeRef | undefined {
  if (typeof value === 'string') {
    return value
  }
  if (!isRecord(value)) {
    return undefined
  }
  if (typeof value.sheetName !== 'string' || typeof value.startAddress !== 'string') {
    return undefined
  }
  return {
    sheetName: value.sheetName,
    startAddress: value.startAddress,
    endAddress: typeof value.endAddress === 'string' ? value.endAddress : value.startAddress,
  }
}

async function sendFrame(serverBaseUrl: string, frame: AgentFrame): Promise<SuccessAgentResponse> {
  const response = await fetch(`${normalizeBaseUrl(serverBaseUrl)}/v2/agent/frames`, {
    method: 'POST',
    headers: {
      'content-type': 'application/octet-stream',
    },
    body: Buffer.from(encodeAgentFrame(frame)),
  })
  if (!response.ok) {
    throw new Error(`Agent request failed with status ${response.status}`)
  }
  const nextFrame = decodeAgentFrame(new Uint8Array(await response.arrayBuffer()))
  if (nextFrame.kind !== 'response') {
    throw new Error(`Expected response frame, received ${nextFrame.kind}`)
  }
  if (nextFrame.response.kind === 'error') {
    throw new Error(`${nextFrame.response.code}: ${nextFrame.response.message}`)
  }
  return nextFrame.response
}

async function runCommand(commandName: CommandName, options: CliOptions & { server: string }, sessionId: string): Promise<unknown> {
  switch (commandName) {
    case 'read-range':
      return sendFrame(options.server, {
        kind: 'request',
        request: {
          kind: 'readRange',
          id: `read:${Date.now()}`,
          sessionId,
          range: parseRange(options.range ?? '', options.sheet),
        },
      })
    case 'write-cell':
      return sendFrame(options.server, {
        kind: 'request',
        request: {
          kind: 'writeRange',
          id: `write-cell:${Date.now()}`,
          sessionId,
          range: {
            sheetName: options.sheet ?? '',
            startAddress: options.addr ?? '',
            endAddress: options.addr ?? '',
          },
          values: [[parseLiteralInput(await loadJsonArgument(options.value ?? '', '--value'), '--value')]],
        },
      })
    case 'write-range':
      return sendFrame(options.server, {
        kind: 'request',
        request: {
          kind: 'writeRange',
          id: `write-range:${Date.now()}`,
          sessionId,
          range: parseRange(options.range ?? '', options.sheet),
          values: parseLiteralMatrix(await loadJsonArgument(options.values ?? '', '--values'), '--values'),
        },
      })
    case 'set-formula':
      return sendFrame(options.server, {
        kind: 'request',
        request: {
          kind: 'setRangeFormulas',
          id: `set-formula:${Date.now()}`,
          sessionId,
          range: {
            sheetName: options.sheet ?? '',
            startAddress: options.addr ?? '',
            endAddress: options.addr ?? '',
          },
          formulas: [[options.formula ?? '']],
        },
      })
    case 'set-formulas':
      return sendFrame(options.server, {
        kind: 'request',
        request: {
          kind: 'setRangeFormulas',
          id: `set-formulas:${Date.now()}`,
          sessionId,
          range: parseRange(options.range ?? '', options.sheet),
          formulas: parseFormulaMatrix(await loadJsonArgument(options.formulas ?? '', '--formulas'), '--formulas'),
        },
      })
    case 'set-range-style':
      return sendFrame(options.server, {
        kind: 'request',
        request: {
          kind: 'setRangeStyle',
          id: `set-range-style:${Date.now()}`,
          sessionId,
          range: parseRange(options.range ?? '', options.sheet),
          patch: parseStylePatch(await loadJsonArgument(options.patch ?? '', '--patch'), '--patch'),
        },
      })
    case 'clear-range-style':
      return sendFrame(options.server, {
        kind: 'request',
        request: {
          kind: 'clearRangeStyle',
          id: `clear-range-style:${Date.now()}`,
          sessionId,
          range: parseRange(options.range ?? '', options.sheet),
          fields: options.fields ? parseStyleFields(await loadJsonArgument(options.fields, '--fields'), '--fields') : undefined,
        },
      })
    case 'set-range-number-format':
      return sendFrame(options.server, {
        kind: 'request',
        request: {
          kind: 'setRangeNumberFormat',
          id: `set-range-number-format:${Date.now()}`,
          sessionId,
          range: parseRange(options.range ?? '', options.sheet),
          format: parseNumberFormatInput(await loadJsonArgument(options.format ?? '', '--format'), '--format'),
        },
      })
    case 'clear-range-number-format':
      return sendFrame(options.server, {
        kind: 'request',
        request: {
          kind: 'clearRangeNumberFormat',
          id: `clear-range-number-format:${Date.now()}`,
          sessionId,
          range: parseRange(options.range ?? '', options.sheet),
        },
      })
    case 'clear-range':
      return sendFrame(options.server, {
        kind: 'request',
        request: {
          kind: 'clearRange',
          id: `clear-range:${Date.now()}`,
          sessionId,
          range: parseRange(options.range ?? '', options.sheet),
        },
      })
    case 'get-metrics':
      return sendFrame(options.server, {
        kind: 'request',
        request: {
          kind: 'getMetrics',
          id: `get-metrics:${Date.now()}`,
          sessionId,
        },
      })
    case 'export-snapshot':
      return sendFrame(options.server, {
        kind: 'request',
        request: {
          kind: 'exportSnapshot',
          id: `export-snapshot:${Date.now()}`,
          sessionId,
        },
      })
    case 'create-pivot':
      return sendFrame(options.server, {
        kind: 'request',
        request: {
          kind: 'createPivotTable',
          id: `create-pivot:${Date.now()}`,
          sessionId,
          name: options.name ?? '',
          sheetName: options.sheet ?? '',
          address: options.addr ?? '',
          source: parseRange(options.source ?? '', options.sheet),
          groupBy: parseStringArray(await loadJsonArgument(options.group ?? '', '--group'), '--group'),
          values: parsePivotValues(await loadJsonArgument(options.values ?? '', '--values'), '--values'),
        },
      })
    case 'batch': {
      const requests = await loadJsonArgument(options.requests ?? '', '--requests')
      if (!Array.isArray(requests)) {
        throw new Error('--requests must be a JSON array')
      }
      return runBatchRequests(requests, options, sessionId)
    }
  }
}

async function dispatchBatchRequest(
  request: unknown,
  index: number,
  options: CliOptions & { server: string },
  sessionId: string,
): Promise<SuccessAgentResponse> {
  if (!isRecord(request) || typeof request.kind !== 'string') {
    throw new Error(`Batch request at index ${index} must be an object with a kind`)
  }
  const id = typeof request.id === 'string' ? request.id : `batch:${index + 1}:${Date.now()}`
  switch (request.kind) {
    case 'readRange':
      return sendFrame(options.server, {
        kind: 'request',
        request: {
          kind: 'readRange',
          id,
          sessionId,
          range: parseRangeLike(parseRangeInput(request.range), options.sheet),
        },
      })
    case 'writeRange':
      return sendFrame(options.server, {
        kind: 'request',
        request: {
          kind: 'writeRange',
          id,
          sessionId,
          range: parseRangeLike(parseRangeInput(request.range), options.sheet),
          values: parseLiteralMatrix(request.values, `--requests[${index}].values`),
        },
      })
    case 'setRangeFormulas':
      return sendFrame(options.server, {
        kind: 'request',
        request: {
          kind: 'setRangeFormulas',
          id,
          sessionId,
          range: parseRangeLike(parseRangeInput(request.range), options.sheet),
          formulas: parseFormulaMatrix(request.formulas, `--requests[${index}].formulas`),
        },
      })
    case 'setRangeStyle':
      return sendFrame(options.server, {
        kind: 'request',
        request: {
          kind: 'setRangeStyle',
          id,
          sessionId,
          range: parseRangeLike(parseRangeInput(request.range), options.sheet),
          patch: parseStylePatch(request.patch, `--requests[${index}].patch`),
        },
      })
    case 'clearRangeStyle':
      return sendFrame(options.server, {
        kind: 'request',
        request: {
          kind: 'clearRangeStyle',
          id,
          sessionId,
          range: parseRangeLike(parseRangeInput(request.range), options.sheet),
          fields: request.fields === undefined ? undefined : parseStyleFields(request.fields, `--requests[${index}].fields`),
        },
      })
    case 'setRangeNumberFormat':
      return sendFrame(options.server, {
        kind: 'request',
        request: {
          kind: 'setRangeNumberFormat',
          id,
          sessionId,
          range: parseRangeLike(parseRangeInput(request.range), options.sheet),
          format: parseNumberFormatInput(request.format, `--requests[${index}].format`),
        },
      })
    case 'clearRangeNumberFormat':
      return sendFrame(options.server, {
        kind: 'request',
        request: {
          kind: 'clearRangeNumberFormat',
          id,
          sessionId,
          range: parseRangeLike(parseRangeInput(request.range), options.sheet),
        },
      })
    case 'clearRange':
      return sendFrame(options.server, {
        kind: 'request',
        request: {
          kind: 'clearRange',
          id,
          sessionId,
          range: parseRangeLike(parseRangeInput(request.range), options.sheet),
        },
      })
    case 'getMetrics':
      return sendFrame(options.server, {
        kind: 'request',
        request: {
          kind: 'getMetrics',
          id,
          sessionId,
        },
      })
    case 'exportSnapshot':
      return sendFrame(options.server, {
        kind: 'request',
        request: {
          kind: 'exportSnapshot',
          id,
          sessionId,
        },
      })
    case 'createPivotTable':
      if (typeof request.name !== 'string' || typeof request.sheetName !== 'string' || typeof request.address !== 'string') {
        throw new Error(`--requests[${index}] createPivotTable requires name, sheetName, and address`)
      }
      return sendFrame(options.server, {
        kind: 'request',
        request: {
          kind: 'createPivotTable',
          id,
          sessionId,
          name: request.name,
          sheetName: request.sheetName,
          address: request.address,
          source: parseRangeLike(parseRangeInput(request.source), options.sheet),
          groupBy: parseStringArray(request.groupBy, `--requests[${index}].groupBy`),
          values: parsePivotValues(request.values, `--requests[${index}].values`),
        },
      })
    default:
      throw new Error(`Unsupported batch request kind: ${request.kind}`)
  }
}

async function runBatchRequests(
  requests: unknown[],
  options: CliOptions & { server: string },
  sessionId: string,
  index = 0,
  responses: SuccessAgentResponse[] = [],
): Promise<SuccessAgentResponse[]> {
  if (index >= requests.length) {
    return responses
  }
  responses.push(await dispatchBatchRequest(requests[index], index, options, sessionId))
  return runBatchRequests(requests, options, sessionId, index + 1, responses)
}

function isCommandName(value: string): value is CommandName {
  return [
    'read-range',
    'write-cell',
    'write-range',
    'set-formula',
    'set-formulas',
    'set-range-style',
    'clear-range-style',
    'set-range-number-format',
    'clear-range-number-format',
    'clear-range',
    'create-pivot',
    'batch',
    'get-metrics',
    'export-snapshot',
  ].includes(value)
}

async function main(): Promise<void> {
  if (!command || command === '--help' || command === 'help') {
    printUsage()
    return
  }
  if (!isCommandName(command)) {
    throw new Error(`Unsupported command: ${command}`)
  }

  const options = parseArgs(argv)
  const server = options.server ?? process.env.BILIG_AGENT_SERVER_URL ?? 'http://127.0.0.1:4321'
  const documentId = options.document ?? process.env.BILIG_DOCUMENT_ID ?? 'bilig-demo'
  const replicaId = options.replica ?? `codex:${Date.now()}`

  const open = await sendFrame(server, {
    kind: 'request',
    request: {
      kind: 'openWorkbookSession',
      id: `open:${Date.now()}`,
      documentId,
      replicaId,
    },
  })
  if (open.kind !== 'ok' || !open.sessionId) {
    throw new Error('Failed to open workbook session')
  }

  const sessionId = open.sessionId
  try {
    const response = await runCommand(command, { ...options, server }, sessionId)
    console.log(JSON.stringify(response, null, 2))
  } finally {
    try {
      await sendFrame(server, {
        kind: 'request',
        request: {
          kind: 'closeWorkbookSession',
          id: `close:${Date.now()}`,
          sessionId,
        },
      })
    } catch {}
  }
}

void (async () => {
  try {
    await main()
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error))
    process.exitCode = 1
  }
})()
