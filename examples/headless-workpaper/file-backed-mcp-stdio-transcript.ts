import { spawn } from 'node:child_process'
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { basename, join } from 'node:path'

import { createWorkPaperFromDocument, exportWorkPaperDocument, parseWorkPaperDocument, serializeWorkPaperDocument } from '@bilig/workpaper'
import { buildDemoWorkPaper } from '@bilig/workpaper/mcp'

type JsonObject = Record<string, unknown>

interface FileBackedTranscriptOutput {
  transport: 'stdio'
  command: string
  requestLines: string[]
  responseSummary: {
    protocolVersion: string
    serverName: string
    toolNames: string[]
    beforeRange: string
    write: {
      editedCell: string
      beforeSerialized: unknown
      afterSerialized: unknown
      persistence: {
        persisted: boolean
        path: string
        serializedBytes: number
      }
      checks: JsonObject
    }
    recalculatedCell: {
      address: string
      formula: string
      value: number
      displayValue: string
    }
  }
  verified: {
    initialized: boolean
    listedFileBackedTools: boolean
    editedInputCell: 'Inputs!B3'
    persistedToDisk: boolean
    formulaRecalculated: boolean
    restoredInputSerialized: 0.4
    restoredMatchesAfter: boolean
  }
}

const workspace = mkdtempSync(join(tmpdir(), 'bilig-workpaper-mcp-file-'))
const workpaperPath = join(workspace, 'pricing.workpaper.json')

try {
  writeFileSync(workpaperPath, serializeWorkPaperDocument(exportWorkPaperDocument(buildDemoWorkPaper(), { includeConfig: true })))

  const requestLines = [
    jsonLine({ jsonrpc: '2.0', id: 1, method: 'initialize' }),
    jsonLine({ jsonrpc: '2.0', method: 'notifications/initialized' }),
    jsonLine({ jsonrpc: '2.0', id: 2, method: 'tools/list' }),
    jsonLine({
      jsonrpc: '2.0',
      id: 3,
      method: 'tools/call',
      params: {
        name: 'read_range',
        arguments: {
          range: 'Summary!A1:B5',
        },
      },
    }),
    jsonLine({
      jsonrpc: '2.0',
      id: 4,
      method: 'tools/call',
      params: {
        name: 'set_cell_contents',
        arguments: {
          sheetName: 'Inputs',
          address: 'B3',
          value: 0.4,
        },
      },
    }),
    jsonLine({
      jsonrpc: '2.0',
      id: 5,
      method: 'tools/call',
      params: {
        name: 'read_cell',
        arguments: {
          sheetName: 'Summary',
          address: 'B3',
        },
      },
    }),
  ]

  const responseLines = await runPublishedMcpBinary(workpaperPath, requestLines)
  const transcriptOutput = createTranscriptOutput(workpaperPath, requestLines, responseLines)
  assertTranscriptOutput(transcriptOutput)

  console.log(JSON.stringify(transcriptOutput, null, 2))
} finally {
  rmSync(workspace, { recursive: true, force: true })
}

function jsonLine(value: unknown): string {
  return JSON.stringify(value)
}

async function runPublishedMcpBinary(path: string, lines: string[]): Promise<string[]> {
  const npmCommand = process.platform === 'win32' ? 'npm.cmd' : 'npm'
  const child = spawn(
    npmCommand,
    ['exec', '--loglevel=error', '--package', '@bilig/workpaper@latest', '--', 'bilig-workpaper-mcp', '--workpaper', path, '--writable'],
    {
      cwd: process.cwd(),
      env: {
        ...process.env,
        NODE_NO_WARNINGS: '1',
      },
      stdio: ['pipe', 'pipe', 'pipe'],
    },
  )

  let stdout = ''
  let stderr = ''
  child.stdout.setEncoding('utf8')
  child.stderr.setEncoding('utf8')
  child.stdout.on('data', (chunk: string) => {
    stdout += chunk
  })
  child.stderr.on('data', (chunk: string) => {
    stderr += chunk
  })

  child.stdin.end(`${lines.join('\n')}\n`)

  const exitCode = await new Promise<number | null>((resolve, reject) => {
    child.on('error', reject)
    child.on('close', resolve)
  })

  if (exitCode !== 0) {
    throw new Error(`File-backed MCP stdio transcript failed with exit ${String(exitCode)}: ${stderr}`)
  }
  if (stderr.trim().length > 0) {
    throw new Error(`File-backed MCP stdio transcript wrote stderr: ${stderr}`)
  }

  return stdout
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0)
}

function createTranscriptOutput(path: string, requests: string[], responses: string[]): FileBackedTranscriptOutput {
  const initializeResponse = requireResponse(responses, 1)
  const toolsResponse = requireResponse(responses, 2)
  const beforeRangeResponse = requireResponse(responses, 3)
  const writeResponse = requireResponse(responses, 4)
  const recalculatedCellResponse = requireResponse(responses, 5)

  const initializeResult = readRecord(initializeResponse.result, 'initialize result')
  const serverInfo = readRecord(initializeResult.serverInfo, 'server info')
  const toolsResult = readRecord(toolsResponse.result, 'tools/list result')
  const beforeRange = readRecord(
    readRecord(beforeRangeResponse.result, 'before range result').structuredContent,
    'before range structured content',
  )
  const write = readRecord(readRecord(writeResponse.result, 'write result').structuredContent, 'write structured content')
  const recalculatedCell = readRecord(
    readRecord(recalculatedCellResponse.result, 'recalculated cell result').structuredContent,
    'recalculated cell structured content',
  )
  const persistence = readRecord(write.persistence, 'write persistence')
  const checks = readRecord(write.checks, 'write checks')
  const restored = createWorkPaperFromDocument(parseWorkPaperDocument(readFileSync(path, 'utf8')))
  const inputsSheet = requireSheet(restored, 'Inputs')

  return {
    transport: 'stdio',
    command: 'npm exec --package @bilig/workpaper@latest -- bilig-workpaper-mcp --workpaper pricing.workpaper.json --writable',
    requestLines: requests,
    responseSummary: {
      protocolVersion: readString(initializeResult.protocolVersion, 'protocolVersion'),
      serverName: readString(serverInfo.name, 'server name'),
      toolNames: readToolNames(toolsResult.tools),
      beforeRange: readString(beforeRange.range, 'before range'),
      write: {
        editedCell: readString(write.editedCell, 'edited cell'),
        beforeSerialized: readRecord(write.before, 'write before').serialized,
        afterSerialized: readRecord(write.after, 'write after').serialized,
        persistence: {
          persisted: readBoolean(persistence.persisted, 'persistence persisted'),
          path: basename(readString(persistence.path, 'persistence path')),
          serializedBytes: readNumber(persistence.serializedBytes, 'persistence serialized bytes'),
        },
        checks,
      },
      recalculatedCell: {
        address: readString(recalculatedCell.address, 'recalculated cell address'),
        formula: readString(recalculatedCell.formula, 'recalculated cell formula'),
        value: readCellNumber(recalculatedCell.value, 'recalculated cell value'),
        displayValue: readString(recalculatedCell.displayValue, 'recalculated cell display value'),
      },
    },
    verified: {
      initialized: true,
      listedFileBackedTools: true,
      editedInputCell: readEditedCell(write.editedCell),
      persistedToDisk: readBoolean(persistence.persisted, 'persistence persisted'),
      formulaRecalculated: readCellNumber(recalculatedCell.value, 'recalculated cell value') === 96000,
      restoredInputSerialized: readRestoredInput(restored, inputsSheet),
      restoredMatchesAfter: readBoolean(checks.restoredMatchesAfter, 'restoredMatchesAfter'),
    },
  }
}

function requireResponse(lines: string[], id: number): JsonObject {
  for (const line of lines) {
    const parsed = readRecord(JSON.parse(line), `response ${id.toString()}`)
    if (parsed.id === id && parsed.jsonrpc === '2.0') {
      return parsed
    }
  }
  throw new Error(`Missing JSON-RPC response id ${id.toString()}: ${lines.join('\n')}`)
}

function readToolNames(value: unknown): string[] {
  if (!Array.isArray(value)) {
    throw new Error(`Expected tools/list result to be an array, got ${JSON.stringify(value)}`)
  }

  return value.map((entry, index) => readString(readRecord(entry, `tool ${index.toString()}`).name, `tool ${index.toString()} name`))
}

function assertTranscriptOutput(output: FileBackedTranscriptOutput): void {
  const expectedTools = [
    'list_sheets',
    'read_range',
    'read_cell',
    'set_cell_contents',
    'get_cell_display_value',
    'export_workpaper_document',
    'validate_formula',
  ]

  if (
    output.responseSummary.protocolVersion !== '2025-06-18' ||
    output.responseSummary.serverName !== 'bilig-headless-workpaper' ||
    JSON.stringify(output.responseSummary.toolNames) !== JSON.stringify(expectedTools) ||
    output.responseSummary.beforeRange !== 'Summary!A1:B5' ||
    output.responseSummary.write.beforeSerialized !== 0.25 ||
    output.responseSummary.write.afterSerialized !== 0.4 ||
    !output.responseSummary.write.persistence.persisted ||
    output.responseSummary.write.persistence.path !== 'pricing.workpaper.json' ||
    output.responseSummary.write.persistence.serializedBytes <= 0 ||
    output.responseSummary.recalculatedCell.address !== 'Summary!B3' ||
    output.responseSummary.recalculatedCell.formula !== '=B2*Inputs!B4' ||
    output.responseSummary.recalculatedCell.value !== 96000 ||
    !output.verified.initialized ||
    !output.verified.listedFileBackedTools ||
    output.verified.editedInputCell !== 'Inputs!B3' ||
    !output.verified.persistedToDisk ||
    !output.verified.formulaRecalculated ||
    output.verified.restoredInputSerialized !== 0.4 ||
    !output.verified.restoredMatchesAfter
  ) {
    throw new Error(`Unexpected file-backed MCP stdio transcript: ${JSON.stringify(output)}`)
  }
}

function readEditedCell(value: unknown): 'Inputs!B3' {
  if (value !== 'Inputs!B3') {
    throw new Error(`Expected edited cell to be Inputs!B3, got ${JSON.stringify(value)}`)
  }
  return value
}

function readRestoredInput(workbook: ReturnType<typeof createWorkPaperFromDocument>, sheet: number): 0.4 {
  const value = workbook.getCellSerialized({ sheet, row: 2, col: 1 })
  if (value !== 0.4) {
    throw new Error(`Expected restored Inputs!B3 to be 0.4, got ${JSON.stringify(value)}`)
  }
  return value
}

function requireSheet(workbook: ReturnType<typeof createWorkPaperFromDocument>, sheetName: string): number {
  const sheetId = workbook.getSheetId(sheetName)
  if (sheetId === undefined) {
    throw new Error(`Expected sheet "${sheetName}" to exist`)
  }
  return sheetId
}

function readCellNumber(value: unknown, label: string): number {
  const record = readRecord(value, label)
  const cellValue = record.value
  if (typeof cellValue !== 'number') {
    throw new Error(`Expected ${label} to contain a numeric value, got ${JSON.stringify(value)}`)
  }
  return cellValue
}

function readRecord(value: unknown, label: string): JsonObject {
  if (!isJsonObject(value)) {
    throw new Error(`Expected ${label} to be an object, got ${JSON.stringify(value)}`)
  }
  return value
}

function isJsonObject(value: unknown): value is JsonObject {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}

function readString(value: unknown, label: string): string {
  if (typeof value !== 'string') {
    throw new Error(`Expected ${label} to be a string, got ${JSON.stringify(value)}`)
  }
  return value
}

function readNumber(value: unknown, label: string): number {
  if (typeof value !== 'number') {
    throw new Error(`Expected ${label} to be a number, got ${JSON.stringify(value)}`)
  }
  return value
}

function readBoolean(value: unknown, label: string): boolean {
  if (typeof value !== 'boolean') {
    throw new Error(`Expected ${label} to be a boolean, got ${JSON.stringify(value)}`)
  }
  return value
}
