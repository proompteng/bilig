import { exportXlsx } from '@bilig/excel-import'
import { describe, expect, it } from 'vitest'
import { strToU8, zipSync } from 'fflate'
import { spawn } from 'node:child_process'
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { createWorkPaperFromDocument, exportWorkPaperDocument, parseWorkPaperDocument, serializeWorkPaperDocument } from '../persistence.js'
import { WorkPaper } from '../work-paper.js'
import { buildDemoWorkPaper } from '../work-paper-mcp-server.js'
import { createFileBackedWorkPaperMcpToolServerFromXlsxFile } from '../work-paper-mcp-xlsx-file.js'

describe('WorkPaper MCP XLSX file bridge', () => {
  it('imports an existing XLSX into writable file-backed MCP tools', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'bilig-workpaper-mcp-xlsx-'))
    const xlsxPath = join(tempDir, 'pricing.xlsx')
    const workpaperPath = join(tempDir, '.bilig', 'pricing.workpaper.json')
    const source = WorkPaper.buildFromSheets({
      Inputs: [
        ['Metric', 'Value'],
        ['Customers', 12],
        ['ARR per customer', 5000],
        ['Win rate', 0.25],
      ],
      Summary: [
        ['Metric', 'Value'],
        ['Expected customers', '=Inputs!B2*Inputs!B4'],
        ['Expected ARR', '=B2*Inputs!B3'],
      ],
    })

    try {
      writeFileSync(xlsxPath, exportXlsx(source.exportSnapshot()))
      const server = createFileBackedWorkPaperMcpToolServerFromXlsxFile({
        fromXlsxPath: xlsxPath,
        workpaperPath,
        writable: true,
      })

      const read = server.handleJsonRpc({
        jsonrpc: '2.0',
        id: 1,
        method: 'tools/call',
        params: {
          name: 'read_cell',
          arguments: {
            sheetName: 'Summary',
            address: 'B3',
          },
        },
      })
      expect(read.result).toMatchObject({
        isError: false,
        structuredContent: {
          address: 'Summary!B3',
          formula: '=B2*Inputs!B3',
          value: {
            value: 15000,
          },
        },
      })

      const write = server.handleJsonRpc({
        jsonrpc: '2.0',
        id: 2,
        method: 'tools/call',
        params: {
          name: 'set_cell_contents_and_readback',
          arguments: {
            sheetName: 'Inputs',
            address: 'B4',
            value: 0.4,
            readbackRange: 'Summary!A1:B3',
          },
        },
      })
      expect(write.result).toMatchObject({
        isError: false,
        structuredContent: {
          editedCell: 'Inputs!B4',
          afterReadback: {
            values: [
              [expect.objectContaining({ value: 'Metric' }), expect.objectContaining({ value: 'Value' })],
              [expect.objectContaining({ value: 'Expected customers' }), expect.objectContaining({ value: expect.any(Number) })],
              [expect.objectContaining({ value: 'Expected ARR' }), expect.objectContaining({ value: expect.any(Number) })],
            ],
          },
          persistence: {
            persisted: true,
            path: workpaperPath,
          },
          checks: {
            persisted: true,
            readbackChanged: true,
            restoredReadbackMatchesAfter: true,
          },
        },
      })
      expect(readCellValue(write.result, ['afterReadback', 'values', '1', '1'])).toBeCloseTo(4.8, 12)
      expect(readCellValue(write.result, ['afterReadback', 'values', '2', '1'])).toBeCloseTo(24000, 9)

      const restored = createWorkPaperFromDocument(parseWorkPaperDocument(readFileSync(workpaperPath, 'utf8')))
      const inputSheet = restored.getSheetId('Inputs')
      if (inputSheet === undefined) {
        throw new Error('Expected restored Inputs sheet')
      }
      expect(restored.getCellSerialized({ sheet: inputSheet, row: 3, col: 1 })).toBe(0.4)
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('refuses to overwrite a generated WorkPaper JSON from XLSX without an explicit flag', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'bilig-workpaper-mcp-xlsx-overwrite-'))
    const xlsxPath = join(tempDir, 'pricing.xlsx')
    const workpaperPath = join(tempDir, 'pricing.workpaper.json')

    try {
      writeFileSync(xlsxPath, exportXlsx(buildDemoWorkPaper().exportSnapshot()))
      writeFileSync(workpaperPath, serializeWorkPaperDocument(exportWorkPaperDocument(buildDemoWorkPaper(), { includeConfig: true })))

      expect(() =>
        createFileBackedWorkPaperMcpToolServerFromXlsxFile({
          fromXlsxPath: xlsxPath,
          workpaperPath,
        }),
      ).toThrow('pass --overwrite-workpaper to replace it')

      expect(() =>
        createFileBackedWorkPaperMcpToolServerFromXlsxFile({
          fromXlsxPath: xlsxPath,
          overwriteWorkPaper: true,
          workpaperPath,
        }),
      ).not.toThrow()
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('rejects large XLSX WorkPaper MCP imports before WorkPaper materialization', () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'bilig-workpaper-mcp-xlsx-large-'))
    const xlsxPath = join(tempDir, 'large.xlsx')
    const workpaperPath = join(tempDir, 'large.workpaper.json')

    try {
      writeFileSync(xlsxPath, largeValidXlsxBytes())

      expect(() =>
        createFileBackedWorkPaperMcpToolServerFromXlsxFile({
          fromXlsxPath: xlsxPath,
          workpaperPath,
        }),
      ).toThrow('small-workbook WorkPaper materialization path')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('starts the stdio bin directly from an XLSX and exposes risk analysis', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'bilig-workpaper-mcp-xlsx-direct-'))
    const xlsxPath = join(tempDir, 'pricing.xlsx')

    try {
      writeFileSync(xlsxPath, exportXlsx(buildDemoWorkPaper().exportSnapshot()))
      const binPath = fileURLToPath(new URL('../work-paper-mcp-stdio-bin.ts', import.meta.url))
      const child = spawn(process.execPath, ['--import', 'tsx', binPath, '--from-xlsx', xlsxPath], {
        stdio: ['pipe', 'pipe', 'pipe'],
      })
      const stdout: string[] = []
      const stderr: string[] = []
      const exitPromise = new Promise<number | null>((resolve, reject) => {
        const timeout = setTimeout(() => {
          child.kill('SIGKILL')
          reject(new Error('Timed out waiting for direct XLSX bilig-workpaper-mcp smoke test process to exit'))
        }, 8000)

        child.once('error', (error) => {
          clearTimeout(timeout)
          reject(error)
        })
        child.once('exit', (code) => {
          clearTimeout(timeout)
          resolve(code)
        })
      })

      child.stdout.setEncoding('utf8')
      child.stdout.on('data', (chunk: string) => stdout.push(chunk))
      child.stderr.setEncoding('utf8')
      child.stderr.on('data', (chunk: string) => stderr.push(chunk))
      child.stdin.end(
        [
          { jsonrpc: '2.0', id: 1, method: 'initialize' },
          { jsonrpc: '2.0', id: 2, method: 'tools/list' },
          {
            jsonrpc: '2.0',
            id: 3,
            method: 'tools/call',
            params: {
              name: 'analyze_workbook_risk',
              arguments: {
                inspectLimit: 'all',
              },
            },
          },
          {
            jsonrpc: '2.0',
            id: 4,
            method: 'tools/call',
            params: {
              name: 'read_cell',
              arguments: {
                sheetName: 'Summary',
                address: 'B3',
              },
            },
          },
        ]
          .map((request) => JSON.stringify(request))
          .join('\n') + '\n',
      )

      await expect(exitPromise).resolves.toBe(0)
      expect(stderr.join('')).toBe('')

      const responses = stdout
        .join('')
        .trim()
        .split('\n')
        .filter((line) => line.length > 0)
        .map((line) => JSON.parse(line))

      expect(responses[1].result.tools.map((tool: { readonly name: string }) => tool.name)).toContain('analyze_workbook_risk')
      expect(responses[2].result.structuredContent).toMatchObject({
        schemaVersion: 'bilig-workbook-compatibility-report.v1',
        verified: true,
        input: {
          fileName: 'pricing.xlsx',
          inspectLimit: 'all',
        },
        excelParity: 'not_proven',
      })
      expect(responses[3].result.structuredContent).toMatchObject({
        address: 'Summary!B3',
        formula: '=B2*Inputs!B4',
        value: {
          value: 60000,
        },
      })
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  }, 20_000)

  it('starts the stdio bin from an XLSX and creates the WorkPaper JSON', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'bilig-workpaper-mcp-xlsx-stdio-'))
    const xlsxPath = join(tempDir, 'pricing.xlsx')
    const workpaperPath = join(tempDir, '.bilig', 'pricing.workpaper.json')

    try {
      writeFileSync(xlsxPath, exportXlsx(buildDemoWorkPaper().exportSnapshot()))
      const binPath = fileURLToPath(new URL('../work-paper-mcp-stdio-bin.ts', import.meta.url))
      const child = spawn(
        process.execPath,
        ['--import', 'tsx', binPath, '--from-xlsx', xlsxPath, '--workpaper', workpaperPath, '--writable'],
        {
          stdio: ['pipe', 'pipe', 'pipe'],
        },
      )
      const stdout: string[] = []
      const stderr: string[] = []
      const exitPromise = new Promise<number | null>((resolve, reject) => {
        const timeout = setTimeout(() => {
          child.kill('SIGKILL')
          reject(new Error('Timed out waiting for XLSX-backed bilig-workpaper-mcp smoke test process to exit'))
        }, 8000)

        child.once('error', (error) => {
          clearTimeout(timeout)
          reject(error)
        })
        child.once('exit', (code) => {
          clearTimeout(timeout)
          resolve(code)
        })
      })

      child.stdout.setEncoding('utf8')
      child.stdout.on('data', (chunk: string) => stdout.push(chunk))
      child.stderr.setEncoding('utf8')
      child.stderr.on('data', (chunk: string) => stderr.push(chunk))
      child.stdin.end(
        [
          { jsonrpc: '2.0', id: 1, method: 'initialize' },
          { jsonrpc: '2.0', id: 2, method: 'tools/list' },
          {
            jsonrpc: '2.0',
            id: 3,
            method: 'tools/call',
            params: {
              name: 'analyze_workbook_risk',
              arguments: {
                inspectLimit: 'all',
              },
            },
          },
          {
            jsonrpc: '2.0',
            id: 4,
            method: 'tools/call',
            params: {
              name: 'read_cell',
              arguments: {
                sheetName: 'Summary',
                address: 'B3',
              },
            },
          },
        ]
          .map((request) => JSON.stringify(request))
          .join('\n') + '\n',
      )

      await expect(exitPromise).resolves.toBe(0)
      expect(stderr.join('')).toBe('')

      const responses = stdout
        .join('')
        .trim()
        .split('\n')
        .filter((line) => line.length > 0)
        .map((line) => JSON.parse(line))

      expect(responses[1].result.tools.map((tool: { readonly name: string }) => tool.name)).toContain('analyze_workbook_risk')
      expect(responses[2].result.structuredContent).toMatchObject({
        schemaVersion: 'bilig-workbook-compatibility-report.v1',
        verified: true,
        input: {
          fileName: 'pricing.xlsx',
          inspectLimit: 'all',
        },
        excelParity: 'not_proven',
      })
      expect(responses[3].result.structuredContent).toMatchObject({
        address: 'Summary!B3',
        formula: '=B2*Inputs!B4',
        value: {
          value: 60000,
        },
      })

      const restored = createWorkPaperFromDocument(parseWorkPaperDocument(readFileSync(workpaperPath, 'utf8')))
      expect(restored.getSheetId('Inputs')).not.toBeUndefined()
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  }, 20_000)
})

function largeValidXlsxBytes(): Uint8Array {
  return zipSync(
    {
      '[Content_Types].xml': strToU8(
        [
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
          '<Default Extension="bin" ContentType="application/octet-stream"/>',
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
          '<Default Extension="xml" ContentType="application/xml"/>',
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
          '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
          '</Types>',
        ].join(''),
      ),
      '_rels/.rels': strToU8(
        [
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
          '</Relationships>',
        ].join(''),
      ),
      'xl/workbook.xml': strToU8(
        [
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
          '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>',
          '</workbook>',
        ].join(''),
      ),
      'xl/_rels/workbook.xml.rels': strToU8(
        [
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>',
          '</Relationships>',
        ].join(''),
      ),
      'xl/worksheets/sheet1.xml': strToU8(
        [
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
          '<dimension ref="A1"/>',
          '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>',
          '</worksheet>',
        ].join(''),
      ),
      'docProps/padding.bin': new Uint8Array(1_100_000),
    },
    { level: 0 },
  )
}

function readCellValue(value: unknown, path: readonly string[]): number {
  let current: unknown = readStructuredContent(value)
  for (const segment of path) {
    if (Array.isArray(current)) {
      current = current[Number(segment)]
      continue
    }
    if (!isRecord(current)) {
      throw new Error(`Expected structured content object while reading ${path.join('.')}, received ${JSON.stringify(current)}`)
    }
    current = current[segment]
  }
  if (!isRecord(current) || typeof current['value'] !== 'number') {
    throw new Error(`Expected numeric cell value at ${path.join('.')}, received ${JSON.stringify(current)}`)
  }
  return current['value']
}

function readStructuredContent(value: unknown): Record<string, unknown> {
  if (!isRecord(value) || !isRecord(value['structuredContent'])) {
    throw new Error(`Expected tool result with structuredContent, received ${JSON.stringify(value)}`)
  }
  return value['structuredContent']
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
