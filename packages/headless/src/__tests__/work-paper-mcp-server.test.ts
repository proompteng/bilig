import { describe, expect, it } from 'vitest'
import { spawn } from 'node:child_process'
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { createWorkPaperFromDocument, exportWorkPaperDocument, parseWorkPaperDocument, serializeWorkPaperDocument } from '../persistence.js'
import { createFileBackedWorkPaperMcpToolServer } from '../work-paper-mcp-file-server.js'
import { parseWorkPaperMcpStdioCliArgs, workPaperMcpStdioHelpText } from '../work-paper-mcp-stdio-cli.js'
import {
  assertWorkPaperMcpDemoOutput,
  buildDemoWorkPaper,
  createWorkPaperMcpDemoOutput,
  createWorkPaperMcpToolServer,
} from '../work-paper-mcp-server.js'

describe('WorkPaper MCP server', () => {
  it('parses stdio bin CLI options without starting the server', () => {
    expect(parseWorkPaperMcpStdioCliArgs(['--workpaper', 'pricing.workpaper.json', '--writable'])).toEqual({
      demoWorkPaperTools: false,
      help: false,
      writable: true,
      workpaperPath: 'pricing.workpaper.json',
    })
    expect(parseWorkPaperMcpStdioCliArgs(['--help'])).toEqual({
      demoWorkPaperTools: false,
      help: true,
      writable: false,
    })
    expect(parseWorkPaperMcpStdioCliArgs(['--demo-workpaper-tools'])).toEqual({
      demoWorkPaperTools: true,
      help: false,
      writable: false,
    })
    expect(workPaperMcpStdioHelpText()).toContain('Usage: bilig-workpaper-mcp')
  })

  it('rejects malformed stdio bin workpaper paths before opening files', () => {
    expect(() => parseWorkPaperMcpStdioCliArgs(['--workpaper', '   '])).toThrow('--workpaper requires a path')
    expect(() => parseWorkPaperMcpStdioCliArgs(['--workpaper', '--writable'])).toThrow('--workpaper requires a path')
    expect(() => parseWorkPaperMcpStdioCliArgs(['--demo-workpaper-tools', '--workpaper', 'pricing.workpaper.json'])).toThrow(
      '--demo-workpaper-tools cannot be combined with --workpaper',
    )
  })

  it('starts the stdio bin and exposes the expected tools', async () => {
    const binPath = fileURLToPath(new URL('../work-paper-mcp-stdio-bin.ts', import.meta.url))
    const child = spawn(process.execPath, ['--import', 'tsx', binPath], {
      stdio: ['pipe', 'pipe', 'pipe'],
    })
    const stdout: string[] = []
    const stderr: string[] = []
    const exitPromise = new Promise<number | null>((resolve, reject) => {
      const timeout = setTimeout(() => {
        child.kill('SIGKILL')
        reject(new Error('Timed out waiting for bilig-workpaper-mcp smoke test process to exit'))
      }, 5000)

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
      `${JSON.stringify({ jsonrpc: '2.0', id: 1, method: 'initialize' })}\n${JSON.stringify({ jsonrpc: '2.0', id: 2, method: 'tools/list' })}\n`,
    )

    await expect(exitPromise).resolves.toBe(0)

    const responses = stdout
      .join('')
      .trim()
      .split('\n')
      .filter((line) => line.length > 0)
      .map((line) => JSON.parse(line))

    expect(stderr.join('')).toBe('')
    expect(responses).toHaveLength(2)
    expect(responses[0]).toMatchObject({
      jsonrpc: '2.0',
      id: 1,
      result: {
        capabilities: {
          tools: {
            listChanged: false,
          },
        },
      },
    })
    expect(responses[1].result.tools.map((tool: { name: string }) => tool.name)).toEqual([
      'read_workpaper_summary',
      'set_workpaper_input_cell',
    ])
  })

  it('exposes stable tool definitions and structured formula readback', () => {
    const output = createWorkPaperMcpDemoOutput()

    assertWorkPaperMcpDemoOutput(output)
    expect(output.listResponse.result.tools.map((tool) => tool.name)).toEqual(['read_workpaper_summary', 'set_workpaper_input_cell'])
    expect(output.listResponse.result.tools).toEqual([
      expect.objectContaining({
        title: 'Read WorkPaper Summary',
        annotations: {
          title: 'Read WorkPaper Summary',
          readOnlyHint: true,
          destructiveHint: false,
          idempotentHint: true,
          openWorldHint: false,
        },
      }),
      expect.objectContaining({
        title: 'Set WorkPaper Input Cell',
        annotations: {
          title: 'Set WorkPaper Input Cell',
          readOnlyHint: false,
          destructiveHint: true,
          idempotentHint: true,
          openWorldHint: false,
        },
      }),
    ])
    expect(output.writeResponse.result.structuredContent).toMatchObject({
      editedCell: 'Inputs!B3',
      after: {
        expectedArr: 96000,
        expansionArr: 105600,
        targetGap: 5600,
      },
      checks: {
        formulasPersisted: true,
        restoredMatchesAfter: true,
      },
    })
  })

  it('rejects unknown tool calls instead of returning a misleading success', () => {
    const server = createWorkPaperMcpToolServer(buildDemoWorkPaper())

    expect(() =>
      server.handleJsonRpc({
        jsonrpc: '2.0',
        id: 1,
        method: 'tools/call',
        params: {
          name: 'unknown_tool',
          arguments: {},
        },
      }),
    ).toThrow('Unknown WorkPaper tool')
  })

  it('exposes file-backed tools for real WorkPaper JSON documents', () => {
    const persistCalls: string[] = []
    const workbook = buildDemoWorkPaper()
    const server = createFileBackedWorkPaperMcpToolServer({
      workbook,
      writable: true,
      sourcePath: '/tmp/pricing.workpaper.json',
      persist(updatedWorkbook) {
        const serialized = serializeWorkPaperDocument(exportWorkPaperDocument(updatedWorkbook, { includeConfig: true }))
        persistCalls.push(serialized)
        return {
          persisted: true,
          path: '/tmp/pricing.workpaper.json',
          serializedBytes: Buffer.byteLength(serialized, 'utf8'),
        }
      },
    })

    const tools = server.handleJsonRpc({
      jsonrpc: '2.0',
      id: 1,
      method: 'tools/list',
    })
    expect(readToolNames(tools.result)).toEqual([
      'list_sheets',
      'read_range',
      'read_cell',
      'set_cell_contents',
      'get_cell_display_value',
      'export_workpaper_document',
      'validate_formula',
    ])

    const read = server.handleJsonRpc({
      jsonrpc: '2.0',
      id: 2,
      method: 'tools/call',
      params: {
        name: 'read_range',
        arguments: {
          range: 'Summary!A1:B5',
        },
      },
    })
    expect(read.result).toMatchObject({
      isError: false,
      structuredContent: {
        range: 'Summary!A1:B5',
      },
    })

    const write = server.handleJsonRpc({
      jsonrpc: '2.0',
      id: 3,
      method: 'tools/call',
      params: {
        name: 'set_cell_contents',
        arguments: {
          sheetName: 'Inputs',
          address: 'B3',
          value: 0.4,
        },
      },
    })
    expect(write.result).toMatchObject({
      isError: false,
      structuredContent: {
        editedCell: 'Inputs!B3',
        before: {
          serialized: 0.25,
        },
        after: {
          serialized: 0.4,
        },
        persistence: {
          persisted: true,
          path: '/tmp/pricing.workpaper.json',
        },
        checks: {
          persisted: true,
          restoredMatchesAfter: true,
        },
      },
    })
    expect(persistCalls).toHaveLength(1)

    const validate = server.handleJsonRpc({
      jsonrpc: '2.0',
      id: 4,
      method: 'tools/call',
      params: {
        name: 'validate_formula',
        arguments: {
          formula: '=SUM(1,2)',
        },
      },
    })
    expect(validate.result).toMatchObject({
      structuredContent: {
        formula: '=SUM(1,2)',
        valid: true,
      },
    })
  })

  it('starts the stdio bin in writable file-backed mode and persists edits', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'bilig-workpaper-mcp-'))
    const workpaperPath = join(tempDir, 'pricing.workpaper.json')
    writeFileSync(workpaperPath, serializeWorkPaperDocument(exportWorkPaperDocument(buildDemoWorkPaper(), { includeConfig: true })))

    try {
      const binPath = fileURLToPath(new URL('../work-paper-mcp-stdio-bin.ts', import.meta.url))
      const child = spawn(process.execPath, ['--import', 'tsx', binPath, '--workpaper', workpaperPath, '--writable'], {
        stdio: ['pipe', 'pipe', 'pipe'],
      })
      const stdout: string[] = []
      const stderr: string[] = []
      const exitPromise = new Promise<number | null>((resolve, reject) => {
        const timeout = setTimeout(() => {
          child.kill('SIGKILL')
          reject(new Error('Timed out waiting for file-backed bilig-workpaper-mcp smoke test process to exit'))
        }, 5000)

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
              name: 'set_cell_contents',
              arguments: {
                sheetName: 'Inputs',
                address: 'B3',
                value: 0.4,
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

      expect(responses[1].result.tools.map((tool: { name: string }) => tool.name)).toContain('set_cell_contents')
      expect(responses[2].result.structuredContent).toMatchObject({
        editedCell: 'Inputs!B3',
        checks: {
          persisted: true,
          restoredMatchesAfter: true,
        },
      })

      const restored = createWorkPaperFromDocument(parseWorkPaperDocument(readFileSync(workpaperPath, 'utf8')))
      const inputSheet = restored.getSheetId('Inputs')
      if (inputSheet === undefined) {
        throw new Error('Expected restored Inputs sheet')
      }
      expect(restored.getCellSerialized({ sheet: inputSheet, row: 2, col: 1 })).toBe(0.4)
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('starts the stdio bin with directory-friendly WorkPaper tools for scanner introspection', async () => {
    const binPath = fileURLToPath(new URL('../work-paper-mcp-stdio-bin.ts', import.meta.url))
    const child = spawn(process.execPath, ['--import', 'tsx', binPath, '--demo-workpaper-tools'], {
      stdio: ['pipe', 'pipe', 'pipe'],
    })
    const stdout: string[] = []
    const stderr: string[] = []
    const exitPromise = new Promise<number | null>((resolve, reject) => {
      const timeout = setTimeout(() => {
        child.kill('SIGKILL')
        reject(new Error('Timed out waiting for demo WorkPaper tool smoke test process to exit'))
      }, 10000)

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

    expect(responses[1].result.tools.map((tool: { name: string }) => tool.name)).toEqual([
      'list_sheets',
      'read_range',
      'read_cell',
      'set_cell_contents',
      'get_cell_display_value',
      'export_workpaper_document',
      'validate_formula',
    ])
  }, 15000)
})

function readToolNames(value: unknown): string[] {
  if (!isRecord(value) || !Array.isArray(value['tools'])) {
    throw new Error(`Expected tools/list result, received ${JSON.stringify(value)}`)
  }

  return value['tools'].map((tool) => {
    if (!isRecord(tool) || typeof tool['name'] !== 'string') {
      throw new Error(`Expected MCP tool definition, received ${JSON.stringify(tool)}`)
    }
    return tool['name']
  })
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
