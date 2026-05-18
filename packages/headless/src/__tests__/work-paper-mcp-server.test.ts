import { describe, expect, it } from 'vitest'
import { spawn } from 'node:child_process'
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { createWorkPaperFromDocument, exportWorkPaperDocument, parseWorkPaperDocument, serializeWorkPaperDocument } from '../persistence.js'
import { WORKPAPER_VERSION } from '../work-paper-version.js'
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
      initDemoWorkPaper: false,
      writable: true,
      workpaperPath: 'pricing.workpaper.json',
    })
    expect(parseWorkPaperMcpStdioCliArgs(['--help'])).toEqual({
      demoWorkPaperTools: false,
      help: true,
      initDemoWorkPaper: false,
      writable: false,
    })
    expect(parseWorkPaperMcpStdioCliArgs(['--demo-workpaper-tools'])).toEqual({
      demoWorkPaperTools: true,
      help: false,
      initDemoWorkPaper: false,
      writable: false,
    })
    expect(parseWorkPaperMcpStdioCliArgs(['--workpaper', 'pricing.workpaper.json', '--init-demo-workpaper', '--writable'])).toEqual({
      demoWorkPaperTools: false,
      help: false,
      initDemoWorkPaper: true,
      writable: true,
      workpaperPath: 'pricing.workpaper.json',
    })
    expect(workPaperMcpStdioHelpText()).toContain('Usage: bilig-workpaper-mcp')
  })

  it('rejects malformed stdio bin workpaper paths before opening files', () => {
    expect(() => parseWorkPaperMcpStdioCliArgs(['--workpaper', '   '])).toThrow('--workpaper requires a path')
    expect(() => parseWorkPaperMcpStdioCliArgs(['--workpaper', '--writable'])).toThrow('--workpaper requires a path')
    expect(() => parseWorkPaperMcpStdioCliArgs(['--demo-workpaper-tools', '--workpaper', 'pricing.workpaper.json'])).toThrow(
      '--demo-workpaper-tools cannot be combined with --workpaper',
    )
    expect(() => parseWorkPaperMcpStdioCliArgs(['--init-demo-workpaper'])).toThrow('--init-demo-workpaper requires --workpaper')
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
        serverInfo: {
          name: 'bilig-headless-workpaper',
          version: WORKPAPER_VERSION,
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
        outputSchema: expect.objectContaining({
          required: ['range', 'values', 'serialized'],
        }),
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
        outputSchema: expect.objectContaining({
          required: ['editedCell', 'before', 'after', 'restored', 'formulaContracts', 'checks'],
        }),
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
    expect(readToolOutputSchemaRequired(tools.result, 'list_sheets')).toEqual(['writable', 'sheets'])
    expect(readToolOutputSchemaRequired(tools.result, 'read_cell')).toEqual(['address', 'value', 'serialized', 'formula', 'displayValue'])
    expect(readToolOutputSchemaRequired(tools.result, 'set_cell_contents')).toEqual([
      'editedCell',
      'before',
      'after',
      'restored',
      'persistence',
      'checks',
    ])
    expect(readToolOutputSchemaRequired(tools.result, 'validate_formula')).toEqual(['formula', 'valid'])
    expect(server.capabilities).toMatchObject({
      resources: {
        listChanged: false,
        subscribe: false,
      },
      prompts: {
        listChanged: false,
      },
    })

    const resources = server.handleJsonRpc({
      jsonrpc: '2.0',
      id: 5,
      method: 'resources/list',
    })
    expect(readResourceUris(resources.result)).toEqual([
      'bilig://workpaper/manifest',
      'bilig://workpaper/agent-handoff',
      'bilig://workpaper/sheets',
      'bilig://workpaper/current-document',
    ])

    const manifest = server.handleJsonRpc({
      jsonrpc: '2.0',
      id: 6,
      method: 'resources/read',
      params: {
        uri: 'bilig://workpaper/manifest',
      },
    })
    expect(JSON.parse(readFirstResourceText(manifest.result))).toMatchObject({
      server: 'bilig-workpaper-mcp',
      writable: true,
      sourcePath: '/tmp/pricing.workpaper.json',
      capabilities: {
        resources: {
          listChanged: false,
        },
        prompts: {
          listChanged: false,
        },
      },
      sheets: expect.arrayContaining([
        expect.objectContaining({
          name: 'Inputs',
          dimensions: {
            width: 2,
            height: 5,
          },
        }),
      ]),
      tools: expect.arrayContaining([
        expect.objectContaining({
          name: 'list_sheets',
        }),
      ]),
      prompts: expect.arrayContaining([
        expect.objectContaining({
          name: 'edit_and_verify_workpaper',
        }),
      ]),
      verificationContract: expect.arrayContaining(['read the dependent computed output after recalculation']),
    })

    const handoff = server.handleJsonRpc({
      jsonrpc: '2.0',
      id: 7,
      method: 'resources/read',
      params: {
        uri: 'bilig://workpaper/agent-handoff',
      },
    })
    expect(readFirstResourceText(handoff.result)).toContain('Do not report success unless computed readback')

    const prompts = server.handleJsonRpc({
      jsonrpc: '2.0',
      id: 8,
      method: 'prompts/list',
    })
    expect(readPromptNames(prompts.result)).toEqual(['edit_and_verify_workpaper', 'debug_workpaper_formula'])

    const prompt = server.handleJsonRpc({
      jsonrpc: '2.0',
      id: 9,
      method: 'prompts/get',
      params: {
        name: 'edit_and_verify_workpaper',
        arguments: {
          task: 'Change win rate and prove expected ARR changed.',
          target_cell: 'Inputs!B3',
          output_range: 'Summary!A1:B5',
        },
      },
    })
    expect(readPromptText(prompt.result)).toContain('Do not claim success from set_cell_contents alone')

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
          { jsonrpc: '2.0', id: 3, method: 'resources/list' },
          { jsonrpc: '2.0', id: 4, method: 'prompts/list' },
          {
            jsonrpc: '2.0',
            id: 5,
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
      expect(responses[2].result.resources.map((resource: { uri: string }) => resource.uri)).toContain('bilig://workpaper/agent-handoff')
      expect(responses[3].result.prompts.map((prompt: { name: string }) => prompt.name)).toContain('edit_and_verify_workpaper')
      expect(responses[4].result.structuredContent).toMatchObject({
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

  it('initializes a missing demo WorkPaper file before starting file-backed stdio tools', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'bilig-workpaper-mcp-init-'))
    const workpaperPath = join(tempDir, 'pricing.workpaper.json')

    try {
      const binPath = fileURLToPath(new URL('../work-paper-mcp-stdio-bin.ts', import.meta.url))
      const child = spawn(
        process.execPath,
        ['--import', 'tsx', binPath, '--workpaper', workpaperPath, '--init-demo-workpaper', '--writable'],
        {
          stdio: ['pipe', 'pipe', 'pipe'],
        },
      )
      const stdout: string[] = []
      const stderr: string[] = []
      const exitPromise = new Promise<number | null>((resolve, reject) => {
        const timeout = setTimeout(() => {
          child.kill('SIGKILL')
          reject(new Error('Timed out waiting for demo-init bilig-workpaper-mcp smoke test process to exit'))
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

      expect(responses[1].result.tools.map((tool: { name: string }) => tool.name)).toContain('set_cell_contents')
      expect(responses[2].result.structuredContent).toMatchObject({
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

function readToolOutputSchemaRequired(value: unknown, toolName: string): string[] {
  if (!isRecord(value) || !Array.isArray(value['tools'])) {
    throw new Error(`Expected tools/list result, received ${JSON.stringify(value)}`)
  }

  const tool = value['tools'].find((candidate) => isRecord(candidate) && candidate['name'] === toolName)
  if (!isRecord(tool)) {
    throw new Error(`Expected ${toolName} tool definition, received ${JSON.stringify(value)}`)
  }

  const outputSchema = tool['outputSchema']
  if (!isRecord(outputSchema) || !Array.isArray(outputSchema['required'])) {
    throw new Error(`Expected ${toolName} output schema, received ${JSON.stringify(tool)}`)
  }
  return outputSchema['required'].map((item) => {
    if (typeof item !== 'string') {
      throw new Error(`Expected ${toolName} output required item to be a string, received ${JSON.stringify(item)}`)
    }
    return item
  })
}

function readResourceUris(value: unknown): string[] {
  if (!isRecord(value) || !Array.isArray(value['resources'])) {
    throw new Error(`Expected resources/list result, received ${JSON.stringify(value)}`)
  }

  return value['resources'].map((resource) => {
    if (!isRecord(resource) || typeof resource['uri'] !== 'string') {
      throw new Error(`Expected MCP resource definition, received ${JSON.stringify(resource)}`)
    }
    return resource['uri']
  })
}

function readPromptNames(value: unknown): string[] {
  if (!isRecord(value) || !Array.isArray(value['prompts'])) {
    throw new Error(`Expected prompts/list result, received ${JSON.stringify(value)}`)
  }

  return value['prompts'].map((prompt) => {
    if (!isRecord(prompt) || typeof prompt['name'] !== 'string') {
      throw new Error(`Expected MCP prompt definition, received ${JSON.stringify(prompt)}`)
    }
    return prompt['name']
  })
}

function readFirstResourceText(value: unknown): string {
  if (!isRecord(value) || !Array.isArray(value['contents'])) {
    throw new Error(`Expected resources/read result, received ${JSON.stringify(value)}`)
  }

  const first = value['contents'][0]
  if (!isRecord(first) || typeof first['text'] !== 'string') {
    throw new Error(`Expected text resource content, received ${JSON.stringify(value)}`)
  }
  return first['text']
}

function readPromptText(value: unknown): string {
  if (!isRecord(value) || !Array.isArray(value['messages'])) {
    throw new Error(`Expected prompts/get result, received ${JSON.stringify(value)}`)
  }

  const first = value['messages'][0]
  const content = isRecord(first) ? first['content'] : undefined
  if (!isRecord(content) || typeof content['text'] !== 'string') {
    throw new Error(`Expected text prompt content, received ${JSON.stringify(value)}`)
  }
  return content['text']
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
