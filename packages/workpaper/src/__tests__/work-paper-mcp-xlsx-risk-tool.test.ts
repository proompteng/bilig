import { exportXlsx } from '@bilig/workpaper/xlsx'
import { WorkPaper } from '@bilig/workpaper'
import { describe, expect, it } from 'vitest'
import { spawn } from 'node:child_process'
import { mkdtempSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

interface McpTestResponse {
  readonly result?: {
    readonly structuredContent?: unknown
    readonly tools?: readonly {
      readonly name: string
    }[]
  }
}

describe('@bilig/workpaper MCP XLSX risk tool', () => {
  it('keeps the scoped package wrapper aligned with the XLSX-backed MCP path', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'bilig-scoped-workpaper-mcp-risk-'))
    const xlsxPath = join(tempDir, 'pricing.xlsx')
    const workpaperPath = join(tempDir, '.bilig', 'pricing.workpaper.json')
    const source = WorkPaper.buildFromSheets({
      Inputs: [
        ['Metric', 'Value'],
        ['Units', 4],
      ],
      Summary: [
        ['Metric', 'Value'],
        ['Units', '=Inputs!B2'],
      ],
    })

    try {
      writeFileSync(xlsxPath, exportXlsx(source.exportSnapshot()))
      const responses = await runMcpBin(fileURLToPath(new URL('../work-paper-mcp-stdio-bin.ts', import.meta.url)), [
        '--from-xlsx',
        xlsxPath,
        '--workpaper',
        workpaperPath,
        '--writable',
      ])

      expect(responses[1]?.result?.tools?.map((tool) => tool.name)).toContain('analyze_workbook_risk')
      expect(responses[2]?.result?.structuredContent).toMatchObject({
        schemaVersion: 'bilig-workbook-compatibility-report.v1',
        verified: true,
        input: {
          fileName: 'pricing.xlsx',
        },
      })
    } finally {
      source.dispose()
      rmSync(tempDir, { recursive: true, force: true })
    }
  })
})

async function runMcpBin(binPath: string, args: readonly string[]): Promise<readonly McpTestResponse[]> {
  const child = spawn(process.execPath, ['--import', 'tsx', binPath, ...args], {
    stdio: ['pipe', 'pipe', 'pipe'],
  })
  const stdout: string[] = []
  const stderr: string[] = []
  const exitPromise = new Promise<number | null>((resolve, reject) => {
    const timeout = setTimeout(() => {
      child.kill('SIGKILL')
      reject(new Error('Timed out waiting for @bilig/workpaper MCP XLSX risk smoke test process to exit'))
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
    ]
      .map((request) => JSON.stringify(request))
      .join('\n') + '\n',
  )

  await expect(exitPromise).resolves.toBe(0)
  expect(stderr.join('')).toBe('')
  return stdout
    .join('')
    .trim()
    .split('\n')
    .filter((line) => line.length > 0)
    .map((line) => readMcpTestResponse(JSON.parse(line)))
}

function readMcpTestResponse(value: unknown): McpTestResponse {
  if (typeof value === 'object' && value !== null) {
    return value as McpTestResponse
  }
  throw new Error(`Expected MCP JSON-RPC response object, received ${JSON.stringify(value)}`)
}
