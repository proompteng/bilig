import { describe, expect, it } from 'vitest'
import { spawn } from 'node:child_process'
import { fileURLToPath } from 'node:url'

import {
  assertWorkPaperMcpDemoOutput,
  buildDemoWorkPaper,
  createWorkPaperMcpDemoOutput,
  createWorkPaperMcpToolServer,
} from '../work-paper-mcp-server.js'

describe('WorkPaper MCP server', () => {
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
})
