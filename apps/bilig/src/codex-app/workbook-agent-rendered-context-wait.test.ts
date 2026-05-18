import { describe, expect, it } from 'vitest'
import type { WorkbookAgentUiContext } from '@bilig/contracts'
import {
  hasRenderedContextAtRevision,
  shouldWaitForRenderedTool,
  waitForWorkbookAgentRenderedContext,
} from './workbook-agent-rendered-context-wait.js'

function context(capturedRevision: number | null): WorkbookAgentUiContext {
  return {
    selection: {
      sheetName: 'Sheet1',
      address: 'A1',
      range: {
        startAddress: 'A1',
        endAddress: 'A1',
      },
    },
    viewport: {
      rowStart: 0,
      rowEnd: 10,
      colStart: 0,
      colEnd: 10,
    },
    rendered: {
      capturedAtUnixMs: 1,
      capturedRevision,
      batchId: 1,
      selection: null,
      visibleRange: null,
    },
  }
}

describe('workbook agent rendered context wait policy', () => {
  it('requires captured revision freshness instead of treating any rendered context as proof', () => {
    expect(hasRenderedContextAtRevision(context(4), 5)).toBe(false)
    expect(hasRenderedContextAtRevision(context(5), 5)).toBe(true)
    expect(hasRenderedContextAtRevision(context(null), 5)).toBe(false)
  })

  it('only waits for tools that claim rendered-browser proof', () => {
    expect(shouldWaitForRenderedTool('read_rendered_selection')).toBe(true)
    expect(shouldWaitForRenderedTool('read_rendered_range')).toBe(true)
    expect(shouldWaitForRenderedTool('apply_and_verify')).toBe(true)
    expect(shouldWaitForRenderedTool('read_range')).toBe(false)
  })

  it('polls until the rendered context reaches the required authoritative revision', async () => {
    const contexts = [context(2), context(4), context(5)]
    const refreshes: number[] = []
    let now = 0

    await expect(
      waitForWorkbookAgentRenderedContext({
        minRevision: 5,
        refreshContext: async () => {
          refreshes.push(now)
          return contexts.shift() ?? context(5)
        },
        delay: async (ms) => {
          now += ms
        },
        now: () => now,
        timeoutMs: 1_000,
        pollIntervalMs: 50,
      }),
    ).resolves.toMatchObject({
      rendered: {
        capturedRevision: 5,
      },
    })
    expect(refreshes).toEqual([0, 50, 100])
  })
})
