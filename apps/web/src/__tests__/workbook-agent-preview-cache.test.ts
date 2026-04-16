import { describe, expect, it, vi } from 'vitest'
import {
  clearWorkbookAgentPreviewCache,
  createWorkbookAgentPreviewRequestKey,
  loadWorkbookAgentPreview,
  readCachedWorkbookAgentPreview,
} from '../workbook-agent-preview-cache.js'

describe('workbook agent preview cache', () => {
  it('reuses in-flight and settled previews for the same bundle selection', async () => {
    clearWorkbookAgentPreviewCache()
    const requestKey = createWorkbookAgentPreviewRequestKey({
      bundle: {
        id: 'bundle-1',
        baseRevision: 7,
      },
      commandIndexes: [0, 2],
    })
    const preview = {
      ranges: [],
      structuralChanges: [],
      cellDiffs: [],
      effectSummary: {
        displayedCellDiffCount: 0,
        truncatedCellDiffs: false,
        inputChangeCount: 0,
        formulaChangeCount: 0,
        styleChangeCount: 0,
        numberFormatChangeCount: 0,
        structuralChangeCount: 0,
      },
    } as const
    let resolvePreview: ((value: typeof preview) => void) | null = null
    const load = vi.fn(
      () =>
        new Promise<typeof preview>((resolve) => {
          resolvePreview = resolve
        }),
    )

    const firstRequest = loadWorkbookAgentPreview({
      requestKey,
      load,
    })
    const secondRequest = loadWorkbookAgentPreview({
      requestKey,
      load,
    })

    expect(load).toHaveBeenCalledTimes(1)
    expect(secondRequest).toBe(firstRequest)

    resolvePreview?.(preview)
    await expect(firstRequest).resolves.toEqual(preview)
    expect(readCachedWorkbookAgentPreview(requestKey)).toEqual(preview)

    await expect(
      loadWorkbookAgentPreview({
        requestKey,
        load,
      }),
    ).resolves.toEqual(preview)
    expect(load).toHaveBeenCalledTimes(1)
  })
})
