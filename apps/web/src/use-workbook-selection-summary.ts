import { useEffect, useState } from 'react'
import { isSelectionAggregateSummary, type GridSelectionSnapshot, type SelectionAggregateSummary } from '@bilig/grid'

export function useWorkbookSelectionSummary(input: {
  readonly selection: GridSelectionSnapshot
  readonly runtimeState: unknown
  readonly runtime: { invoke(method: string, ...args: unknown[]): Promise<unknown> } | null
  readonly onError: (error: unknown) => void
}): SelectionAggregateSummary | null {
  const { selection, runtimeState, runtime, onError } = input
  const { sheetName, range, kind } = selection
  const key = `${sheetName}\u001f${range.startAddress}\u001f${range.endAddress}\u001f${kind}`
  const [result, setResult] = useState<{ key: string; summary: SelectionAggregateSummary } | null>(null)

  useEffect(() => {
    if (!runtime || kind === 'cell') return
    let cancelled = false
    void (async () => {
      try {
        const summary = await runtime.invoke('getSelectionSummary', {
          sheetName,
          startAddress: range.startAddress,
          endAddress: range.endAddress,
        })
        if (cancelled) return
        if (!isSelectionAggregateSummary(summary)) throw new Error('Worker returned an invalid selection summary')
        setResult({ key, summary })
      } catch (error) {
        if (!cancelled) onError(error)
      }
    })()
    return () => {
      cancelled = true
    }
  }, [key, kind, onError, range.startAddress, range.endAddress, runtime, runtimeState, sheetName])
  return result?.key === key ? result.summary : null
}
