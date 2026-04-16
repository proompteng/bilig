import { useEffect, useMemo, useState } from 'react'
import { queries } from '@bilig/zero-sync'
import { normalizeWorkbookChangeRows, selectWorkbookChangeEntries, type WorkbookChangeEntry } from './workbook-changes-model.js'

interface ZeroLiveView<T> {
  readonly data: T
  addListener(listener: (value: T) => void): () => void
  destroy(): void
}

export interface ZeroWorkbookChangeQuerySource {
  materialize(query: unknown): unknown
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function isZeroLiveView<T>(value: unknown): value is ZeroLiveView<T> {
  return isRecord(value) && 'data' in value && typeof value['addListener'] === 'function' && typeof value['destroy'] === 'function'
}

export function useWorkbookChanges(input: {
  readonly documentId: string
  readonly sheetNames: readonly string[]
  readonly zero: ZeroWorkbookChangeQuerySource
  readonly enabled: boolean
}): readonly WorkbookChangeEntry[] {
  const { documentId, enabled, sheetNames, zero } = input
  const [rows, setRows] = useState([] as readonly ReturnType<typeof normalizeWorkbookChangeRows>[number][])

  useEffect(() => {
    if (!enabled) {
      setRows([])
      return
    }
    const view = zero.materialize(queries.workbookChange.byWorkbook({ documentId }))
    if (!isZeroLiveView<unknown>(view)) {
      throw new Error('Zero workbook changes query returned an invalid live view')
    }
    const publishRows = (value: unknown) => {
      setRows(normalizeWorkbookChangeRows(value))
    }
    publishRows(view.data)
    const cleanup = view.addListener((value) => {
      publishRows(value)
    })
    return () => {
      cleanup()
      view.destroy()
    }
  }, [documentId, enabled, zero])

  return useMemo(
    () =>
      selectWorkbookChangeEntries({
        rows,
        knownSheetNames: sheetNames,
      }),
    [rows, sheetNames],
  )
}
