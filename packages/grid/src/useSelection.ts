import { useCallback, useSyncExternalStore } from 'react'
import { selectors, type SpreadsheetEngine } from '@bilig/core'

export function useSelection(engine: SpreadsheetEngine) {
  const selection = useSyncExternalStore(
    (listener) => engine.subscribeSelection(listener),
    () => selectors.selectSelectionState(engine),
    () => selectors.selectSelectionState(engine),
  )
  const select = useCallback(
    (nextSheetName: string, nextAddress: string | null) => {
      engine.setSelection(nextSheetName, nextAddress)
    },
    [engine],
  )

  return {
    sheetName: selection.sheetName,
    address: selection.address,
    select,
  }
}
