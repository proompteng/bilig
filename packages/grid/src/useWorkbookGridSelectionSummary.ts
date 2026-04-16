import { useLayoutEffect, useMemo } from 'react'
import { formatSelectionSummary, selectionToAddresses } from './gridSelection.js'
import type { GridSelection } from './gridTypes.js'

export function useWorkbookGridSelectionSummary(input: {
  gridSelection: GridSelection
  selectedAddr: string
  onSelectionLabelChange?: ((label: string) => void) | undefined
  onSelectionRangeChange?: ((range: { startAddress: string; endAddress: string }) => void) | undefined
}) {
  const { gridSelection, onSelectionLabelChange, onSelectionRangeChange, selectedAddr } = input
  const selectionSummary = useMemo(() => formatSelectionSummary(gridSelection, selectedAddr), [gridSelection, selectedAddr])
  const selectionRange = useMemo(() => selectionToAddresses(gridSelection, selectedAddr), [gridSelection, selectedAddr])

  useLayoutEffect(() => {
    onSelectionLabelChange?.(selectionSummary)
    onSelectionRangeChange?.(selectionRange)
  }, [onSelectionLabelChange, onSelectionRangeChange, selectionRange, selectionSummary])
}
