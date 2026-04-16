import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { WorkPaperCellAddress, WorkPaperSheet, WorkPaperSheetDimensions } from './work-paper-types.js'

export function replaceWorkPaperSheetContent(args: {
  sheetId: number
  sheetName: string
  content: WorkPaperSheet
  duringInitialization: boolean
  listSpills: () => Array<{
    sheetName: string
    address: string
    rows: number
    cols: number
  }>
  getSheetDimensions: (sheetId: number) => WorkPaperSheetDimensions
  clearRange: (input: { sheetName: string; startAddress: string; endAddress: string }) => void
  applyMatrixContents: (
    address: WorkPaperCellAddress,
    content: WorkPaperSheet,
    options: {
      captureUndo: boolean
      deferLiteralAddresses: ReadonlySet<string>
      skipNulls: boolean
    },
  ) => void
  clearHistoryStacks: () => void
  getUndoStackLength: () => number
  mergeUndoHistory: (stackStart: number) => void
}): void {
  const undoStackStart = args.duringInitialization ? 0 : args.getUndoStackLength()
  const deferredLiteralAddresses = new Set<string>()
  args
    .listSpills()
    .filter((spill) => spill.sheetName === args.sheetName)
    .forEach((spill) => {
      const owner = parseCellAddress(spill.address, spill.sheetName)
      for (let rowOffset = 0; rowOffset < spill.rows; rowOffset += 1) {
        for (let colOffset = 0; colOffset < spill.cols; colOffset += 1) {
          if (rowOffset === 0 && colOffset === 0) {
            continue
          }
          deferredLiteralAddresses.add(formatAddress(owner.row + rowOffset, owner.col + colOffset))
        }
      }
    })
  const dimensions = args.getSheetDimensions(args.sheetId)
  if (dimensions.width > 0 && dimensions.height > 0) {
    args.clearRange({
      sheetName: args.sheetName,
      startAddress: 'A1',
      endAddress: formatAddress(dimensions.height - 1, dimensions.width - 1),
    })
  }
  args.applyMatrixContents({ sheet: args.sheetId, row: 0, col: 0 }, args.content, {
    captureUndo: !args.duringInitialization,
    deferLiteralAddresses: deferredLiteralAddresses,
    skipNulls: true,
  })
  if (args.duringInitialization) {
    args.clearHistoryStacks()
    return
  }
  args.mergeUndoHistory(undoStackStart)
}
