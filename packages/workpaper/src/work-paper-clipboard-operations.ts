import type { CellValue } from '@bilig/protocol'
import { assertRange } from './work-paper-runtime-helpers.js'
import { buildWorkPaperFillRangeData, buildWorkPaperNullMatrixForRange } from './work-paper-fill-helpers.js'
import { createWorkPaperClipboardPayload, type WorkPaperClipboardPayload } from './work-paper-clipboard.js'
import { WorkPaperNothingToPasteError } from './work-paper-errors.js'
import type { RawCellContent, WorkPaperCellAddress, WorkPaperCellRange, WorkPaperChange, WorkPaperSheet } from './work-paper-types.js'

export interface WorkPaperClipboardOperationsRuntime {
  readonly assertReadable: () => void
  readonly assertNotDisposed: () => void
  readonly getClipboard: () => WorkPaperClipboardPayload | null
  readonly setClipboard: (clipboard: WorkPaperClipboardPayload | null) => void
  readonly getRangeSerialized: (range: WorkPaperCellRange) => RawCellContent[][]
  readonly getRangeValues: (range: WorkPaperCellRange) => CellValue[][]
  readonly batch: (operations: () => void) => WorkPaperChange[]
  readonly setCellContents: (address: WorkPaperCellAddress, content: RawCellContent | WorkPaperSheet) => WorkPaperChange[]
  readonly captureChanges: (mutate: () => void) => WorkPaperChange[]
  readonly applySerializedMatrix: (
    targetLeftCorner: WorkPaperCellAddress,
    content: RawCellContent[][],
    sourceAnchor: WorkPaperCellAddress,
  ) => void
}

export interface WorkPaperClipboardOperations {
  readonly copy: (range: WorkPaperCellRange) => CellValue[][]
  readonly cut: (range: WorkPaperCellRange) => CellValue[][]
  readonly paste: (targetLeftCorner: WorkPaperCellAddress) => WorkPaperChange[]
  readonly isClipboardEmpty: () => boolean
  readonly clearClipboard: () => void
  readonly getFillRangeData: (source: WorkPaperCellRange, target: WorkPaperCellRange, offsetsFromTarget?: boolean) => RawCellContent[][]
}

export function createWorkPaperClipboardOperations(runtime: WorkPaperClipboardOperationsRuntime): WorkPaperClipboardOperations {
  const copy = (range: WorkPaperCellRange): CellValue[][] => {
    runtime.assertReadable()
    assertRange(range)
    const serialized = runtime.getRangeSerialized(range)
    const values = runtime.getRangeValues(range)
    runtime.setClipboard(
      createWorkPaperClipboardPayload({
        sourceAnchor: { ...range.start },
        serialized,
        values,
      }),
    )
    return values
  }

  return {
    copy,

    cut(range) {
      runtime.assertReadable()
      const values = copy(range)
      runtime.batch(() => {
        runtime.setCellContents(range.start, buildWorkPaperNullMatrixForRange(range))
      })
      return values
    },

    paste(targetLeftCorner) {
      runtime.assertNotDisposed()
      const clipboard = runtime.getClipboard()
      if (!clipboard) {
        throw new WorkPaperNothingToPasteError()
      }
      return runtime.captureChanges(() => {
        runtime.applySerializedMatrix(targetLeftCorner, clipboard.serialized, clipboard.sourceAnchor)
      })
    },

    isClipboardEmpty() {
      return runtime.getClipboard() === null
    },

    clearClipboard() {
      runtime.setClipboard(null)
    },

    getFillRangeData(source, target, offsetsFromTarget = false) {
      assertRange(source)
      assertRange(target)
      return buildWorkPaperFillRangeData({
        source,
        target,
        sourceSerialized: runtime.getRangeSerialized(source),
        offsetsFromTarget,
      })
    },
  }
}
