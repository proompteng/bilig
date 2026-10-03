import type { CellValue } from '@bilig/protocol'
import { cloneCellValue } from './work-paper-runtime-helpers.js'
import type { RawCellContent, WorkPaperCellAddress } from './work-paper-types.js'

export interface WorkPaperClipboardPayload {
  readonly sourceAnchor: WorkPaperCellAddress
  readonly serialized: RawCellContent[][]
  readonly values: CellValue[][]
}

export function createWorkPaperClipboardPayload(input: {
  readonly sourceAnchor: WorkPaperCellAddress
  readonly serialized: RawCellContent[][]
  readonly values: CellValue[][]
}): WorkPaperClipboardPayload {
  return {
    sourceAnchor: { ...input.sourceAnchor },
    serialized: input.serialized,
    values: input.values,
  }
}

export function cloneWorkPaperClipboardPayload(payload: WorkPaperClipboardPayload | null): WorkPaperClipboardPayload | null {
  if (payload === null) {
    return null
  }
  return {
    sourceAnchor: { ...payload.sourceAnchor },
    serialized: payload.serialized.map((row) => [...row]),
    values: payload.values.map((row) => row.map((value) => cloneCellValue(value))),
  }
}
