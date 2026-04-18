import type { GridSelectionSnapshot } from '@bilig/grid'
import type { WorkbookAgentUiContext } from '@bilig/contracts'
import type { Viewport } from '@bilig/protocol'
import type { WorkerRuntimeSelection } from './runtime-session.js'

export function createSingleCellSelectionSnapshot(selection: WorkerRuntimeSelection): GridSelectionSnapshot {
  return {
    sheetName: selection.sheetName,
    address: selection.address,
    kind: 'cell',
    range: {
      startAddress: selection.address,
      endAddress: selection.address,
    },
  }
}

export function buildWorkbookAgentContext(input: {
  readonly selection: GridSelectionSnapshot
  readonly viewport: Viewport
}): WorkbookAgentUiContext {
  return {
    selection: {
      sheetName: input.selection.sheetName,
      address: input.selection.address,
      range: {
        startAddress: input.selection.range.startAddress,
        endAddress: input.selection.range.endAddress,
      },
    },
    viewport: input.viewport,
  }
}
