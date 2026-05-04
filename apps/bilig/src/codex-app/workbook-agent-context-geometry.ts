import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { CellRangeRef } from '@bilig/protocol'
import type { WorkbookAgentUiContext, WorkbookViewport } from '@bilig/contracts'

export function workbookAgentViewportToRange(sheetName: string, viewport: WorkbookViewport): CellRangeRef {
  return {
    sheetName,
    startAddress: formatAddress(viewport.rowStart, viewport.colStart),
    endAddress: formatAddress(viewport.rowEnd, viewport.colEnd),
  }
}

export function resolveWorkbookAgentSelectionRange(context: WorkbookAgentUiContext | null): CellRangeRef {
  if (!context) {
    throw new Error('No browser workbook context is attached to this chat session')
  }
  return {
    sheetName: context.selection.sheetName,
    startAddress: context.selection.range?.startAddress ?? context.selection.address,
    endAddress: context.selection.range?.endAddress ?? context.selection.address,
  }
}

export function resolveWorkbookAgentVisibleRange(context: WorkbookAgentUiContext | null): CellRangeRef {
  if (!context) {
    throw new Error('No browser workbook context is attached to this chat session')
  }
  return workbookAgentViewportToRange(context.selection.sheetName, context.viewport)
}

export function resolveWorkbookAgentInspectionTarget(
  context: WorkbookAgentUiContext | null,
  input: {
    readonly sheetName?: string | undefined
    readonly address?: string | undefined
  },
): {
  sheetName: string
  address: string
} {
  if (input.sheetName && input.address) {
    return {
      sheetName: input.sheetName,
      address: input.address,
    }
  }
  if (!context) {
    throw new Error('sheetName and address are required when no browser workbook context exists')
  }
  return context.selection
}

export function workbookAgentViewportAroundAddress(sheetName: string, address: string, base?: WorkbookViewport): WorkbookViewport {
  const parsed = parseCellAddress(address, sheetName)
  if (base) {
    const rowCount = Math.max(1, base.rowEnd - base.rowStart + 1)
    const colCount = Math.max(1, base.colEnd - base.colStart + 1)
    return {
      rowStart: Math.max(0, parsed.row),
      rowEnd: Math.max(0, parsed.row + rowCount - 1),
      colStart: Math.max(0, parsed.col),
      colEnd: Math.max(0, parsed.col + colCount - 1),
    }
  }
  return {
    rowStart: parsed.row,
    rowEnd: parsed.row + 20,
    colStart: parsed.col,
    colEnd: parsed.col + 10,
  }
}
