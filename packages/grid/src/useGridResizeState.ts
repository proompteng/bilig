import type { HeaderSelection } from './gridPointer.js'

export function resolveResizeGuideColumn(input: {
  readonly activeResizeColumn: number | null
  readonly cursor: string
  readonly header: HeaderSelection | null
}): number | null {
  if (input.activeResizeColumn !== null) {
    return input.activeResizeColumn
  }
  if (input.cursor === 'col-resize' && input.header?.kind === 'column') {
    return input.header.index
  }
  return null
}

export function resolveResizeGuideRow(input: {
  readonly activeResizeRow: number | null
  readonly cursor: string
  readonly header: HeaderSelection | null
}): number | null {
  if (input.activeResizeRow !== null) {
    return input.activeResizeRow
  }
  if (input.cursor === 'row-resize' && input.header?.kind === 'row') {
    return input.header.index
  }
  return null
}
