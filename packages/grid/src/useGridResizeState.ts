import type { HeaderSelection } from './gridPointer.js'

export function resolveResizeGuideColumn(input: {
  readonly activeResizeColumn: number | null
  readonly cursor: string
  readonly header: HeaderSelection | null
}): number | null {
  return input.activeResizeColumn ?? (input.cursor === 'col-resize' && input.header?.kind === 'column' ? input.header.index : null)
}

export function resolveResizeGuideRow(input: {
  readonly activeResizeRow: number | null
  readonly cursor: string
  readonly header: HeaderSelection | null
}): number | null {
  return input.activeResizeRow ?? (input.cursor === 'row-resize' && input.header?.kind === 'row' ? input.header.index : null)
}
