import type { EngineCellMutationRef } from '@bilig/core/headless-runtime'
import type { LiteralInput } from '@bilig/protocol'
import { isBlankRawCellContent, isDeferredBatchLiteralContent, isFormulaContent, stripLeadingEquals } from './work-paper-runtime-helpers.js'
import type { RawCellContent } from './work-paper-types.js'

export function buildWorkPaperRawCellMutation(args: {
  readonly row: number
  readonly col: number
  readonly content: RawCellContent
  readonly rewriteFormulaForStorage: (formula: string) => string
}): EngineCellMutationRef['mutation'] {
  if (isBlankRawCellContent(args.content)) {
    return { kind: 'clearCell', row: args.row, col: args.col }
  }
  if (isFormulaContent(args.content)) {
    return {
      kind: 'setCellFormula',
      row: args.row,
      col: args.col,
      formula: args.rewriteFormulaForStorage(stripLeadingEquals(args.content)),
    }
  }
  return {
    kind: 'setCellValue',
    row: args.row,
    col: args.col,
    value: args.content,
  }
}

export function buildWorkPaperLiteralCellValueMutation(args: {
  readonly row: number
  readonly col: number
  readonly content: LiteralInput
}): EngineCellMutationRef['mutation'] {
  if (isBlankRawCellContent(args.content)) {
    return { kind: 'clearCell', row: args.row, col: args.col }
  }
  return {
    kind: 'setCellValue',
    row: args.row,
    col: args.col,
    value: args.content,
  }
}

export function tryEnqueueWorkPaperLiteralMutation(args: {
  readonly enabled: boolean
  readonly queue: EngineCellMutationRef[]
  readonly sheetId: number
  readonly row: number
  readonly col: number
  readonly content: RawCellContent
  readonly cellIndex: number | undefined
  readonly addPotentialNewCell: () => void
}): boolean {
  if (!args.enabled || !isDeferredBatchLiteralContent(args.content) || isFormulaContent(args.content)) {
    return false
  }
  args.queue.push({
    sheetId: args.sheetId,
    mutation: buildWorkPaperLiteralCellValueMutation({
      row: args.row,
      col: args.col,
      content: args.content,
    }),
    ...(args.cellIndex !== undefined ? { cellIndex: args.cellIndex } : {}),
  })
  if (!isBlankRawCellContent(args.content) && args.cellIndex === undefined) {
    args.addPotentialNewCell()
  }
  return true
}
