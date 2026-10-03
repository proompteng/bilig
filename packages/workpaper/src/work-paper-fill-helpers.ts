import { translateFormulaReferences } from '@bilig/formula'
import type { RawCellContent, WorkPaperCellRange } from './work-paper-types.js'

export function buildWorkPaperNullMatrixForRange(range: WorkPaperCellRange): RawCellContent[][] {
  const height = range.end.row - range.start.row + 1
  const width = range.end.col - range.start.col + 1
  return Array.from({ length: height }, () => Array.from({ length: width }, () => null))
}

export function buildWorkPaperFillRangeData(input: {
  readonly source: WorkPaperCellRange
  readonly target: WorkPaperCellRange
  readonly sourceSerialized: RawCellContent[][]
  readonly offsetsFromTarget: boolean
}): RawCellContent[][] {
  const targetHeight = input.target.end.row - input.target.start.row + 1
  const targetWidth = input.target.end.col - input.target.start.col + 1
  const sourceHeight = Math.max(input.sourceSerialized.length, 1)
  const sourceWidth = Math.max(input.sourceSerialized[0]?.length ?? 0, 1)
  const output: RawCellContent[][] = []
  for (let rowOffset = 0; rowOffset < targetHeight; rowOffset += 1) {
    const row: RawCellContent[] = []
    for (let colOffset = 0; colOffset < targetWidth; colOffset += 1) {
      const targetRow = input.target.start.row + rowOffset
      const targetCol = input.target.start.col + colOffset
      const sourceRow =
        (((targetRow - (input.offsetsFromTarget ? input.target.start.row : input.source.start.row)) % sourceHeight) + sourceHeight) %
        sourceHeight
      const sourceCol =
        (((targetCol - (input.offsetsFromTarget ? input.target.start.col : input.source.start.col)) % sourceWidth) + sourceWidth) %
        sourceWidth
      const raw = input.sourceSerialized[sourceRow]?.[sourceCol] ?? null
      const formulaBodyStart = typeof raw === 'string' ? readFormulaBodyStart(raw) : -1
      if (typeof raw === 'string' && formulaBodyStart >= 0) {
        row.push(
          `${raw.slice(0, formulaBodyStart)}${translateFormulaReferences(
            raw.slice(formulaBodyStart),
            targetRow - (input.source.start.row + sourceRow),
            targetCol - (input.source.start.col + sourceCol),
          )}`,
        )
      } else {
        row.push(raw)
      }
    }
    output.push(row)
  }
  return output
}

function readFormulaBodyStart(raw: string): number {
  let index = 0
  while (index < raw.length && isFormulaPrefixWhitespace(raw.charCodeAt(index))) {
    index += 1
  }
  return raw.charCodeAt(index) === 61 ? index + 1 : -1
}

function isFormulaPrefixWhitespace(code: number): boolean {
  return code === 32 || code === 9 || code === 10 || code === 13
}
