import { formatAddress, parseCellAddress } from '@bilig/formula'
import { ValueTag, type CellRangeRef } from '@bilig/protocol'
import type { GridEngineLike, SelectionAggregateSummary } from '@bilig/grid'

export function collectWorkbookSelectionSummary(
  engine: Pick<GridEngineLike, 'workbook' | 'getCell'>,
  range: CellRangeRef,
): SelectionAggregateSummary {
  const start = parseCellAddress(range.startAddress, range.sheetName)
  const end = parseCellAddress(range.endAddress, range.sheetName)
  const rowStart = Math.min(start.row, end.row)
  const rowEnd = Math.max(start.row, end.row)
  const colStart = Math.min(start.col, end.col)
  const colEnd = Math.max(start.col, end.col)
  let nonEmptyCount = 0
  let numericCount = 0
  let sum = 0
  let correction = 0
  let min: number | null = null
  let max: number | null = null

  engine.workbook.getSheet(range.sheetName)?.grid.forEachCellEntry((_index, row, col) => {
    if (row < rowStart || row > rowEnd || col < colStart || col > colEnd) return
    const { value } = engine.getCell(range.sheetName, formatAddress(row, col))
    if (value.tag !== ValueTag.Empty) nonEmptyCount += 1
    if (value.tag !== ValueTag.Number || !Number.isFinite(value.value)) return
    numericCount += 1
    const next = sum + value.value
    if (Number.isFinite(next)) {
      correction += Math.abs(sum) >= Math.abs(value.value) ? sum - next + value.value : value.value - next + sum
    } else {
      correction = 0
    }
    sum = next
    min = min === null ? value.value : Math.min(min, value.value)
    max = max === null ? value.value : Math.max(max, value.value)
  })
  return { nonEmptyCount, numericCount, sum: sum + correction, min, max }
}
