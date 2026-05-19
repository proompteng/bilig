import { BLOCK_COLS, BLOCK_ROWS, type SheetRecord, type SpreadsheetEngine } from '@bilig/core/headless-runtime'
import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { WorkPaperCellRange, WorkPaperRangeValueBlock } from './work-paper-types.js'

const EMPTY_CELL_VALUE: CellValue = Object.freeze({ tag: ValueTag.Empty })
const FAST_RANGE_MATRIX_AREA_LIMIT = 262_144
const FAST_RANGE_VALUE_BLOCK_AREA_LIMIT = 2_097_152

export function readFastRangeValues(engine: SpreadsheetEngine, range: WorkPaperCellRange): CellValue[][] | undefined {
  const sheet = engine.workbook.getSheetById(range.start.sheet)
  if (!sheet) {
    return undefined
  }
  const height = range.end.row - range.start.row + 1
  const width = range.end.col - range.start.col + 1
  if (height <= 0 || width <= 0 || height * width > FAST_RANGE_MATRIX_AREA_LIMIT) {
    return undefined
  }
  if (sheet.structureVersion !== 1) {
    return readFastLogicalRangeValues(engine, sheet, range, height, width)
  }

  const rows: CellValue[][] = []
  rows.length = height
  for (let rowOffset = 0; rowOffset < height; rowOffset += 1) {
    const row: CellValue[] = []
    row.length = width
    rows[rowOffset] = row
  }
  const readCellValue = createFastCellValueReader(engine)
  let filledCells = 0
  const blockRowStart = Math.floor(range.start.row / BLOCK_ROWS)
  const blockRowEnd = Math.floor(range.end.row / BLOCK_ROWS)
  const blockColStart = Math.floor(range.start.col / BLOCK_COLS)
  const blockColEnd = Math.floor(range.end.col / BLOCK_COLS)
  for (let blockRow = blockRowStart; blockRow <= blockRowEnd; blockRow += 1) {
    const absoluteBlockRow = blockRow * BLOCK_ROWS
    const localRowStart = Math.max(range.start.row - absoluteBlockRow, 0)
    const localRowEnd = Math.min(range.end.row - absoluteBlockRow, BLOCK_ROWS - 1)
    for (let blockCol = blockColStart; blockCol <= blockColEnd; blockCol += 1) {
      const block = sheet.grid.blocks.get(blockRow * 1_000_000 + blockCol)
      if (!block) {
        continue
      }
      const absoluteBlockCol = blockCol * BLOCK_COLS
      const localColStart = Math.max(range.start.col - absoluteBlockCol, 0)
      const localColEnd = Math.min(range.end.col - absoluteBlockCol, BLOCK_COLS - 1)
      for (let localRow = localRowStart; localRow <= localRowEnd; localRow += 1) {
        const row = rows[absoluteBlockRow + localRow - range.start.row]!
        const blockRowOffset = localRow * BLOCK_COLS
        for (let localCol = localColStart; localCol <= localColEnd; localCol += 1) {
          const encodedCellIndex = block[blockRowOffset + localCol]!
          if (encodedCellIndex === 0) {
            continue
          }
          const cellIndex = encodedCellIndex - 1
          const outputCol = absoluteBlockCol + localCol - range.start.col
          row[outputCol] = readCellValue(cellIndex)
          filledCells += 1
        }
      }
    }
  }
  if (filledCells < height * width) {
    for (let rowOffset = 0; rowOffset < height; rowOffset += 1) {
      const row = rows[rowOffset]!
      for (let colOffset = 0; colOffset < width; colOffset += 1) {
        row[colOffset] ??= EMPTY_CELL_VALUE
      }
    }
  }
  return rows
}

export function readFastRangeValueBlock(engine: SpreadsheetEngine, range: WorkPaperCellRange): WorkPaperRangeValueBlock | undefined {
  const sheet = engine.workbook.getSheetById(range.start.sheet)
  if (!sheet) {
    return undefined
  }
  const rowCount = range.end.row - range.start.row + 1
  const colCount = range.end.col - range.start.col + 1
  if (rowCount <= 0 || colCount <= 0 || rowCount * colCount > FAST_RANGE_VALUE_BLOCK_AREA_LIMIT) {
    return undefined
  }
  const block = createEmptyValueBlock(rowCount, colCount)
  const appendString = createValueBlockCellWriter(engine, block)
  if (sheet.structureVersion !== 1) {
    for (let rowOffset = 0; rowOffset < rowCount; rowOffset += 1) {
      const visibleRow = range.start.row + rowOffset
      const outputRowOffset = rowOffset * colCount
      for (let colOffset = 0; colOffset < colCount; colOffset += 1) {
        const cellIndex = sheet.grid.get(visibleRow, range.start.col + colOffset)
        if (cellIndex !== -1) {
          appendString(cellIndex, outputRowOffset + colOffset)
        }
      }
    }
    return finalizeValueBlock(block)
  }

  const blockRowStart = Math.floor(range.start.row / BLOCK_ROWS)
  const blockRowEnd = Math.floor(range.end.row / BLOCK_ROWS)
  const blockColStart = Math.floor(range.start.col / BLOCK_COLS)
  const blockColEnd = Math.floor(range.end.col / BLOCK_COLS)
  for (let blockRow = blockRowStart; blockRow <= blockRowEnd; blockRow += 1) {
    const absoluteBlockRow = blockRow * BLOCK_ROWS
    const localRowStart = Math.max(range.start.row - absoluteBlockRow, 0)
    const localRowEnd = Math.min(range.end.row - absoluteBlockRow, BLOCK_ROWS - 1)
    for (let blockCol = blockColStart; blockCol <= blockColEnd; blockCol += 1) {
      const gridBlock = sheet.grid.blocks.get(blockRow * 1_000_000 + blockCol)
      if (!gridBlock) {
        continue
      }
      const absoluteBlockCol = blockCol * BLOCK_COLS
      const localColStart = Math.max(range.start.col - absoluteBlockCol, 0)
      const localColEnd = Math.min(range.end.col - absoluteBlockCol, BLOCK_COLS - 1)
      for (let localRow = localRowStart; localRow <= localRowEnd; localRow += 1) {
        const outputRowOffset = (absoluteBlockRow + localRow - range.start.row) * colCount
        const blockRowOffset = localRow * BLOCK_COLS
        for (let localCol = localColStart; localCol <= localColEnd; localCol += 1) {
          const encodedCellIndex = gridBlock[blockRowOffset + localCol]!
          if (encodedCellIndex !== 0) {
            appendString(encodedCellIndex - 1, outputRowOffset + absoluteBlockCol + localCol - range.start.col)
          }
        }
      }
    }
  }
  return finalizeValueBlock(block)
}

export function materializeRangeValueBlock(block: WorkPaperRangeValueBlock): CellValue[][] {
  const rows: CellValue[][] = []
  rows.length = block.rowCount
  for (let rowOffset = 0; rowOffset < block.rowCount; rowOffset += 1) {
    const row: CellValue[] = []
    row.length = block.colCount
    rows[rowOffset] = row
    const outputRowOffset = rowOffset * block.colCount
    for (let colOffset = 0; colOffset < block.colCount; colOffset += 1) {
      const index = outputRowOffset + colOffset
      switch ((block.tags[index] as ValueTag | undefined) ?? ValueTag.Empty) {
        case ValueTag.Number:
          row[colOffset] = { tag: ValueTag.Number, value: block.numbers[index] ?? 0 }
          break
        case ValueTag.Boolean:
          row[colOffset] = {
            tag: ValueTag.Boolean,
            value: (block.numbers[index] ?? 0) !== 0,
          }
          break
        case ValueTag.String: {
          const stringId = block.stringIds[index] ?? 0
          row[colOffset] = {
            tag: ValueTag.String,
            value: stringId === 0 ? '' : (block.strings?.get(stringId) ?? ''),
            stringId,
          }
          break
        }
        case ValueTag.Error:
          row[colOffset] = {
            tag: ValueTag.Error,
            code: (block.errors[index] as ErrorCode | undefined) ?? ErrorCode.None,
          }
          break
        case ValueTag.Empty:
        default:
          row[colOffset] = EMPTY_CELL_VALUE
          break
      }
    }
  }
  return rows
}

function readFastLogicalRangeValues(
  engine: SpreadsheetEngine,
  sheet: SheetRecord,
  range: WorkPaperCellRange,
  height: number,
  width: number,
): CellValue[][] {
  const readCellValue = createFastCellValueReader(engine)
  const rows: CellValue[][] = []
  rows.length = height
  for (let rowOffset = 0; rowOffset < height; rowOffset += 1) {
    const row: CellValue[] = []
    row.length = width
    rows[rowOffset] = row
    const visibleRow = range.start.row + rowOffset
    for (let colOffset = 0; colOffset < width; colOffset += 1) {
      const cellIndex = sheet.grid.get(visibleRow, range.start.col + colOffset)
      row[colOffset] = cellIndex === -1 ? EMPTY_CELL_VALUE : readCellValue(cellIndex)
    }
  }
  return rows
}

function createFastCellValueReader(engine: SpreadsheetEngine): (cellIndex: number) => CellValue {
  const cellStore = engine.workbook.cellStore
  const cellTags = cellStore.tags
  const cellNumbers = cellStore.numbers
  const cellStringIds = cellStore.stringIds
  const cellErrors = cellStore.errors
  const strings = engine.strings
  return (cellIndex: number): CellValue => {
    switch ((cellTags[cellIndex] as ValueTag | undefined) ?? ValueTag.Empty) {
      case ValueTag.Number:
        return { tag: ValueTag.Number, value: cellNumbers[cellIndex] ?? 0 }
      case ValueTag.Boolean:
        return {
          tag: ValueTag.Boolean,
          value: (cellNumbers[cellIndex] ?? 0) !== 0,
        }
      case ValueTag.String: {
        const stringId = cellStringIds[cellIndex] ?? 0
        return {
          tag: ValueTag.String,
          value: stringId === 0 ? '' : strings.get(stringId),
          stringId,
        }
      }
      case ValueTag.Error:
        return {
          tag: ValueTag.Error,
          code: (cellErrors[cellIndex] as ErrorCode | undefined) ?? ErrorCode.None,
        }
      case ValueTag.Empty:
      default:
        return EMPTY_CELL_VALUE
    }
  }
}

function createEmptyValueBlock(rowCount: number, colCount: number): WorkPaperRangeValueBlock & { strings?: Map<number, string> } {
  const area = rowCount * colCount
  return {
    rowCount,
    colCount,
    tags: new Uint8Array(area),
    numbers: new Float64Array(area),
    stringIds: new Uint32Array(area),
    errors: new Uint16Array(area),
  }
}

function createValueBlockCellWriter(
  engine: SpreadsheetEngine,
  block: WorkPaperRangeValueBlock & { strings?: Map<number, string> },
): (cellIndex: number, outputIndex: number) => void {
  const cellStore = engine.workbook.cellStore
  const cellTags = cellStore.tags
  const cellNumbers = cellStore.numbers
  const cellStringIds = cellStore.stringIds
  const cellErrors = cellStore.errors
  const strings = engine.strings
  return (cellIndex: number, outputIndex: number): void => {
    const tag = (cellTags[cellIndex] as ValueTag | undefined) ?? ValueTag.Empty
    block.tags[outputIndex] = tag
    switch (tag) {
      case ValueTag.Number:
      case ValueTag.Boolean:
        block.numbers[outputIndex] = cellNumbers[cellIndex] ?? 0
        break
      case ValueTag.String: {
        const stringId = cellStringIds[cellIndex] ?? 0
        block.stringIds[outputIndex] = stringId
        if (stringId !== 0) {
          block.strings ??= new Map()
          if (!block.strings.has(stringId)) {
            block.strings.set(stringId, strings.get(stringId))
          }
        }
        break
      }
      case ValueTag.Error:
        block.errors[outputIndex] = (cellErrors[cellIndex] as ErrorCode | undefined) ?? ErrorCode.None
        break
      case ValueTag.Empty:
        break
    }
  }
}

function finalizeValueBlock(block: WorkPaperRangeValueBlock & { strings?: Map<number, string> }): WorkPaperRangeValueBlock {
  return block
}
