import { MAX_COLS, MAX_ROWS } from '@bilig/protocol'
import { WorkPaperInvalidArgumentsError } from './work-paper-errors.js'
import { normalizeAxisIntervals, normalizeAxisSwapMappings } from './work-paper-axis-helpers.js'
import { assertRange, assertRowAndColumn, isCellRange } from './work-paper-runtime-helpers.js'
import type {
  RawCellContent,
  WorkPaperAddressLike,
  WorkPaperAxisInterval,
  WorkPaperAxisSwapMapping,
  WorkPaperCellAddress,
  WorkPaperCellRange,
  WorkPaperConfig,
  WorkPaperSheet,
} from './work-paper-types.js'

type WorkPaperAxis = 'row' | 'column'

interface WorkPaperCapabilityContext {
  readonly config: WorkPaperConfig
  readonly requireSheet: (sheetId: number) => void
  readonly doesSheetExist: (sheetName: string) => boolean
  readonly getSheetIdByName: (sheetName: string) => number | undefined
  readonly isCellContentProtected?: (address: WorkPaperCellAddress) => boolean
}

const limitForAxis = (config: WorkPaperConfig, axis: WorkPaperAxis): number =>
  axis === 'row' ? (config.maxRows ?? MAX_ROWS) : (config.maxColumns ?? MAX_COLS)

const validateSheetName = (sheetName: string): string => {
  const trimmed = sheetName.trim()
  if (trimmed.length === 0) {
    throw new WorkPaperInvalidArgumentsError('Sheet name must be non-empty')
  }
  return trimmed
}

function protectedCellContentsIntersect(
  context: WorkPaperCapabilityContext,
  sheet: number,
  startRow: number,
  startCol: number,
  rowCount: number,
  colCount: number,
): boolean {
  if (!context.isCellContentProtected || rowCount <= 0 || colCount <= 0) {
    return false
  }
  for (let rowOffset = 0; rowOffset < rowCount; rowOffset += 1) {
    for (let colOffset = 0; colOffset < colCount; colOffset += 1) {
      if (context.isCellContentProtected({ sheet, row: startRow + rowOffset, col: startCol + colOffset })) {
        return true
      }
    }
  }
  return false
}

export function isWorkPaperSetCellContentsPossible(
  context: WorkPaperCapabilityContext,
  addressOrRange: WorkPaperAddressLike,
  content?: RawCellContent | WorkPaperSheet,
): boolean {
  if (isCellRange(addressOrRange)) {
    assertRange(addressOrRange)
    context.requireSheet(addressOrRange.start.sheet)
    return (
      addressOrRange.end.row < (context.config.maxRows ?? MAX_ROWS) &&
      addressOrRange.end.col < (context.config.maxColumns ?? MAX_COLS) &&
      !protectedCellContentsIntersect(
        context,
        addressOrRange.start.sheet,
        addressOrRange.start.row,
        addressOrRange.start.col,
        addressOrRange.end.row - addressOrRange.start.row + 1,
        addressOrRange.end.col - addressOrRange.start.col + 1,
      )
    )
  }

  context.requireSheet(addressOrRange.sheet)
  assertRowAndColumn(addressOrRange.row, 'address.row')
  assertRowAndColumn(addressOrRange.col, 'address.col')
  if (content === undefined) {
    return (
      addressOrRange.row < (context.config.maxRows ?? MAX_ROWS) &&
      addressOrRange.col < (context.config.maxColumns ?? MAX_COLS) &&
      !context.isCellContentProtected?.(addressOrRange)
    )
  }
  if (Array.isArray(content)) {
    if (!content.every((row) => Array.isArray(row))) {
      throw new WorkPaperInvalidArgumentsError('Content matrix must be a two-dimensional array')
    }
    const height = content.length
    const width = Math.max(0, ...content.map((row) => row.length))
    return (
      addressOrRange.row + height <= (context.config.maxRows ?? MAX_ROWS) &&
      addressOrRange.col + width <= (context.config.maxColumns ?? MAX_COLS) &&
      !protectedCellContentsIntersect(context, addressOrRange.sheet, addressOrRange.row, addressOrRange.col, height, width)
    )
  }
  return (
    addressOrRange.row < (context.config.maxRows ?? MAX_ROWS) &&
    addressOrRange.col < (context.config.maxColumns ?? MAX_COLS) &&
    !context.isCellContentProtected?.(addressOrRange)
  )
}

export function isWorkPaperAxisSwapPossible(
  context: WorkPaperCapabilityContext,
  axis: WorkPaperAxis,
  sheetId: number,
  firstOrMappings: number | readonly WorkPaperAxisSwapMapping[],
  second?: number,
): boolean {
  context.requireSheet(sheetId)
  const limit = limitForAxis(context.config, axis)
  const [firstLabel, secondLabel] = axis === 'row' ? ['rowA', 'rowB'] : ['columnA', 'columnB']
  return normalizeAxisSwapMappings(axis, firstOrMappings, second).every(([first, mappedSecond]) => {
    assertRowAndColumn(first, firstLabel)
    assertRowAndColumn(mappedSecond, secondLabel)
    return first < limit && mappedSecond < limit
  })
}

export function isWorkPaperAxisOrderPossible(
  context: WorkPaperCapabilityContext,
  axis: WorkPaperAxis,
  sheetId: number,
  order: readonly number[],
): boolean {
  context.requireSheet(sheetId)
  if (new Set(order).size !== order.length || order.some((value) => !Number.isInteger(value) || value < 0)) {
    return false
  }
  const limit = limitForAxis(context.config, axis)
  return order.every((value) => value < limit)
}

export function isWorkPaperAxisIntervalEditPossible(
  context: WorkPaperCapabilityContext,
  axis: WorkPaperAxis,
  sheetId: number,
  startOrInterval: number | WorkPaperAxisInterval,
  countOrInterval: number | WorkPaperAxisInterval | undefined,
  restIntervals: readonly WorkPaperAxisInterval[],
): boolean {
  context.requireSheet(sheetId)
  const limit = limitForAxis(context.config, axis)
  return normalizeAxisIntervals(startOrInterval, countOrInterval, restIntervals).every(([start, count]) => {
    assertRowAndColumn(start, 'start')
    assertRowAndColumn(count, 'count')
    return count > 0 && start + count <= limit
  })
}

export function isWorkPaperMoveCellsPossible(
  context: WorkPaperCapabilityContext,
  source: WorkPaperCellRange,
  target: WorkPaperCellAddress,
): boolean {
  assertRange(source)
  assertRowAndColumn(target.sheet, 'target.sheet')
  assertRowAndColumn(target.row, 'target.row')
  assertRowAndColumn(target.col, 'target.col')
  if (source.start.sheet !== target.sheet) {
    return false
  }
  const height = source.end.row - source.start.row + 1
  const width = source.end.col - source.start.col + 1
  return target.row + height <= (context.config.maxRows ?? MAX_ROWS) && target.col + width <= (context.config.maxColumns ?? MAX_COLS)
}

export function isWorkPaperMoveAxisPossible(
  context: WorkPaperCapabilityContext,
  axis: WorkPaperAxis,
  sheetId: number,
  start: number,
  count: number,
  target: number,
): boolean {
  context.requireSheet(sheetId)
  assertRowAndColumn(start, 'start')
  assertRowAndColumn(count, 'count')
  assertRowAndColumn(target, 'target')
  const limit = limitForAxis(context.config, axis)
  return count > 0 && start + count <= limit && target + count <= limit
}

export function isWorkPaperSheetNameAvailable(context: WorkPaperCapabilityContext, sheetName: string, currentSheetId?: number): boolean {
  const trimmed = validateSheetName(sheetName)
  if (currentSheetId === undefined) {
    return !context.doesSheetExist(trimmed)
  }
  context.requireSheet(currentSheetId)
  const existingId = context.getSheetIdByName(trimmed)
  return existingId === undefined || existingId === currentSheetId
}

export function isWorkPaperSheetContentReplaceable(context: WorkPaperCapabilityContext, sheetId: number, content: WorkPaperSheet): boolean {
  context.requireSheet(sheetId)
  if (!content.every((row) => Array.isArray(row))) {
    throw new WorkPaperInvalidArgumentsError('Sheet content must be a two-dimensional array')
  }
  const height = content.length
  const width = Math.max(0, ...content.map((row) => row.length))
  return height <= (context.config.maxRows ?? MAX_ROWS) && width <= (context.config.maxColumns ?? MAX_COLS)
}
