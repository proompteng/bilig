import type { SpreadsheetEngine, SheetRecord } from '@bilig/core/headless-runtime'
import { ValueTag, type CellRangeRef, type CellSnapshot, type CellValue } from '@bilig/protocol'
import {
  collectSerializedWorkPaperSheets,
  collectWorkPaperSheetsByName,
  readWorkPaperSheetRange,
  type WorkPaperSheetFormulas,
  type WorkPaperSheetValues,
} from './work-paper-sheet-read.js'
import {
  formatWorkPaperCellAddressText,
  formatWorkPaperCellRangeText,
  parseWorkPaperCellAddressText,
  parseWorkPaperCellRangeText,
  resolveDefaultWorkPaperSheetName,
} from './work-paper-address-format.js'
import type { WorkPaperSheetDimensionCache } from './work-paper-sheet-dimension-cache.js'
import type {
  RawCellContent,
  WorkPaperAddressFormatOptions,
  WorkPaperCellAddress,
  WorkPaperCellRange,
  WorkPaperRangeValueBlock,
  WorkPaperCellType,
  WorkPaperCellValueDetailedType,
  WorkPaperCellValueType,
  WorkPaperSheetDimensions,
} from './work-paper-types.js'
import {
  getVisibleWorkPaperCellIndexInSheet,
  readWorkPaperCellClassification,
  readWorkPaperCellFormula,
  readWorkPaperCellFormulaByIndex,
  readWorkPaperCellHasSimpleValue,
  readWorkPaperCellHyperlink,
  readWorkPaperCellPartOfArray,
  readWorkPaperCellValue,
  readWorkPaperCellValueDetailedType,
  readWorkPaperCellValueType,
  readWorkPaperRangeValueBlock,
  readWorkPaperRangeValues,
} from './work-paper-cell-read.js'
import { assertRange } from './work-paper-runtime-helpers.js'
import { workPaperFormulaMayResizeDynamically } from './work-paper-sheet-inspection.js'

export interface WorkPaperReadOperations {
  readonly getCellValue: (address: WorkPaperCellAddress) => CellValue
  readonly getCellFormula: (address: WorkPaperCellAddress) => string | undefined
  readonly getCellHyperlink: (address: WorkPaperCellAddress) => string | undefined
  readonly getCellSerialized: (address: WorkPaperCellAddress) => RawCellContent
  readonly getRangeValues: (range: WorkPaperCellRange) => CellValue[][]
  readonly getRangeValueBlock: (range: WorkPaperCellRange) => WorkPaperRangeValueBlock
  readonly getRangeFormulas: (range: WorkPaperCellRange) => WorkPaperSheetFormulas
  readonly getRangeSerialized: (range: WorkPaperCellRange) => RawCellContent[][]
  readonly getSheetValues: (sheetId: number) => WorkPaperSheetValues
  readonly getSheetFormulas: (sheetId: number) => WorkPaperSheetFormulas
  readonly getSheetSerialized: (sheetId: number) => RawCellContent[][]
  readonly getAllSheetsValues: () => Record<string, CellValue[][]>
  readonly getAllSheetsFormulas: () => Record<string, WorkPaperSheetFormulas>
  readonly getAllSheetsSerialized: () => Record<string, RawCellContent[][]>
  readonly getAllSheetsDimensions: () => Record<string, WorkPaperSheetDimensions>
  readonly getSheetDimensions: (sheetId: number) => WorkPaperSheetDimensions
  readonly simpleCellAddressFromString: (value: string, defaultSheetId?: number) => WorkPaperCellAddress | undefined
  readonly simpleCellRangeFromString: (value: string, defaultSheetId?: number) => WorkPaperCellRange | undefined
  readonly simpleCellAddressToString: (
    address: WorkPaperCellAddress,
    optionsOrContextSheetId?: WorkPaperAddressFormatOptions | number,
  ) => string
  readonly simpleCellRangeToString: (range: WorkPaperCellRange, optionsOrContextSheetId?: WorkPaperAddressFormatOptions | number) => string
  readonly getCellType: (address: WorkPaperCellAddress) => WorkPaperCellType
  readonly doesCellHaveSimpleValue: (address: WorkPaperCellAddress) => boolean
  readonly doesCellHaveFormula: (address: WorkPaperCellAddress) => boolean
  readonly isCellEmpty: (address: WorkPaperCellAddress) => boolean
  readonly isCellPartOfArray: (address: WorkPaperCellAddress) => boolean
  readonly getCellValueType: (address: WorkPaperCellAddress) => WorkPaperCellValueType
  readonly getCellValueDetailedType: (address: WorkPaperCellAddress) => WorkPaperCellValueDetailedType
  readonly getCellValueFormat: (address: WorkPaperCellAddress) => string | undefined
}

interface WorkPaperReadOperationsRuntime {
  readonly getEngine: () => SpreadsheetEngine
  readonly getSheetDimensionCache: () => WorkPaperSheetDimensionCache
  readonly assertReadable: () => void
  readonly assertNotDisposed: () => void
  readonly prepareReadableState: () => void
  readonly flushPendingBatchOps: () => void
  readonly sheetRecord: (sheetId: number) => SheetRecord
  readonly sheetName: (sheetId: number) => string
  readonly a1: (address: Pick<WorkPaperCellAddress, 'row' | 'col'>) => string
  readonly rangeRef: (range: WorkPaperCellRange) => CellRangeRef
  readonly listSheetRecords: () => readonly SheetRecord[]
  readonly requireSheetId: (name: string) => number
  readonly restorePublicFormula: (formula: string, ownerSheetId: number) => string
  readonly cellSnapshotToRawContent: (cell: CellSnapshot, ownerSheetId: number) => RawCellContent
}

export function createWorkPaperReadOperations(runtime: WorkPaperReadOperationsRuntime): WorkPaperReadOperations {
  const engine = (): SpreadsheetEngine => runtime.getEngine()

  const getCellValue = (address: WorkPaperCellAddress): CellValue => {
    runtime.assertReadable()
    return readWorkPaperCellValue({ engine: engine(), sheet: runtime.sheetRecord(address.sheet), address })
  }

  const getCellFormula = (address: WorkPaperCellAddress): string | undefined => {
    runtime.prepareReadableState()
    return readWorkPaperCellFormula({
      engine: engine(),
      sheetName: runtime.sheetName(address.sheet),
      a1: runtime.a1(address),
      ownerSheetId: address.sheet,
      restorePublicFormula: (formula, ownerSheetId) => runtime.restorePublicFormula(formula, ownerSheetId),
    })
  }

  const getCellSerialized = (address: WorkPaperCellAddress): RawCellContent => {
    runtime.prepareReadableState()
    return runtime.cellSnapshotToRawContent(engine().getCell(runtime.sheetName(address.sheet), runtime.a1(address)), address.sheet)
  }

  const getRangeValues = (range: WorkPaperCellRange): CellValue[][] => {
    runtime.assertReadable()
    return readWorkPaperRangeValues({ engine: engine(), range, rangeRef: () => runtime.rangeRef(range) })
  }

  const getRangeValueBlock = (range: WorkPaperCellRange): WorkPaperRangeValueBlock => {
    runtime.assertReadable()
    return readWorkPaperRangeValueBlock({ engine: engine(), range, rangeRef: () => runtime.rangeRef(range) })
  }

  const getRangeFormulas = (range: WorkPaperCellRange): WorkPaperSheetFormulas => {
    runtime.prepareReadableState()
    return buildWorkPaperDenseFormulaRange({
      engine: engine(),
      range,
      sheet: runtime.sheetRecord(range.start.sheet),
      restorePublicFormula: (formula, ownerSheetId) => runtime.restorePublicFormula(formula, ownerSheetId),
    })
  }

  const getRangeSerialized = (range: WorkPaperCellRange): RawCellContent[][] => {
    runtime.prepareReadableState()
    return buildWorkPaperDenseSerializedRange({
      engine: engine(),
      range,
      sheet: runtime.sheetRecord(range.start.sheet),
      cellSnapshotToRawContent: (cell, ownerSheetId) => runtime.cellSnapshotToRawContent(cell, ownerSheetId),
    })
  }

  const getSheetDimensions = (sheetId: number): WorkPaperSheetDimensions => {
    runtime.prepareReadableState()
    const sheet = runtime.sheetRecord(sheetId)
    const sheetDimensionCache = runtime.getSheetDimensionCache()
    const cached = sheetDimensionCache.get(sheetId)
    if (cached) {
      return { width: cached.width, height: cached.height }
    }
    const scanned = scanSheetDimensionsAndDynamicFormula(engine(), sheet)
    const dimensions = scanned.dimensions
    sheetDimensionCache.cacheScanned(sheetId, dimensions, {
      mayResizeDynamically: scanned.mayResizeDynamically,
    })
    return dimensions
  }

  const getCellValueFormat = (address: WorkPaperCellAddress): string | undefined => {
    runtime.flushPendingBatchOps()
    return engine().getCell(runtime.sheetName(address.sheet), runtime.a1(address)).format
  }

  const isCellEmpty = (address: WorkPaperCellAddress): boolean => {
    runtime.flushPendingBatchOps()
    return engine().getCellValue(runtime.sheetName(address.sheet), runtime.a1(address)).tag === ValueTag.Empty
  }

  const isCellPartOfArray = (address: WorkPaperCellAddress): boolean => {
    runtime.flushPendingBatchOps()
    return readWorkPaperCellPartOfArray({
      address,
      spillRanges: engine().getSpillRanges(),
      requireSheetId: (sheetName) => runtime.requireSheetId(sheetName),
    })
  }

  return {
    getCellValue,
    getCellFormula,
    getCellHyperlink(address) {
      return readWorkPaperCellHyperlink(getCellFormula(address))
    },
    getCellSerialized,
    getRangeValues,
    getRangeValueBlock,
    getRangeFormulas,
    getRangeSerialized,
    getSheetValues(sheetId) {
      return readWorkPaperSheetRange({
        sheetId,
        dimensions: getSheetDimensions(sheetId),
        readRange: (range) => getRangeValues(range),
      })
    },
    getSheetFormulas(sheetId) {
      return readWorkPaperSheetRange({
        sheetId,
        dimensions: getSheetDimensions(sheetId),
        readRange: (range) => getRangeFormulas(range),
      })
    },
    getSheetSerialized(sheetId) {
      return readWorkPaperSheetRange({
        sheetId,
        dimensions: getSheetDimensions(sheetId),
        readRange: (range) => getRangeSerialized(range),
      })
    },
    getAllSheetsValues() {
      runtime.assertReadable()
      return collectWorkPaperSheetsByName(runtime.listSheetRecords(), (sheet) => this.getSheetValues(sheet.id))
    },
    getAllSheetsFormulas() {
      return collectWorkPaperSheetsByName(runtime.listSheetRecords(), (sheet) => this.getSheetFormulas(sheet.id))
    },
    getAllSheetsSerialized() {
      return collectSerializedWorkPaperSheets({
        sheets: runtime.listSheetRecords(),
        readSheet: (sheet) => this.getSheetSerialized(sheet.id),
        runtimeSnapshot: engine().exportSnapshot(),
      })
    },
    getAllSheetsDimensions() {
      return collectWorkPaperSheetsByName(runtime.listSheetRecords(), (sheet) => getSheetDimensions(sheet.id))
    },
    getSheetDimensions,
    simpleCellAddressFromString(value, defaultSheetId) {
      runtime.assertNotDisposed()
      const defaultSheetName = resolveDefaultWorkPaperSheetName({
        ...(defaultSheetId !== undefined ? { defaultSheetId } : {}),
        sheets: runtime.listSheetRecords(),
        sheetName: (sheetId) => runtime.sheetName(sheetId),
      })
      return parseWorkPaperCellAddressText({
        value,
        ...(defaultSheetName !== undefined ? { defaultSheetName } : {}),
        requireSheetId: (sheetName) => runtime.requireSheetId(sheetName),
      })
    },
    simpleCellRangeFromString(value, defaultSheetId) {
      runtime.assertNotDisposed()
      const defaultSheetName = resolveDefaultWorkPaperSheetName({
        ...(defaultSheetId !== undefined ? { defaultSheetId } : {}),
        sheets: runtime.listSheetRecords(),
        sheetName: (sheetId) => runtime.sheetName(sheetId),
      })
      return parseWorkPaperCellRangeText({
        value,
        ...(defaultSheetName !== undefined ? { defaultSheetName } : {}),
        requireSheetId: (sheetName) => runtime.requireSheetId(sheetName),
      })
    },
    simpleCellAddressToString(address, optionsOrContextSheetId = {}) {
      runtime.assertNotDisposed()
      return formatWorkPaperCellAddressText({
        address,
        optionsOrContextSheetId,
        sheetName: (sheetId) => runtime.sheetName(sheetId),
      })
    },
    simpleCellRangeToString(range, optionsOrContextSheetId = {}) {
      return formatWorkPaperCellRangeText({
        range,
        optionsOrContextSheetId,
        sheetName: (sheetId) => runtime.sheetName(sheetId),
      })
    },
    getCellType(address) {
      runtime.flushPendingBatchOps()
      const cell = engine().getCell(runtime.sheetName(address.sheet), runtime.a1(address))
      return readWorkPaperCellClassification({
        hasFormula: cell.formula !== undefined,
        isEmpty: isCellEmpty(address),
        isPartOfArray: isCellPartOfArray(address),
      })
    },
    doesCellHaveSimpleValue(address) {
      runtime.flushPendingBatchOps()
      const cell = engine().getCell(runtime.sheetName(address.sheet), runtime.a1(address))
      return readWorkPaperCellHasSimpleValue({
        hasFormula: cell.formula !== undefined,
        isEmpty: isCellEmpty(address),
      })
    },
    doesCellHaveFormula(address) {
      runtime.flushPendingBatchOps()
      return engine().getCell(runtime.sheetName(address.sheet), runtime.a1(address)).formula !== undefined
    },
    isCellEmpty,
    isCellPartOfArray,
    getCellValueType(address) {
      return readWorkPaperCellValueType(getCellValue(address))
    },
    getCellValueDetailedType(address) {
      const value = getCellValue(address)
      return readWorkPaperCellValueDetailedType({
        value,
        format: value.tag === ValueTag.Number ? getCellValueFormat(address) : undefined,
      })
    },
    getCellValueFormat,
  }
}

function scanSheetDimensionsAndDynamicFormula(
  engine: SpreadsheetEngine,
  sheet: SheetRecord,
): { readonly dimensions: WorkPaperSheetDimensions; readonly mayResizeDynamically: boolean } {
  let width = 0
  let height = 0
  let mayResize = false
  const formulaIds = engine.workbook.cellStore.formulaIds
  sheet.grid.forEachCellEntry((_cellIndex, row, col) => {
    height = Math.max(height, row + 1)
    width = Math.max(width, col + 1)
    if (!mayResize && (formulaIds[_cellIndex] ?? 0) !== 0) {
      const formula = engine.getCellByIndex(_cellIndex).formula
      if (formula !== undefined && workPaperFormulaMayResizeDynamically(formula)) {
        mayResize = true
      }
    }
  })
  return { dimensions: { width, height }, mayResizeDynamically: mayResize }
}

function buildWorkPaperDenseFormulaRange(args: {
  readonly engine: SpreadsheetEngine
  readonly range: WorkPaperCellRange
  readonly sheet: SheetRecord
  readonly restorePublicFormula: (formula: string, ownerSheetId: number) => string
}): WorkPaperSheetFormulas {
  assertRange(args.range)
  const height = args.range.end.row - args.range.start.row + 1
  const width = args.range.end.col - args.range.start.col + 1
  const ownerSheetId = args.range.start.sheet
  const formulaIds = args.engine.workbook.cellStore.formulaIds
  const formulas: WorkPaperSheetFormulas = Array.from({ length: height }, () => Array<string | undefined>(width).fill(undefined))
  for (let rowOffset = 0; rowOffset < height; rowOffset += 1) {
    const outputRow = formulas[rowOffset]!
    const row = args.range.start.row + rowOffset
    for (let colOffset = 0; colOffset < width; colOffset += 1) {
      const col = args.range.start.col + colOffset
      const cellIndex = getVisibleWorkPaperCellIndexInSheet(args.sheet, row, col)
      if (cellIndex !== undefined && (formulaIds[cellIndex] ?? 0) !== 0) {
        outputRow[colOffset] = readWorkPaperCellFormulaByIndex({
          engine: args.engine,
          cellIndex,
          ownerSheetId,
          restorePublicFormula: args.restorePublicFormula,
        })
      }
    }
  }
  return formulas
}

function buildWorkPaperDenseSerializedRange(args: {
  readonly engine: SpreadsheetEngine
  readonly range: WorkPaperCellRange
  readonly sheet: SheetRecord
  readonly cellSnapshotToRawContent: (cell: CellSnapshot, ownerSheetId: number) => RawCellContent
}): RawCellContent[][] {
  assertRange(args.range)
  const height = args.range.end.row - args.range.start.row + 1
  const width = args.range.end.col - args.range.start.col + 1
  const ownerSheetId = args.range.start.sheet
  const serialized: RawCellContent[][] = Array.from({ length: height }, () => Array<RawCellContent>(width).fill(null))
  for (let rowOffset = 0; rowOffset < height; rowOffset += 1) {
    const outputRow = serialized[rowOffset]!
    const row = args.range.start.row + rowOffset
    for (let colOffset = 0; colOffset < width; colOffset += 1) {
      const col = args.range.start.col + colOffset
      const cellIndex = getVisibleWorkPaperCellIndexInSheet(args.sheet, row, col)
      if (cellIndex !== undefined) {
        outputRow[colOffset] = args.cellSnapshotToRawContent(args.engine.getCellByIndex(cellIndex), ownerSheetId)
      }
    }
  }
  return serialized
}
