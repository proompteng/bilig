import { Effect } from 'effect'
import {
  formulaLooksDateLike,
  isDateLikeHeaderValue,
  isLikelyExcelDateSerialValue,
  type CellRangeRef,
  type CellSnapshot,
  type CellValue,
  type DependencySnapshot,
  type ExplainCellSnapshot,
} from '@bilig/protocol'
import { parseCellAddress } from '@bilig/formula'
import { CellFlags } from '../../cell-store.js'
import { entityPayload, isExactLookupColumnEntity, isRangeEntity, isSortedLookupColumnEntity, makeCellEntity } from '../../entity-ids.js'
import { normalizeRange } from '../../engine-range-utils.js'
import { emptyValue } from '../../engine-value-utils.js'
import { WorkbookStore } from '../../workbook-store.js'
import type { EngineRuntimeState } from '../runtime-state.js'
import { getRuntimeFormulaSource } from '../runtime-formula-source.js'
import type { EngineRuntimeColumnStoreService } from './runtime-column-store-service.js'

export interface EngineReadService {
  readonly exportSheetCsv: (sheetName: string) => Effect.Effect<string>
  readonly getCellValue: (sheetName: string, address: string) => Effect.Effect<CellValue>
  readonly getRangeValues: (range: CellRangeRef) => Effect.Effect<CellValue[][]>
  readonly getCell: (sheetName: string, address: string) => Effect.Effect<CellSnapshot>
  readonly getCellByIndex: (cellIndex: number) => Effect.Effect<CellSnapshot>
  readonly getDependencies: (sheetName: string, address: string) => Effect.Effect<DependencySnapshot>
  readonly getDependents: (sheetName: string, address: string) => Effect.Effect<DependencySnapshot>
  readonly explainCell: (sheetName: string, address: string) => Effect.Effect<ExplainCellSnapshot>
}

export function createEngineReadService(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook' | 'strings' | 'formulas' | 'ranges'>
  readonly runtimeColumnStore: EngineRuntimeColumnStoreService
  readonly forEachFormulaDependencyCell: (cellIndex: number, fn: (dependencyCellIndex: number) => void) => void
  readonly getEntityDependents: (entityId: number) => Uint32Array
  readonly cellToCsvValue: (cell: CellSnapshot) => string
  readonly serializeCsv: (rows: string[][]) => string
}): EngineReadService {
  const inferDateFormat = (
    sheetName: string,
    row: number,
    col: number,
    value: CellValue,
    formula: string | undefined,
  ): string | undefined => {
    if (!isLikelyExcelDateSerialValue(value)) {
      return undefined
    }
    if (formulaLooksDateLike(formula)) {
      return 'date:short'
    }
    if (row === 0) {
      return undefined
    }
    const headerValue = readCellValueAt(sheetName, row - 1, col)
    return isDateLikeHeaderValue(headerValue) ? 'date:short' : undefined
  }

  const readCellValueByIndex = (cellIndex: number | undefined): CellValue => {
    if (cellIndex === undefined) {
      return emptyValue()
    }
    return args.state.workbook.cellStore.getValue(cellIndex, (id) => args.state.strings.get(id))
  }

  const readCellValueAt = (sheetName: string, row: number, col: number): CellValue => {
    const sheet = args.state.workbook.getSheet(sheetName)
    return readCellValueByIndex(sheet?.logical.getVisibleCell(row, col))
  }

  const getCellByIndex = (cellIndex: number): CellSnapshot => {
    const position = args.state.workbook.getCellPosition(cellIndex)
    const row = position?.row ?? args.state.workbook.cellStore.rows[cellIndex]!
    const col = position?.col ?? args.state.workbook.cellStore.cols[cellIndex]!
    const address = args.state.workbook.getAddress(cellIndex)
    const sheetName = args.state.workbook.getSheetNameById(args.state.workbook.cellStore.sheetIds[cellIndex]!)
    const runtimeFormula = args.state.formulas.get(cellIndex)
    const formula = runtimeFormula ? getRuntimeFormulaSource(runtimeFormula) : undefined
    const snapshot: CellSnapshot = {
      sheetName,
      address,
      value: readCellValueByIndex(cellIndex),
      flags: args.state.workbook.cellStore.flags[cellIndex]!,
      version: args.state.workbook.cellStore.versions[cellIndex] ?? 0,
    }
    const styleId = args.state.workbook.getStyleId(sheetName, row, col)
    if (styleId !== WorkbookStore.defaultStyleId) {
      snapshot.styleId = styleId
    }
    const explicitFormat = args.state.workbook.getCellFormat(cellIndex)
    const numberFormatId =
      explicitFormat !== undefined
        ? args.state.workbook.internCellNumberFormat(explicitFormat).id
        : args.state.workbook.getRangeFormatId(sheetName, row, col)
    const formatRecord = args.state.workbook.getCellNumberFormat(numberFormatId)
    if (numberFormatId !== WorkbookStore.defaultFormatId) {
      snapshot.numberFormatId = numberFormatId
    }
    if (explicitFormat !== undefined) {
      snapshot.format = explicitFormat
    } else if (formatRecord && numberFormatId !== WorkbookStore.defaultFormatId) {
      snapshot.format = formatRecord.code
    } else {
      const inferredFormat = inferDateFormat(sheetName, row, col, snapshot.value, formula)
      if (inferredFormat !== undefined) {
        snapshot.format = inferredFormat
      }
    }
    if (formula !== undefined) {
      snapshot.formula = formula
    }
    return snapshot
  }

  const getCell = (sheetName: string, address: string): CellSnapshot => {
    const cellIndex = args.state.workbook.getCellIndex(sheetName, address)
    if (cellIndex === undefined) {
      const parsed = parseCellAddress(address, sheetName)
      const styleId = args.state.workbook.getStyleId(sheetName, parsed.row, parsed.col)
      const numberFormatId = args.state.workbook.getRangeFormatId(sheetName, parsed.row, parsed.col)
      const formatRecord = args.state.workbook.getCellNumberFormat(numberFormatId)
      return {
        sheetName,
        address,
        ...(styleId !== WorkbookStore.defaultStyleId ? { styleId } : {}),
        ...(numberFormatId !== WorkbookStore.defaultFormatId ? { numberFormatId } : {}),
        ...(formatRecord && numberFormatId !== WorkbookStore.defaultFormatId ? { format: formatRecord.code } : {}),
        value: emptyValue(),
        flags: 0,
        version: 0,
      }
    }
    return getCellByIndex(cellIndex)
  }

  const getCellValue = (sheetName: string, address: string): CellValue => {
    const parsed = parseCellAddress(address, sheetName)
    return readCellValueAt(sheetName, parsed.row, parsed.col)
  }

  const readRangeValueMatrix = (range: CellRangeRef): CellValue[][] => {
    const bounds = normalizeRange(range)
    return args.runtimeColumnStore.readRangeValueMatrix({
      sheetName: range.sheetName,
      rowStart: bounds.startRow,
      rowEnd: bounds.endRow,
      colStart: bounds.startCol,
      colEnd: bounds.endCol,
    })
  }

  const getDependencies = (sheetName: string, address: string): DependencySnapshot => {
    const cellIndex = args.state.workbook.getCellIndex(sheetName, address)
    if (cellIndex === undefined) {
      return { directDependents: [], directPrecedents: [] }
    }
    const directDependents = new Set<number>()
    const directPrecedents = new Set<number>()
    args.forEachFormulaDependencyCell(cellIndex, (dependencyCellIndex) => {
      directPrecedents.add(dependencyCellIndex)
    })
    const formula = args.state.formulas.get(cellIndex)
    formula?.rangeDependencies.forEach((rangeIndex) => {
      const members = args.state.ranges.expandToCells(rangeIndex)
      for (let index = 0; index < members.length; index += 1) {
        directPrecedents.add(members[index]!)
      }
    })
    const dependents = args.getEntityDependents(makeCellEntity(cellIndex))
    for (let index = 0; index < dependents.length; index += 1) {
      const dependent = dependents[index]!
      if (isRangeEntity(dependent) || isExactLookupColumnEntity(dependent) || isSortedLookupColumnEntity(dependent)) {
        const syntheticDependents = args.getEntityDependents(dependent)
        for (let syntheticIndex = 0; syntheticIndex < syntheticDependents.length; syntheticIndex += 1) {
          const formulaEntity = syntheticDependents[syntheticIndex]!
          if (isRangeEntity(formulaEntity) || isExactLookupColumnEntity(formulaEntity) || isSortedLookupColumnEntity(formulaEntity)) {
            continue
          }
          directDependents.add(entityPayload(formulaEntity))
        }
        continue
      }
      directDependents.add(entityPayload(dependent))
    }
    return {
      directPrecedents: [...directPrecedents].map((dependencyCellIndex) => args.state.workbook.getQualifiedAddress(dependencyCellIndex)),
      directDependents: [...directDependents].map((dependentCellIndex) => args.state.workbook.getQualifiedAddress(dependentCellIndex)),
    }
  }

  const explainCell = (sheetName: string, address: string): ExplainCellSnapshot => {
    const cellIndex = args.state.workbook.getCellIndex(sheetName, address)
    if (cellIndex === undefined) {
      return {
        sheetName,
        address,
        value: emptyValue(),
        flags: 0,
        version: 0,
        inCycle: false,
        directPrecedents: [],
        directDependents: [],
      }
    }

    const snapshot = getCellByIndex(cellIndex)
    const formula = args.state.formulas.get(cellIndex)
    const flags = args.state.workbook.cellStore.flags[cellIndex] ?? 0
    const isFormula = (flags & CellFlags.HasFormula) !== 0 && formula !== undefined
    const dependencies = getDependencies(sheetName, address)

    const explanation: ExplainCellSnapshot = {
      ...snapshot,
      version: args.state.workbook.cellStore.versions[cellIndex] ?? 0,
      inCycle: (flags & CellFlags.InCycle) !== 0,
      directPrecedents: dependencies.directPrecedents,
      directDependents: dependencies.directDependents,
    }

    if (formula?.source !== undefined) {
      explanation.formula = formula.source
    }
    if (isFormula) {
      explanation.mode = formula.compiled.mode
      explanation.topoRank = args.state.workbook.cellStore.topoRanks[cellIndex] ?? 0
    }

    return explanation
  }

  const exportSheetCsv = (sheetName: string): string => {
    const sheet = args.state.workbook.getSheet(sheetName)
    if (!sheet) {
      return ''
    }

    let maxRow = -1
    let maxCol = -1
    const cells = new Map<string, string>()

    sheet.grid.forEachCell((cellIndex) => {
      const cell = getCellByIndex(cellIndex)
      const parsed = parseCellAddress(cell.address, sheetName)
      maxRow = Math.max(maxRow, parsed.row)
      maxCol = Math.max(maxCol, parsed.col)
      cells.set(`${parsed.row}:${parsed.col}`, args.cellToCsvValue(cell))
    })

    if (maxRow < 0 || maxCol < 0) {
      return ''
    }

    const rows = Array.from({ length: maxRow + 1 }, (_rowEntry, row) =>
      Array.from({ length: maxCol + 1 }, (_colEntry, col) => cells.get(`${row}:${col}`) ?? ''),
    )

    return args.serializeCsv(rows)
  }

  return {
    exportSheetCsv(sheetName) {
      return Effect.sync(() => exportSheetCsv(sheetName))
    },
    getCellValue(sheetName, address) {
      return Effect.sync(() => getCellValue(sheetName, address))
    },
    getRangeValues(range) {
      return Effect.sync(() => readRangeValueMatrix(range))
    },
    getCell(sheetName, address) {
      return Effect.sync(() => getCell(sheetName, address))
    },
    getCellByIndex(cellIndex) {
      return Effect.sync(() => getCellByIndex(cellIndex))
    },
    getDependencies(sheetName, address) {
      return Effect.sync(() => getDependencies(sheetName, address))
    },
    getDependents(sheetName, address) {
      return Effect.sync(() => getDependencies(sheetName, address))
    },
    explainCell(sheetName, address) {
      return Effect.sync(() => explainCell(sheetName, address))
    },
  }
}
