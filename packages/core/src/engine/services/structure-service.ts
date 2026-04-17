import { Effect } from 'effect'
import type { SheetFormatRangeSnapshot, SheetStyleRangeSnapshot } from '@bilig/protocol'
import {
  formatAddress,
  rewriteAddressForStructuralTransform,
  rewriteCompiledFormulaForStructuralTransform,
  rewriteFormulaForStructuralTransform,
  rewriteRangeForStructuralTransform,
  type StructuralAxisTransform,
} from '@bilig/formula'
import type { CompiledFormula } from '@bilig/formula'
import type { EngineOp } from '@bilig/workbook-domain'
import { CellFlags } from '../../cell-store.js'
import { emptyValue } from '../../engine-value-utils.js'
import { ValueTag } from '@bilig/protocol'
import { mapStructuralAxisIndex, mapStructuralBoundary, structuralTransformForOp } from '../../engine-structural-utils.js'
import type { FormulaTable } from '../../formula-table.js'
import type { RuntimeFormula } from '../runtime-state.js'
import { EngineStructureError } from '../errors.js'
import { makeCellKey, normalizeDefinedName, type WorkbookPivotRecord, type WorkbookStore } from '../../workbook-store.js'
import type { RangeRegistry } from '../../range-registry.js'

type StructuralAxisOp = Extract<
  EngineOp,
  {
    kind: 'insertRows' | 'deleteRows' | 'moveRows' | 'insertColumns' | 'deleteColumns' | 'moveColumns'
  }
>

interface EngineStructureState {
  readonly workbook: WorkbookStore
  readonly formulas: FormulaTable<RuntimeFormula>
  readonly ranges: RangeRegistry
  readonly pivotOutputOwners: Map<number, string>
}

interface StructuralFormulaRebindInput {
  readonly cellIndex: number
  readonly ownerSheetName: string
  readonly source: string
  readonly compiled?: CompiledFormula
  readonly preservesValue?: boolean
}

function dependencyTouchesSheet(dependency: string, sheetName: string): boolean {
  if (!dependency.includes('!')) {
    return false
  }
  const [qualifiedSheetName] = dependency.split('!')
  return qualifiedSheetName?.replace(/^'(.*)'$/, '$1') === sheetName
}

function rangeDependencyAxisAffected(
  rangeDescriptor: { sheetId: number; row1: number; row2: number; col1: number; col2: number },
  targetSheetId: number,
  transform: StructuralAxisTransform,
): boolean {
  if (rangeDescriptor.sheetId !== targetSheetId) {
    return false
  }
  const start = transform.axis === 'row' ? rangeDescriptor.row1 : rangeDescriptor.col1
  const end = transform.axis === 'row' ? rangeDescriptor.row2 : rangeDescriptor.col2
  return !(end < transform.start || start >= transform.start + transform.count)
}

function runtimeDirectRangeAxisAffected(
  targetSheetId: number | undefined,
  targetSheetName: string,
  transform: StructuralAxisTransform,
  range: { sheetName: string; rowStart: number; rowEnd: number; col: number } | undefined,
): boolean {
  if (!range || targetSheetId === undefined || range.sheetName !== targetSheetName) {
    return false
  }
  const descriptor =
    transform.axis === 'row'
      ? {
          sheetId: targetSheetId,
          row1: range.rowStart,
          row2: range.rowEnd,
          col1: range.col,
          col2: range.col,
        }
      : {
          sheetId: targetSheetId,
          row1: range.rowStart,
          row2: range.rowEnd,
          col1: range.col,
          col2: range.col,
        }
  return rangeDependencyAxisAffected(descriptor, targetSheetId, transform)
}

function isStructurallyStableSimpleFormulaNode(node: CompiledFormula['optimizedAst']): boolean {
  switch (node.kind) {
    case 'NumberLiteral':
    case 'BooleanLiteral':
    case 'StringLiteral':
    case 'ErrorLiteral':
    case 'CellRef':
      return true
    case 'UnaryExpr':
      return isStructurallyStableSimpleFormulaNode(node.argument)
    case 'BinaryExpr':
      return isStructurallyStableSimpleFormulaNode(node.left) && isStructurallyStableSimpleFormulaNode(node.right)
    case 'NameRef':
    case 'StructuredRef':
    case 'SpillRef':
    case 'RowRef':
    case 'ColumnRef':
    case 'RangeRef':
    case 'CallExpr':
    case 'InvokeExpr':
      return false
  }
}

function structuralRewritePreservesValue(
  formula: RuntimeFormula,
  rewritten: { compiled: CompiledFormula; reusedProgram: boolean },
  transform: StructuralAxisTransform,
): boolean {
  return (
    transform.kind !== 'delete' &&
    rewritten.reusedProgram &&
    !rewritten.compiled.volatile &&
    formula.compiled.symbolicNames.length === 0 &&
    formula.compiled.symbolicTables.length === 0 &&
    formula.compiled.symbolicSpills.length === 0 &&
    formula.directLookup === undefined &&
    formula.directAggregate === undefined &&
    formula.directCriteria === undefined &&
    isStructurallyStableSimpleFormulaNode(rewritten.compiled.optimizedAst)
  )
}

function structuralRemapScope(transform: StructuralAxisTransform): { start: number; end?: number } {
  switch (transform.kind) {
    case "insert":
    case "delete":
      return { start: transform.start };
    case "move":
      if (transform.target < transform.start) {
        return { start: transform.target, end: transform.start + transform.count };
      }
      if (transform.target > transform.start) {
        return { start: transform.start, end: transform.target + transform.count };
      }
      return { start: transform.start, end: transform.start };
    default: {
      const exhaustive: never = transform;
      return exhaustive;
    }
  }
}

export interface EngineStructureService {
  readonly captureSheetCellState: (sheetName: string) => Effect.Effect<EngineOp[], EngineStructureError>
  readonly captureRowRangeCellState: (sheetName: string, start: number, count: number) => Effect.Effect<EngineOp[], EngineStructureError>
  readonly captureColumnRangeCellState: (sheetName: string, start: number, count: number) => Effect.Effect<EngineOp[], EngineStructureError>
  readonly applyStructuralAxisOp: (op: StructuralAxisOp) => Effect.Effect<
    {
      changedCellIndices: number[]
      formulaCellIndices: number[]
    },
    EngineStructureError
  >
}

export function createEngineStructureService(args: {
  readonly state: EngineStructureState
  readonly captureStoredCellOps: (
    cellIndex: number,
    sheetName: string,
    address: string,
    sourceSheetName?: string,
    sourceAddress?: string,
  ) => EngineOp[]
  readonly removeFormula: (cellIndex: number) => boolean
  readonly clearOwnedPivot: (pivot: WorkbookPivotRecord) => number[]
  readonly refreshRangeDependencies: (rangeIndices: readonly number[]) => void
  readonly rebindFormulaCells: (inputs: readonly StructuralFormulaRebindInput[]) => void
}): EngineStructureService {
  const shouldCaptureStoredCell = (cellIndex: number): boolean => {
    const value = args.state.workbook.cellStore.getValue(cellIndex, () => '')
    const explicitFormat = args.state.workbook.getCellFormat(cellIndex)
    const flags = args.state.workbook.cellStore.flags[cellIndex] ?? 0
    const formula = args.state.formulas.get(cellIndex)
    if ((flags & (CellFlags.SpillChild | CellFlags.PivotOutput)) !== 0) {
      return false
    }
    return !(
      formula === undefined &&
      explicitFormat === undefined &&
      (flags & CellFlags.AuthoredBlank) === 0 &&
      (value.tag === ValueTag.Empty || value.tag === ValueTag.Error)
    )
  }

  const captureStoredCellState = (
    cellIndex: number,
    sheetName: string,
    address: string,
    sourceSheetName?: string,
    sourceAddress?: string,
  ): EngineOp[] => args.captureStoredCellOps(cellIndex, sheetName, address, sourceSheetName, sourceAddress)

  const captureAxisRangeCellState = (sheetName: string, axis: 'row' | 'column', start: number, count: number): EngineOp[] => {
    const sheet = args.state.workbook.getSheet(sheetName)
    if (!sheet) {
      return []
    }
    const captured: Array<{ cellIndex: number; row: number; col: number }> = []
    sheet.grid.forEachCellEntry((cellIndex, row, col) => {
      if (!shouldCaptureStoredCell(cellIndex)) {
        return
      }
      const index = axis === 'row' ? row : col
      if (index >= start && index < start + count) {
        captured.push({ cellIndex, row, col })
      }
    })
    return captured
      .toSorted((left, right) => left.row - right.row || left.col - right.col)
      .flatMap(({ cellIndex, row, col }) => captureStoredCellState(cellIndex, sheetName, formatAddress(row, col)))
  }

  const captureSheetCellState = (sheetName: string): EngineOp[] => {
    const sheet = args.state.workbook.getSheet(sheetName)
    if (!sheet) {
      return []
    }
    const captured: Array<{ cellIndex: number; row: number; col: number }> = []
    sheet.grid.forEachCellEntry((cellIndex, row, col) => {
      if (!shouldCaptureStoredCell(cellIndex)) {
        return
      }
      captured.push({ cellIndex, row, col })
    })
    return captured
      .toSorted((left, right) => left.row - right.row || left.col - right.col)
      .flatMap(({ cellIndex }) => captureStoredCellState(cellIndex, sheetName, args.state.workbook.getAddress(cellIndex)))
  }

  const rewriteDefinedNamesForStructuralTransform = (sheetName: string, transform: StructuralAxisTransform): Set<string> => {
    const changedNames = new Set<string>()
    args.state.workbook.listDefinedNames().forEach((record) => {
      if (typeof record.value === 'string' && record.value.startsWith('=')) {
        const nextFormula = rewriteFormulaForStructuralTransform(record.value.slice(1), sheetName, sheetName, transform)
        if (`=${nextFormula}` !== record.value) {
          args.state.workbook.setDefinedName(record.name, `=${nextFormula}`)
        }
        return
      }
      if (typeof record.value !== 'object' || !record.value) {
        return
      }
      switch (record.value.kind) {
        case 'formula': {
          const nextFormula = rewriteFormulaForStructuralTransform(
            record.value.formula.startsWith('=') ? record.value.formula.slice(1) : record.value.formula,
            sheetName,
            sheetName,
            transform,
          )
          const nextValue = {
            ...record.value,
            formula: record.value.formula.startsWith('=') ? `=${nextFormula}` : nextFormula,
          }
          if (nextValue.formula !== record.value.formula) {
            args.state.workbook.setDefinedName(record.name, nextValue)
            changedNames.add(normalizeDefinedName(record.name))
          }
          return
        }
        case 'cell-ref': {
          if (record.value.sheetName !== sheetName) {
            return
          }
          const nextAddress = rewriteAddressForStructuralTransform(record.value.address, transform)
          if (!nextAddress) {
            args.state.workbook.deleteDefinedName(record.name)
            changedNames.add(normalizeDefinedName(record.name))
            return
          }
          if (nextAddress !== record.value.address) {
            args.state.workbook.setDefinedName(record.name, {
              ...record.value,
              address: nextAddress,
            })
            changedNames.add(normalizeDefinedName(record.name))
          }
          return
        }
        case 'range-ref': {
          if (record.value.sheetName !== sheetName) {
            return
          }
          const nextRange = rewriteRangeForStructuralTransform(record.value.startAddress, record.value.endAddress, transform)
          if (!nextRange) {
            args.state.workbook.deleteDefinedName(record.name)
            changedNames.add(normalizeDefinedName(record.name))
            return
          }
          if (nextRange.startAddress !== record.value.startAddress || nextRange.endAddress !== record.value.endAddress) {
            args.state.workbook.setDefinedName(record.name, {
              ...record.value,
              startAddress: nextRange.startAddress,
              endAddress: nextRange.endAddress,
            })
            changedNames.add(normalizeDefinedName(record.name))
          }
          return
        }
        case 'scalar':
        case 'structured-ref':
          return
      }
    })
    return changedNames
  }

  const rewriteStructuralFormulaCompiled = (
    formula: RuntimeFormula,
    ownerSheetName: string,
    sheetName: string,
    transform: StructuralAxisTransform,
  ): ReturnType<typeof rewriteCompiledFormulaForStructuralTransform> | undefined => {
    const rewritten = rewriteCompiledFormulaForStructuralTransform(formula.compiled, ownerSheetName, sheetName, transform)
    return rewritten.source === formula.source ? undefined : rewritten
  }

  const rewriteFormulaSourceFallback = (
    source: string,
    ownerSheetName: string,
    sheetName: string,
    transform: StructuralAxisTransform,
  ): string => rewriteFormulaForStructuralTransform(source, ownerSheetName, sheetName, transform)

  const resolveStructuralFormulaRebindInputs = (argsForResolve: {
    readonly formulaCellIndices: readonly number[]
    readonly sheetName: string
    readonly transform: StructuralAxisTransform
    readonly changedDefinedNames: ReadonlySet<string>
    readonly changedTableNames: ReadonlySet<string>
  }): StructuralFormulaRebindInput[] => {
    const inputs: StructuralFormulaRebindInput[] = []
    argsForResolve.formulaCellIndices.forEach((cellIndex) => {
      const formula = args.state.formulas.get(cellIndex)
      if (!formula) {
        return
      }
      const ownerSheetName = args.state.workbook.getSheetNameById(args.state.workbook.cellStore.sheetIds[cellIndex]!)
      if (!ownerSheetName) {
        return
      }
      const rewritten = rewriteStructuralFormulaCompiled(formula, ownerSheetName, argsForResolve.sheetName, argsForResolve.transform)
      const touchesChangedName = formula.compiled.symbolicNames.some((name) =>
        argsForResolve.changedDefinedNames.has(normalizeDefinedName(name)),
      )
      const touchesChangedTable = formula.compiled.symbolicTables.some((name) => argsForResolve.changedTableNames.has(name))
      if (!rewritten) {
        inputs.push({
          cellIndex,
          ownerSheetName,
          source: formula.source,
        })
        return
      }
      if (touchesChangedName || touchesChangedTable) {
        inputs.push({
          cellIndex,
          ownerSheetName,
          source: rewriteFormulaSourceFallback(formula.source, ownerSheetName, argsForResolve.sheetName, argsForResolve.transform),
        })
        return
      }
      inputs.push({
        cellIndex,
        ownerSheetName,
        source: rewritten.source,
        compiled: rewritten.compiled,
        preservesValue: structuralRewritePreservesValue(formula, rewritten, argsForResolve.transform),
      })
    })
    return inputs
  }

  const collectStructuralRangeDependencies = (argsForCollect: { readonly formulaCellIndices: readonly number[] }): number[] => {
    const rangeIndices = new Set<number>()
    argsForCollect.formulaCellIndices.forEach((cellIndex) => {
      const formula = args.state.formulas.get(cellIndex)
      if (!formula) {
        return
      }
      formula.rangeDependencies.forEach((rangeIndex) => {
        rangeIndices.add(rangeIndex)
      })
    })
    return [...rangeIndices]
  }

  const clearSpillMetadataForSheet = (sheetName: string): void => {
    args.state.workbook.listSpills().forEach((spill) => {
      if (spill.sheetName !== sheetName) {
        return
      }
      args.state.workbook.deleteSpill(spill.sheetName, spill.address)
    })
  }

  const clearPivotOutputsForSheet = (sheetName: string): void => {
    args.state.workbook
      .listPivots()
      .filter((pivot) => pivot.sheetName === sheetName)
      .forEach((pivot) => {
        args.clearOwnedPivot(pivot)
      })
  }

  const clearDerivedCellArtifacts = (cellIndex: number): void => {
    args.state.pivotOutputOwners.delete(cellIndex)
  }

  const rewriteWorkbookMetadataForStructuralTransform = (
    sheetName: string,
    transform: StructuralAxisTransform,
  ): { changedTableNames: Set<string> } => {
    const changedTableNames = new Set<string>()
    args.state.workbook
      .listTables()
      .filter((table) => table.sheetName === sheetName)
      .forEach((table) => {
        const range = rewriteRangeForStructuralTransform(table.startAddress, table.endAddress, transform)
        if (!range) {
          changedTableNames.add(table.name)
          args.state.workbook.deleteTable(table.name)
          return
        }
        changedTableNames.add(table.name)
        args.state.workbook.setTable({
          ...table,
          startAddress: range.startAddress,
          endAddress: range.endAddress,
        })
      })
    args.state.workbook.listFilters(sheetName).forEach((filter) => {
      const range = rewriteRangeForStructuralTransform(filter.range.startAddress, filter.range.endAddress, transform)
      args.state.workbook.deleteFilter(sheetName, filter.range)
      if (range) {
        args.state.workbook.setFilter(sheetName, {
          ...filter.range,
          startAddress: range.startAddress,
          endAddress: range.endAddress,
        })
      }
    })
    args.state.workbook.listSorts(sheetName).forEach((sort) => {
      const range = rewriteRangeForStructuralTransform(sort.range.startAddress, sort.range.endAddress, transform)
      args.state.workbook.deleteSort(sheetName, sort.range)
      if (!range) {
        return
      }
      args.state.workbook.setSort(
        sheetName,
        { ...sort.range, startAddress: range.startAddress, endAddress: range.endAddress },
        sort.keys.map((key) => ({
          ...key,
          keyAddress: rewriteAddressForStructuralTransform(key.keyAddress, transform) ?? key.keyAddress,
        })),
      )
    })
    args.state.workbook.listDataValidations(sheetName).forEach((validation) => {
      const range = rewriteRangeForStructuralTransform(validation.range.startAddress, validation.range.endAddress, transform)
      args.state.workbook.deleteDataValidation(sheetName, validation.range)
      if (!range) {
        return
      }
      const nextValidation = structuredClone(validation)
      nextValidation.range = {
        ...validation.range,
        startAddress: range.startAddress,
        endAddress: range.endAddress,
      }
      if (nextValidation.rule.kind === 'list' && nextValidation.rule.source) {
        switch (nextValidation.rule.source.kind) {
          case 'cell-ref': {
            if (nextValidation.rule.source.sheetName !== sheetName) {
              break
            }
            const nextAddress = rewriteAddressForStructuralTransform(nextValidation.rule.source.address, transform)
            if (!nextAddress) {
              return
            }
            nextValidation.rule.source.address = nextAddress
            break
          }
          case 'range-ref': {
            if (nextValidation.rule.source.sheetName !== sheetName) {
              break
            }
            const nextSourceRange = rewriteRangeForStructuralTransform(
              nextValidation.rule.source.startAddress,
              nextValidation.rule.source.endAddress,
              transform,
            )
            if (!nextSourceRange) {
              return
            }
            nextValidation.rule.source.startAddress = nextSourceRange.startAddress
            nextValidation.rule.source.endAddress = nextSourceRange.endAddress
            break
          }
          case 'named-range':
          case 'structured-ref':
            break
        }
      }
      args.state.workbook.setDataValidation(nextValidation)
    })
    args.state.workbook.listConditionalFormats(sheetName).forEach((format) => {
      const range = rewriteRangeForStructuralTransform(format.range.startAddress, format.range.endAddress, transform)
      args.state.workbook.deleteConditionalFormat(format.id)
      if (!range) {
        return
      }
      args.state.workbook.setConditionalFormat({
        ...format,
        range: {
          ...format.range,
          startAddress: range.startAddress,
          endAddress: range.endAddress,
        },
      })
    })
    args.state.workbook.listRangeProtections(sheetName).forEach((protection) => {
      const range = rewriteRangeForStructuralTransform(protection.range.startAddress, protection.range.endAddress, transform)
      args.state.workbook.deleteRangeProtection(protection.id)
      if (!range) {
        return
      }
      args.state.workbook.setRangeProtection({
        ...protection,
        range: {
          ...protection.range,
          startAddress: range.startAddress,
          endAddress: range.endAddress,
        },
      })
    })
    args.state.workbook.listCommentThreads(sheetName).forEach((thread) => {
      const nextAddress = rewriteAddressForStructuralTransform(thread.address, transform)
      args.state.workbook.deleteCommentThread(sheetName, thread.address)
      if (!nextAddress) {
        return
      }
      args.state.workbook.setCommentThread({
        ...thread,
        address: nextAddress,
      })
    })
    args.state.workbook.listNotes(sheetName).forEach((note) => {
      const nextAddress = rewriteAddressForStructuralTransform(note.address, transform)
      args.state.workbook.deleteNote(sheetName, note.address)
      if (!nextAddress) {
        return
      }
      args.state.workbook.setNote({
        ...note,
        address: nextAddress,
      })
    })
    const rewrittenStyleRanges: SheetStyleRangeSnapshot[] = []
    const rewrittenFormatRanges: SheetFormatRangeSnapshot[] = []
    args.state.workbook.listStyleRanges(sheetName).forEach((record) => {
      const range = rewriteRangeForStructuralTransform(record.range.startAddress, record.range.endAddress, transform)
      if (!range) {
        return
      }
      rewrittenStyleRanges.push({
        range: {
          ...record.range,
          startAddress: range.startAddress,
          endAddress: range.endAddress,
        },
        styleId: record.styleId,
      })
    })
    args.state.workbook.setStyleRanges(sheetName, rewrittenStyleRanges)
    args.state.workbook.listFormatRanges(sheetName).forEach((record) => {
      const range = rewriteRangeForStructuralTransform(record.range.startAddress, record.range.endAddress, transform)
      if (!range) {
        return
      }
      rewrittenFormatRanges.push({
        range: {
          ...record.range,
          startAddress: range.startAddress,
          endAddress: range.endAddress,
        },
        formatId: record.formatId,
      })
    })
    args.state.workbook.setFormatRanges(sheetName, rewrittenFormatRanges)
    const freezePane = args.state.workbook.getFreezePane(sheetName)
    if (freezePane) {
      const nextRows = transform.axis === 'row' ? mapStructuralBoundary(freezePane.rows, transform) : freezePane.rows
      const nextCols = transform.axis === 'column' ? mapStructuralBoundary(freezePane.cols, transform) : freezePane.cols
      if (nextRows <= 0 && nextCols <= 0) {
        args.state.workbook.clearFreezePane(sheetName)
      } else {
        args.state.workbook.setFreezePane(sheetName, nextRows, nextCols)
      }
    }
    args.state.workbook.listPivots().forEach((pivot) => {
      const nextAddress = pivot.sheetName === sheetName ? rewriteAddressForStructuralTransform(pivot.address, transform) : pivot.address
      const nextSource =
        pivot.source.sheetName === sheetName
          ? rewriteRangeForStructuralTransform(pivot.source.startAddress, pivot.source.endAddress, transform)
          : { startAddress: pivot.source.startAddress, endAddress: pivot.source.endAddress }
      if (!nextAddress || !nextSource) {
        args.clearOwnedPivot(pivot)
        args.state.workbook.deletePivot(pivot.sheetName, pivot.address)
        return
      }
      if (nextAddress !== pivot.address) {
        args.clearOwnedPivot(pivot)
        args.state.workbook.deletePivot(pivot.sheetName, pivot.address)
      }
      args.state.workbook.setPivot({
        ...pivot,
        address: nextAddress,
        source: {
          ...pivot.source,
          startAddress: nextSource.startAddress,
          endAddress: nextSource.endAddress,
        },
      })
    })
    args.state.workbook.listCharts().forEach((chart) => {
      const nextAddress = chart.sheetName === sheetName ? rewriteAddressForStructuralTransform(chart.address, transform) : chart.address
      const nextSource =
        chart.source.sheetName === sheetName
          ? rewriteRangeForStructuralTransform(chart.source.startAddress, chart.source.endAddress, transform)
          : { startAddress: chart.source.startAddress, endAddress: chart.source.endAddress }
      if (!nextAddress || !nextSource) {
        args.state.workbook.deleteChart(chart.id)
        return
      }
      args.state.workbook.setChart({
        ...chart,
        address: nextAddress,
        source: {
          ...chart.source,
          startAddress: nextSource.startAddress,
          endAddress: nextSource.endAddress,
        },
      })
    })
    args.state.workbook.listImages().forEach((image) => {
      if (image.sheetName !== sheetName) {
        return
      }
      const nextAddress = rewriteAddressForStructuralTransform(image.address, transform)
      if (!nextAddress) {
        args.state.workbook.deleteImage(image.id)
        return
      }
      args.state.workbook.setImage({
        ...image,
        address: nextAddress,
      })
    })
    args.state.workbook.listShapes().forEach((shape) => {
      if (shape.sheetName !== sheetName) {
        return
      }
      const nextAddress = rewriteAddressForStructuralTransform(shape.address, transform)
      if (!nextAddress) {
        args.state.workbook.deleteShape(shape.id)
        return
      }
      args.state.workbook.setShape({
        ...shape,
        address: nextAddress,
      })
    })
    return { changedTableNames }
  }

  const isCellIndexMapped = (cellIndex: number): boolean => {
    const sheetId = args.state.workbook.cellStore.sheetIds[cellIndex]
    const row = args.state.workbook.cellStore.rows[cellIndex]
    const col = args.state.workbook.cellStore.cols[cellIndex]
    if (sheetId === undefined || row === undefined || col === undefined || !Number.isFinite(row) || !Number.isFinite(col)) {
      return false
    }
    return args.state.workbook.cellKeyToIndex.get(makeCellKey(sheetId, row, col)) === cellIndex
  }

  const structuralAxisIndexAffected = (axisIndex: number, transform: StructuralAxisTransform): boolean => {
    const nextIndex = mapStructuralAxisIndex(axisIndex, transform)
    return nextIndex === undefined || nextIndex !== axisIndex
  }

  const collectStructuralFormulaImpacts = (argsForImpact: {
    readonly targetSheetId: number | undefined
    readonly transform: StructuralAxisTransform
    readonly sheetName: string
    readonly changedDefinedNames: ReadonlySet<string>
    readonly changedTableNames: ReadonlySet<string>
  }): {
    formulaCellIndices: number[]
    rebindCellIndices: number[]
  } => {
    const formulaCellIndices = new Set<number>()
    const rebindCellIndices = new Set<number>()
    args.state.formulas.forEach((formula, cellIndex) => {
      if (!isCellIndexMapped(cellIndex)) {
        return
      }
      const ownerSheetName = args.state.workbook.getSheetNameById(args.state.workbook.cellStore.sheetIds[cellIndex]!)
      if (!ownerSheetName) {
        return
      }
      const axisIndex =
        argsForImpact.transform.axis === 'row'
          ? args.state.workbook.cellStore.rows[cellIndex]
          : args.state.workbook.cellStore.cols[cellIndex]
      const ownerPositionAffected =
        ownerSheetName === argsForImpact.sheetName &&
        axisIndex !== undefined &&
        structuralAxisIndexAffected(axisIndex, argsForImpact.transform)
      const dependencyPositionAffected =
        argsForImpact.targetSheetId !== undefined &&
        (formula.dependencyIndices.some((dependencyCellIndex) => {
          if (args.state.workbook.cellStore.sheetIds[dependencyCellIndex] !== argsForImpact.targetSheetId) {
            return false
          }
          const dependencyAxisIndex =
            argsForImpact.transform.axis === 'row'
              ? args.state.workbook.cellStore.rows[dependencyCellIndex]
              : args.state.workbook.cellStore.cols[dependencyCellIndex]
          return dependencyAxisIndex !== undefined && structuralAxisIndexAffected(dependencyAxisIndex, argsForImpact.transform)
        }) ||
          formula.rangeDependencies.some((rangeIndex) =>
            rangeDependencyAxisAffected(args.state.ranges.getDescriptor(rangeIndex), argsForImpact.targetSheetId!, argsForImpact.transform),
          ) ||
          runtimeDirectRangeAxisAffected(
            argsForImpact.targetSheetId,
            argsForImpact.sheetName,
            argsForImpact.transform,
            formula.directAggregate,
          ) ||
          runtimeDirectRangeAxisAffected(
            argsForImpact.targetSheetId,
            argsForImpact.sheetName,
            argsForImpact.transform,
            formula.directCriteria?.aggregateRange,
          ) ||
          formula.directCriteria?.criteriaPairs.some((pair) =>
            runtimeDirectRangeAxisAffected(argsForImpact.targetSheetId, argsForImpact.sheetName, argsForImpact.transform, pair.range),
          ) ||
          runtimeDirectRangeAxisAffected(
            argsForImpact.targetSheetId,
            argsForImpact.sheetName,
            argsForImpact.transform,
            formula.directLookup?.kind === 'exact' || formula.directLookup?.kind === 'approximate'
              ? {
                  sheetName: formula.directLookup.prepared.sheetName,
                  rowStart: formula.directLookup.prepared.rowStart,
                  rowEnd: formula.directLookup.prepared.rowEnd,
                  col: formula.directLookup.prepared.col,
                }
              : formula.directLookup?.kind === 'exact-uniform-numeric' || formula.directLookup?.kind === 'approximate-uniform-numeric'
                ? {
                    sheetName: formula.directLookup.sheetName,
                    rowStart: formula.directLookup.rowStart,
                    rowEnd: formula.directLookup.rowEnd,
                    col: formula.directLookup.col,
                  }
                : undefined,
          ))
      const touchesSheetDependency = formula.compiled.deps.some((dependency) => dependencyTouchesSheet(dependency, argsForImpact.sheetName))
      const touchesChangedName = formula.compiled.symbolicNames.some((name) =>
        argsForImpact.changedDefinedNames.has(normalizeDefinedName(name)),
      )
      const touchesChangedTable = formula.compiled.symbolicTables.some((name) => argsForImpact.changedTableNames.has(name))
      if (!ownerPositionAffected && !dependencyPositionAffected && !touchesSheetDependency && !touchesChangedName && !touchesChangedTable) {
        return
      }
      formulaCellIndices.add(cellIndex)
      if (ownerPositionAffected || dependencyPositionAffected || touchesSheetDependency || touchesChangedName || touchesChangedTable) {
        rebindCellIndices.add(cellIndex)
      }
    })
    return {
      formulaCellIndices: [...formulaCellIndices],
      rebindCellIndices: [...rebindCellIndices],
    }
  }

  return {
    captureSheetCellState(sheetName) {
      return Effect.try({
        try: () => captureSheetCellState(sheetName),
        catch: (cause) =>
          new EngineStructureError({
            message: `Failed to capture sheet cell state for ${sheetName}`,
            cause,
          }),
      })
    },
    captureRowRangeCellState(sheetName, start, count) {
      return Effect.try({
        try: () => captureAxisRangeCellState(sheetName, 'row', start, count),
        catch: (cause) =>
          new EngineStructureError({
            message: `Failed to capture row state for ${sheetName}`,
            cause,
          }),
      })
    },
    captureColumnRangeCellState(sheetName, start, count) {
      return Effect.try({
        try: () => captureAxisRangeCellState(sheetName, 'column', start, count),
        catch: (cause) =>
          new EngineStructureError({
            message: `Failed to capture column state for ${sheetName}`,
            cause,
          }),
      })
    },
    applyStructuralAxisOp(op) {
      return Effect.try({
        try: () => {
          const axis = op.kind.includes('Rows') ? 'row' : 'column'
          const transform = structuralTransformForOp(op)
          const sheetName = op.sheetName
          const targetSheetId = args.state.workbook.getSheet(sheetName)?.id

          clearPivotOutputsForSheet(sheetName)
          const changedDefinedNames = rewriteDefinedNamesForStructuralTransform(sheetName, transform)
          const { changedTableNames } = rewriteWorkbookMetadataForStructuralTransform(sheetName, transform)
          const impactedFormulas = collectStructuralFormulaImpacts({
            targetSheetId,
            transform,
            sheetName,
            changedDefinedNames,
            changedTableNames,
          })

          switch (op.kind) {
            case 'insertRows':
              args.state.workbook.insertRows(sheetName, op.start, op.count, op.entries)
              break
            case 'deleteRows':
              args.state.workbook.deleteRows(sheetName, op.start, op.count)
              break
            case 'moveRows':
              args.state.workbook.moveRows(sheetName, op.start, op.count, op.target)
              break
            case 'insertColumns':
              args.state.workbook.insertColumns(sheetName, op.start, op.count, op.entries)
              break
            case 'deleteColumns':
              args.state.workbook.deleteColumns(sheetName, op.start, op.count)
              break
            case 'moveColumns':
              args.state.workbook.moveColumns(sheetName, op.start, op.count, op.target)
              break
          }

          const structuralRangeDependencies = collectStructuralRangeDependencies({
            formulaCellIndices: impactedFormulas.formulaCellIndices,
          })

          const remapped = args.state.workbook.remapSheetCells(
            sheetName,
            axis,
            (index) => mapStructuralAxisIndex(index, transform),
            structuralRemapScope(transform),
          );
          remapped.removedCellIndices.forEach((cellIndex) => {
            clearDerivedCellArtifacts(cellIndex)
            args.removeFormula(cellIndex)
            args.state.workbook.setCellFormat(cellIndex, null)
            args.state.workbook.cellStore.setValue(cellIndex, emptyValue())
            args.state.workbook.cellStore.flags[cellIndex] =
              (args.state.workbook.cellStore.flags[cellIndex] ?? 0) &
              ~(CellFlags.HasFormula | CellFlags.JsOnly | CellFlags.InCycle | CellFlags.SpillChild | CellFlags.PivotOutput)
            args.state.workbook.detachCellIndex(cellIndex)
          })

          clearSpillMetadataForSheet(sheetName)
          const formulaCellIndices = impactedFormulas.formulaCellIndices.filter((cellIndex) => isCellIndexMapped(cellIndex))
          args.refreshRangeDependencies(structuralRangeDependencies)
          const rebindInputs = resolveStructuralFormulaRebindInputs({
            formulaCellIndices: impactedFormulas.rebindCellIndices.filter((cellIndex) => isCellIndexMapped(cellIndex)),
            sheetName,
            transform,
            changedDefinedNames,
            changedTableNames,
          })
          args.rebindFormulaCells(rebindInputs)
          const preservedFormulaCellIndices = new Set(rebindInputs.filter((input) => input.preservesValue).map((input) => input.cellIndex))
          return {
            changedCellIndices: [...remapped.changedCellIndices, ...remapped.removedCellIndices],
            formulaCellIndices: formulaCellIndices.filter((cellIndex) => !preservedFormulaCellIndices.has(cellIndex)),
          }
        },
        catch: (cause) =>
          new EngineStructureError({
            message: `Failed to apply structural operation ${op.kind}`,
            cause,
          }),
      })
    },
  }
}
