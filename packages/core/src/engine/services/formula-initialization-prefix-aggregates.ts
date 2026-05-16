import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import { CellFlags } from '../../cell-store.js'
import type { InitialFormulaCellIndexList } from './formula-initialization-refs.js'
import type { EngineFormulaInitializationServiceArgs } from './formula-initialization-service-types.js'

export type InitialPrefixAggregateKind = 'sum' | 'count' | 'average' | 'min' | 'max'

export interface InitialPrefixAggregateGroup {
  readonly sheetName: string
  readonly col: number
  readonly colEnd: number
  readonly aggregateKind: InitialPrefixAggregateKind
  maxRowEnd: number
  lastRowEnd: number
  formulasAreOrdered: boolean
  readonly formulas: Array<{ cellIndex: number; rowEnd: number; resultOffset?: number }>
}

export function evaluateInitialPrefixAggregateGroups(
  args: Pick<EngineFormulaInitializationServiceArgs, 'state' | 'checkEvaluationBudget'>,
  orderedCellIndices: InitialFormulaCellIndexList,
  pushChangedCellIndex: (cellIndex: number) => void,
  writeFormulaValue: (cellIndex: number, value: CellValue) => void,
): Set<number> | undefined {
  const groups = new Map<string, InitialPrefixAggregateGroup>()
  for (let index = 0; index < orderedCellIndices.length; index += 1) {
    args.checkEvaluationBudget()
    const cellIndex = orderedCellIndices[index]!
    const formula = args.state.formulas.get(cellIndex)
    const aggregate = formula?.directAggregate
    if (!formula || !aggregate || aggregate.rowStart !== 0 || formula.dependencyIndices.length !== 0) {
      continue
    }
    const key = `${aggregate.sheetName}\t${aggregate.col}\t${aggregate.colEnd}\t${aggregate.aggregateKind}`
    let group = groups.get(key)
    if (!group) {
      group = {
        sheetName: aggregate.sheetName,
        col: aggregate.col,
        colEnd: aggregate.colEnd,
        aggregateKind: aggregate.aggregateKind,
        maxRowEnd: aggregate.rowEnd,
        lastRowEnd: aggregate.rowEnd,
        formulasAreOrdered: true,
        formulas: [],
      }
      groups.set(key, group)
    } else {
      group.maxRowEnd = Math.max(group.maxRowEnd, aggregate.rowEnd)
      if (aggregate.rowEnd < group.lastRowEnd) {
        group.formulasAreOrdered = false
      }
      group.lastRowEnd = aggregate.rowEnd
    }
    group.formulas.push({
      cellIndex,
      rowEnd: aggregate.rowEnd,
      ...(aggregate.resultOffset !== undefined ? { resultOffset: aggregate.resultOffset } : {}),
    })
  }
  if (groups.size === 0) {
    return undefined
  }

  const handled = new Set<number>()
  groups.forEach((group) => {
    const sheet = args.state.workbook.getSheet(group.sheetName)
    if (!sheet) {
      return
    }
    const formulas = group.formulasAreOrdered ? group.formulas : group.formulas.toSorted((left, right) => left.rowEnd - right.rowEnd)
    let sum = 0
    let count = 0
    let averageCount = 0
    let errorCode = ErrorCode.None
    let errorCount = 0
    let minimum = Number.POSITIVE_INFINITY
    let maximum = Number.NEGATIVE_INFINITY
    let formulaIndex = 0
    let encounteredFormulaMember = false
    for (let row = 0; row <= group.maxRowEnd && !encounteredFormulaMember; row += 1) {
      for (let col = group.col; col <= group.colEnd; col += 1) {
        const memberCellIndex = sheet.structureVersion === 1 ? sheet.grid.getPhysical(row, col) : sheet.grid.get(row, col)
        if (memberCellIndex !== -1) {
          if (((args.state.workbook.cellStore.flags[memberCellIndex] ?? 0) & CellFlags.HasFormula) !== 0) {
            encounteredFormulaMember = true
            break
          }
          const tag = (args.state.workbook.cellStore.tags[memberCellIndex] as ValueTag | undefined) ?? ValueTag.Empty
          if (tag === ValueTag.Number) {
            const numeric = args.state.workbook.cellStore.numbers[memberCellIndex] ?? 0
            sum += numeric
            count += 1
            averageCount += 1
            minimum = Math.min(minimum, numeric)
            maximum = Math.max(maximum, numeric)
          } else if (tag === ValueTag.Boolean) {
            const numeric = (args.state.workbook.cellStore.numbers[memberCellIndex] ?? 0) !== 0 ? 1 : 0
            sum += numeric
            count += 1
            averageCount += 1
            minimum = Math.min(minimum, numeric)
            maximum = Math.max(maximum, numeric)
          } else if (tag === ValueTag.Empty) {
            minimum = Math.min(minimum, 0)
            maximum = Math.max(maximum, 0)
          } else if (tag === ValueTag.Error) {
            errorCode ||= (args.state.workbook.cellStore.errors[memberCellIndex] as ErrorCode | undefined) ?? ErrorCode.None
            errorCount += 1
          }
        }
      }
      while (formulaIndex < formulas.length && formulas[formulaIndex]!.rowEnd <= row) {
        const formula = formulas[formulaIndex]!
        const aggregateValue =
          group.aggregateKind === 'sum'
            ? errorCount > 0 && errorCode !== ErrorCode.None
              ? { tag: ValueTag.Error as const, code: errorCode }
              : { tag: ValueTag.Number as const, value: sum }
            : group.aggregateKind === 'count'
              ? { tag: ValueTag.Number as const, value: count }
              : group.aggregateKind === 'average'
                ? errorCount > 0 && errorCode !== ErrorCode.None
                  ? { tag: ValueTag.Error as const, code: errorCode }
                  : averageCount === 0
                    ? { tag: ValueTag.Error as const, code: ErrorCode.Div0 }
                    : { tag: ValueTag.Number as const, value: sum / averageCount }
                : group.aggregateKind === 'min'
                  ? { tag: ValueTag.Number as const, value: minimum === Number.POSITIVE_INFINITY ? 0 : minimum }
                  : { tag: ValueTag.Number as const, value: maximum === Number.NEGATIVE_INFINITY ? 0 : maximum }
        const value =
          formula.resultOffset !== undefined && aggregateValue.tag === ValueTag.Number
            ? { tag: ValueTag.Number as const, value: aggregateValue.value + formula.resultOffset }
            : aggregateValue
        writeFormulaValue(formula.cellIndex, value)
        handled.add(formula.cellIndex)
        pushChangedCellIndex(formula.cellIndex)
        formulaIndex += 1
      }
    }
  })
  return handled.size === 0 ? undefined : handled
}
