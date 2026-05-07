import type { FormulaFamilyStore } from '../../formula/formula-family-store.js'
import type { EngineRuntimeState } from '../runtime-state.js'
import { getFormulaBindingFamilyShapeKey, type FormulaBindingFamilyShapeKeyCache } from './formula-binding-family-shape-key.js'

export interface DeferredFormulaFamilyIndexRun {
  readonly sheetId: number
  readonly templateId: number
  readonly shapeKey: string
  readonly fixedIndex: number
  start: number
  step: number
  lastIndex: number
  ordered: boolean
  cellIndices: number[]
  rows?: number[]
}

export function noteDeferredFormulaFamilyIndexRunMember(
  runs: Map<string, DeferredFormulaFamilyIndexRun>,
  args: {
    readonly cellIndex: number
    readonly sheetId: number
    readonly templateId: number
    readonly shapeKey: string
    readonly row: number
    readonly col: number
  },
): void {
  const familyKey = `${args.sheetId}\t${args.templateId}\t${args.shapeKey}\t${args.col}`
  let run = runs.get(familyKey)
  if (!run) {
    run = {
      sheetId: args.sheetId,
      templateId: args.templateId,
      shapeKey: args.shapeKey,
      fixedIndex: args.col,
      start: args.row,
      step: 0,
      lastIndex: args.row,
      ordered: true,
      cellIndices: [],
    }
    runs.set(familyKey, run)
  } else {
    const nextStep = args.row - run.lastIndex
    let breaksOrder = false
    if (run.cellIndices.length === 1) {
      run.step = nextStep
    } else if (run.step !== nextStep) {
      breaksOrder = true
    }
    if (args.row <= run.lastIndex || args.col !== run.fixedIndex) {
      breaksOrder = true
    }
    if (breaksOrder) {
      if (!run.rows) {
        const priorStep = run.cellIndices.length <= 1 ? 1 : run.step
        const start = run.start
        run.rows = Array.from({ length: run.cellIndices.length }, (_value, index) => start + priorStep * index)
      }
      run.ordered = false
    }
    run.lastIndex = args.row
  }
  run.cellIndices.push(args.cellIndex)
  run.rows?.push(args.row)
}

export function registerDeferredFormulaFamilyIndexRun(
  store: Pick<FormulaFamilyStore, 'registerFormulaRun' | 'registerFreshUniformRun'>,
  run: DeferredFormulaFamilyIndexRun,
): void {
  const step = run.cellIndices.length <= 1 ? 1 : run.step
  if (
    run.ordered &&
    step > 0 &&
    store.registerFreshUniformRun({
      sheetId: run.sheetId,
      templateId: run.templateId,
      shapeKey: run.shapeKey,
      axis: 'row',
      fixedIndex: run.fixedIndex,
      start: run.start,
      step,
      cellIndices: run.cellIndices,
    })
  ) {
    return
  }
  store.registerFormulaRun({
    sheetId: run.sheetId,
    templateId: run.templateId,
    shapeKey: run.shapeKey,
    members: run.cellIndices.map((cellIndex, index) => ({
      cellIndex,
      row: run.ordered ? run.start + step * index : run.rows![index]!,
      col: run.fixedIndex,
    })),
  })
}

export function rebuildDeferredFormulaFamilyIndex(args: {
  readonly state: Pick<EngineRuntimeState, 'workbook' | 'formulas'>
  readonly store: FormulaFamilyStore
  readonly shapeKeyCache: FormulaBindingFamilyShapeKeyCache
}): void {
  const runs = new Map<string, DeferredFormulaFamilyIndexRun>()
  args.store.clear()
  args.shapeKeyCache.clear()
  args.state.formulas.forEach((formula, cellIndex) => {
    const templateId = formula.templateId
    const sheetId = args.state.workbook.cellStore.sheetIds[cellIndex]
    const position = args.state.workbook.getCellPosition(cellIndex)
    if (templateId === undefined || sheetId === undefined || !position) {
      return
    }
    noteDeferredFormulaFamilyIndexRunMember(runs, {
      cellIndex,
      sheetId,
      templateId,
      shapeKey: getFormulaBindingFamilyShapeKey(args.shapeKeyCache, formula),
      row: position.row,
      col: position.col,
    })
  })
  runs.forEach((run) => {
    registerDeferredFormulaFamilyIndexRun(args.store, run)
  })
}
