import type { DeferredInitialFormulaFamilyRun } from '../engine/services/formula-initialization-family-runs.js'
import type { FormulaFamilyRunAxis } from '../formula/formula-family-store.js'

export interface RuntimeImageFormulaFamilyRunSnapshot {
  readonly sheetName: string
  readonly templateId: number
  readonly shapeKey: string
  readonly axis: FormulaFamilyRunAxis
  readonly fixedIndex: number
  readonly start: number
  readonly step: number
  readonly cellIndices: readonly number[]
}

interface RestoredRuntimeFormulaFamilyRuns {
  readonly runs: DeferredInitialFormulaFamilyRun[]
  readonly fallbackCount: number
}

export function restoreAlignedRuntimeFormulaFamilyRuns(args: {
  readonly runs: readonly RuntimeImageFormulaFamilyRunSnapshot[] | undefined
  readonly sheetIdsByName: ReadonlyMap<string, number>
}): RestoredRuntimeFormulaFamilyRuns | undefined {
  const runtimeRuns = args.runs
  if (!runtimeRuns || runtimeRuns.length === 0) {
    return undefined
  }
  const restoredRuns: DeferredInitialFormulaFamilyRun[] = []
  let fallbackCount = 0
  for (const run of runtimeRuns) {
    const sheetId = args.sheetIdsByName.get(run.sheetName)
    if (
      sheetId === undefined ||
      run.cellIndices.length === 0 ||
      !Number.isInteger(run.templateId) ||
      !Number.isInteger(run.fixedIndex) ||
      !Number.isInteger(run.start) ||
      !Number.isInteger(run.step) ||
      run.step <= 0 ||
      (run.axis !== 'row' && run.axis !== 'column')
    ) {
      fallbackCount += 1
      continue
    }
    restoredRuns.push({
      sheetId,
      templateId: run.templateId,
      shapeKey: run.shapeKey,
      axis: run.axis,
      fixedIndex: run.fixedIndex,
      start: run.start,
      step: run.step,
      lastIndex: run.start + run.step * (run.cellIndices.length - 1),
      ordered: true,
      cellIndices: [...run.cellIndices],
    })
  }
  return { runs: restoredRuns, fallbackCount }
}
