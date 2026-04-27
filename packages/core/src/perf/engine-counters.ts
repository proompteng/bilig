export const ENGINE_COUNTER_KEYS = [
  'cellsRemapped',
  'rangesMaterialized',
  'rangeMembersExpanded',
  'formulasParsed',
  'formulasBound',
  'columnSliceBuilds',
  'exactIndexBuilds',
  'approxIndexBuilds',
  'topoRebuilds',
  'changedCellPayloadsBuilt',
  'snapshotOpsReplayed',
  'wasmFullUploads',
  'directAggregateScanEvaluations',
  'directAggregateScanCells',
  'directAggregatePrefixEvaluations',
  'directAggregateDeltaApplications',
  'directAggregateDeltaOnlyRecalcSkips',
  'directScalarDeltaApplications',
  'directScalarDeltaOnlyRecalcSkips',
  'kernelSyncOnlyRecalcSkips',
  'directFormulaKernelSyncOnlyRecalcSkips',
  'directFormulaInitialEvaluations',
  'structuralTransactions',
  'structuralPlannedCells',
  'structuralSurvivorCellsRemapped',
  'structuralRemovedCells',
  'structuralUndoCapturedCells',
  'structuralFormulaImpactCandidates',
  'structuralFormulaRebindInputs',
  'structuralRangeRetargets',
  'sheetGridBlockScans',
  'axisMapSplices',
  'axisMapMoves',
  'regionQueryIndexBuilds',
  'columnOwnerBuilds',
  'lookupOwnerBuilds',
  'calcChainFullScans',
  'cycleFormulaScans',
  'topoRepairs',
  'topoRepairFailures',
  'topoRepairAffectedFormulas',
] as const

export type EngineCounterKey = (typeof ENGINE_COUNTER_KEYS)[number]

export type EngineCounters = Record<EngineCounterKey, number>

export function createEngineCounters(): EngineCounters {
  return {
    cellsRemapped: 0,
    rangesMaterialized: 0,
    rangeMembersExpanded: 0,
    formulasParsed: 0,
    formulasBound: 0,
    columnSliceBuilds: 0,
    exactIndexBuilds: 0,
    approxIndexBuilds: 0,
    topoRebuilds: 0,
    changedCellPayloadsBuilt: 0,
    snapshotOpsReplayed: 0,
    wasmFullUploads: 0,
    directAggregateScanEvaluations: 0,
    directAggregateScanCells: 0,
    directAggregatePrefixEvaluations: 0,
    directAggregateDeltaApplications: 0,
    directAggregateDeltaOnlyRecalcSkips: 0,
    directScalarDeltaApplications: 0,
    directScalarDeltaOnlyRecalcSkips: 0,
    kernelSyncOnlyRecalcSkips: 0,
    directFormulaKernelSyncOnlyRecalcSkips: 0,
    directFormulaInitialEvaluations: 0,
    structuralTransactions: 0,
    structuralPlannedCells: 0,
    structuralSurvivorCellsRemapped: 0,
    structuralRemovedCells: 0,
    structuralUndoCapturedCells: 0,
    structuralFormulaImpactCandidates: 0,
    structuralFormulaRebindInputs: 0,
    structuralRangeRetargets: 0,
    sheetGridBlockScans: 0,
    axisMapSplices: 0,
    axisMapMoves: 0,
    regionQueryIndexBuilds: 0,
    columnOwnerBuilds: 0,
    lookupOwnerBuilds: 0,
    calcChainFullScans: 0,
    cycleFormulaScans: 0,
    topoRepairs: 0,
    topoRepairFailures: 0,
    topoRepairAffectedFormulas: 0,
  }
}

export function cloneEngineCounters(counters: Readonly<EngineCounters>): EngineCounters {
  return { ...counters }
}

export function addEngineCounter(counters: EngineCounters, key: EngineCounterKey, delta = 1): number {
  const nextValue = counters[key] + delta
  counters[key] = nextValue
  return nextValue
}

export function addEngineCounters(target: EngineCounters, delta: Readonly<Partial<EngineCounters>>): EngineCounters {
  for (const key of ENGINE_COUNTER_KEYS) {
    target[key] += delta[key] ?? 0
  }
  return target
}

export function resetEngineCounters(counters: EngineCounters): EngineCounters {
  for (const key of ENGINE_COUNTER_KEYS) {
    counters[key] = 0
  }
  return counters
}
