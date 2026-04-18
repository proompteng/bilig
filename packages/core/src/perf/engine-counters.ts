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
