import type { LiteralInput } from "@bilig/protocol";

export type ReplicaId = string;
export type OpId = string;

export interface Clock {
  replicaId: ReplicaId;
  counter: number;
}

export type EngineOp =
  | { kind: "upsertWorkbook"; name: string }
  | { kind: "upsertSheet"; name: string; order: number }
  | { kind: "deleteSheet"; name: string }
  | { kind: "setCellValue"; sheetName: string; address: string; value: LiteralInput }
  | { kind: "setCellFormula"; sheetName: string; address: string; formula: string }
  | { kind: "clearCell"; sheetName: string; address: string };

export interface EngineOpBatch {
  id: OpId;
  replicaId: ReplicaId;
  clock: Clock;
  ops: EngineOp[];
}

export interface ReplicaState {
  replicaId: ReplicaId;
  counter: number;
  appliedBatchIds: Set<OpId>;
}

export function createReplicaState(replicaId: ReplicaId): ReplicaState {
  return {
    replicaId,
    counter: 0,
    appliedBatchIds: new Set<OpId>()
  };
}

export function nextClock(state: ReplicaState): Clock {
  state.counter += 1;
  return { replicaId: state.replicaId, counter: state.counter };
}

export function createBatch(state: ReplicaState, ops: EngineOp[]): EngineOpBatch {
  const clock = nextClock(state);
  const batch: EngineOpBatch = {
    id: `${clock.replicaId}:${clock.counter}`,
    replicaId: state.replicaId,
    clock,
    ops
  };
  state.appliedBatchIds.add(batch.id);
  return batch;
}

export function compactLog(batches: EngineOpBatch[]): EngineOpBatch[] {
  const byId = new Map<OpId, EngineOpBatch>();
  for (const batch of batches) {
    if (!byId.has(batch.id)) {
      byId.set(batch.id, batch);
    }
  }
  return mergeBatches([...byId.values()]);
}

export function mergeBatches(batches: EngineOpBatch[]): EngineOpBatch[] {
  return [...batches].sort((left, right) => {
    if (left.clock.counter !== right.clock.counter) {
      return left.clock.counter - right.clock.counter;
    }
    if (left.replicaId !== right.replicaId) {
      return left.replicaId.localeCompare(right.replicaId);
    }
    return left.id.localeCompare(right.id);
  });
}

export function shouldApplyBatch(state: ReplicaState, batch: EngineOpBatch): boolean {
  if (state.appliedBatchIds.has(batch.id)) {
    return false;
  }
  state.appliedBatchIds.add(batch.id);
  state.counter = Math.max(state.counter, batch.clock.counter);
  return true;
}
