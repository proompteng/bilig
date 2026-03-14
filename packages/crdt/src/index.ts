import type { LiteralInput } from "@bilig/protocol";

export type ReplicaId = string;
export type OpId = string;

export interface Clock {
  counter: number;
}

export type EngineOp =
  | { kind: "upsertWorkbook"; name: string }
  | { kind: "upsertSheet"; name: string; order: number }
  | { kind: "deleteSheet"; name: string }
  | { kind: "setCellValue"; sheetName: string; address: string; value: LiteralInput }
  | { kind: "setCellFormula"; sheetName: string; address: string; formula: string }
  | { kind: "setCellFormat"; sheetName: string; address: string; format: string | null }
  | { kind: "clearCell"; sheetName: string; address: string };

export interface EngineOpBatch {
  id: OpId;
  replicaId: ReplicaId;
  clock: Clock;
  ops: EngineOp[];
}

export interface ReplicaState {
  replicaId: ReplicaId;
  clock: Clock;
  appliedBatchIds: Set<OpId>;
}

export interface ReplicaSnapshot {
  replicaId: ReplicaId;
  counter: number;
  appliedBatchIds: OpId[];
}

export interface OpOrder {
  counter: number;
  replicaId: ReplicaId;
  batchId: OpId;
  opIndex: number;
}

export interface ReplicaVersionSnapshot {
  entityKey: string;
  order: OpOrder;
}

export function createReplicaState(replicaId: ReplicaId): ReplicaState {
  return {
    replicaId,
    clock: { counter: 0 },
    appliedBatchIds: new Set<OpId>()
  };
}

export function hydrateReplicaState(state: ReplicaState, snapshot: ReplicaSnapshot): void {
  state.replicaId = snapshot.replicaId;
  state.clock.counter = snapshot.counter;
  state.appliedBatchIds.clear();
  snapshot.appliedBatchIds.forEach((id) => state.appliedBatchIds.add(id));
}

export function exportReplicaSnapshot(state: ReplicaState, limit = 2048): ReplicaSnapshot {
  const appliedBatchIds = [...state.appliedBatchIds].sort();
  const trimmed = appliedBatchIds.slice(Math.max(0, appliedBatchIds.length - limit));
  return {
    replicaId: state.replicaId,
    counter: state.clock.counter,
    appliedBatchIds: trimmed
  };
}

export function importReplicaSnapshot(snapshot: ReplicaSnapshot): ReplicaState {
  const state = createReplicaState(snapshot.replicaId);
  hydrateReplicaState(state, snapshot);
  return state;
}

export function nextClock(state: ReplicaState): Clock {
  state.clock.counter += 1;
  return { counter: state.clock.counter };
}

export function createBatch(state: ReplicaState, ops: EngineOp[]): EngineOpBatch {
  const clock = nextClock(state);
  const batch: EngineOpBatch = {
    id: `${state.replicaId}:${clock.counter}`,
    replicaId: state.replicaId,
    clock,
    ops
  };
  markBatchApplied(state, batch);
  return batch;
}

export function markBatchApplied(state: ReplicaState, batch: EngineOpBatch): void {
  state.appliedBatchIds.add(batch.id);
  state.clock.counter = Math.max(state.clock.counter, batch.clock.counter);
}

export function shouldApplyBatch(state: ReplicaState, batch: EngineOpBatch): boolean {
  return !state.appliedBatchIds.has(batch.id);
}

export function compareBatches(left: EngineOpBatch, right: EngineOpBatch): number {
  return (
    left.clock.counter - right.clock.counter ||
    left.replicaId.localeCompare(right.replicaId) ||
    left.id.localeCompare(right.id)
  );
}

export function batchOpOrder(batch: EngineOpBatch, opIndex: number): OpOrder {
  return {
    counter: batch.clock.counter,
    replicaId: batch.replicaId,
    batchId: batch.id,
    opIndex
  };
}

export function compareOpOrder(left: OpOrder, right: OpOrder): number {
  return (
    left.counter - right.counter ||
    left.replicaId.localeCompare(right.replicaId) ||
    left.batchId.localeCompare(right.batchId) ||
    left.opIndex - right.opIndex
  );
}

function entityKeyForOp(op: EngineOp): string {
  switch (op.kind) {
    case "upsertWorkbook":
      return "workbook";
    case "upsertSheet":
    case "deleteSheet":
      return `sheet:${op.name}`;
    case "setCellValue":
    case "setCellFormula":
    case "clearCell":
      return `cell:${op.sheetName}!${op.address}`;
    case "setCellFormat":
      return `format:${op.sheetName}!${op.address}`;
  }
}

function sheetDeleteBarrierForOp(op: EngineOp, latestSheetDeletes: Map<string, OpOrder>): OpOrder | undefined {
  switch (op.kind) {
    case "setCellValue":
    case "setCellFormula":
    case "setCellFormat":
    case "clearCell":
      return latestSheetDeletes.get(op.sheetName);
    case "upsertSheet":
      return latestSheetDeletes.get(op.name);
    default:
      return undefined;
  }
}

export function compactLog(batches: EngineOpBatch[]): EngineOpBatch[] {
  const deduped = new Map<OpId, EngineOpBatch>();
  batches.forEach((batch) => {
    deduped.set(batch.id, batch);
  });
  const ordered = [...deduped.values()].sort(compareBatches);
  const latestByEntity = new Map<string, OpOrder>();
  const latestSheetDeletes = new Map<string, OpOrder>();

  ordered.forEach((batch) => {
    batch.ops.forEach((op, opIndex) => {
      const order = batchOpOrder(batch, opIndex);
      latestByEntity.set(entityKeyForOp(op), order);
      if (op.kind === "deleteSheet") {
        latestSheetDeletes.set(op.name, order);
      }
    });
  });

  return ordered
    .map((batch) => ({
      ...batch,
      ops: batch.ops.filter((op, opIndex) => {
        const order = batchOpOrder(batch, opIndex);
        const sheetDeleteBarrier = sheetDeleteBarrierForOp(op, latestSheetDeletes);
        if (sheetDeleteBarrier && compareOpOrder(order, sheetDeleteBarrier) <= 0) {
          return false;
        }

        const latestOrder = latestByEntity.get(entityKeyForOp(op));
        return latestOrder !== undefined && compareOpOrder(order, latestOrder) === 0;
      })
    }))
    .filter((batch) => batch.ops.length > 0);
}

export function mergeBatches(batches: EngineOpBatch[]): EngineOpBatch[] {
  const deduped = new Map<OpId, EngineOpBatch>();
  batches.forEach((batch) => {
    deduped.set(batch.id, batch);
  });
  return [...deduped.values()].sort(compareBatches);
}
