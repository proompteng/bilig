import type {
  Clock as WorkbookDomainClock,
  EngineOp,
  EngineOpBatch,
  OpId,
  ReplicaId,
} from "@bilig/workbook-domain";

type Clock = WorkbookDomainClock;

export interface ReplicaState {
  replicaId: ReplicaId;
  clock: Clock;
  appliedBatchIds: Set<OpId>;
}

export interface OpOrder {
  counter: number;
  replicaId: ReplicaId;
  batchId: OpId;
  opIndex: number;
}

export interface ReplicaSnapshot {
  replicaId: ReplicaId;
  counter: number;
  appliedBatchIds: OpId[];
}

export interface ReplicaVersionSnapshot {
  entityKey: string;
  order: OpOrder;
}

function normalizedDefinedName(name: string): string {
  return name.trim().toUpperCase();
}

function pivotEntityKey(sheetName: string, address: string): string {
  return `pivot:${sheetName}!${address}`;
}

export function createReplicaState(replicaId: ReplicaId): ReplicaState {
  return {
    replicaId,
    clock: { counter: 0 },
    appliedBatchIds: new Set<OpId>(),
  };
}

export function hydrateReplicaState(state: ReplicaState, snapshot: ReplicaSnapshot): void {
  state.replicaId = snapshot.replicaId;
  state.clock.counter = snapshot.counter;
  state.appliedBatchIds.clear();
  snapshot.appliedBatchIds.forEach((id) => state.appliedBatchIds.add(id));
}

export function exportReplicaSnapshot(state: ReplicaState, limit = 2048): ReplicaSnapshot {
  const appliedBatchIds = [...state.appliedBatchIds].toSorted();
  const trimmed = appliedBatchIds.slice(Math.max(0, appliedBatchIds.length - limit));
  return {
    replicaId: state.replicaId,
    counter: state.clock.counter,
    appliedBatchIds: trimmed,
  };
}

export function importReplicaSnapshot(snapshot: ReplicaSnapshot): ReplicaState {
  const state = createReplicaState(snapshot.replicaId);
  hydrateReplicaState(state, snapshot);
  return state;
}

function nextClock(state: ReplicaState): Clock {
  state.clock.counter += 1;
  return { counter: state.clock.counter };
}

export function createBatch(state: ReplicaState, ops: EngineOp[]): EngineOpBatch {
  const clock = nextClock(state);
  const batch: EngineOpBatch = {
    id: `${state.replicaId}:${clock.counter}`,
    replicaId: state.replicaId,
    clock,
    ops,
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

function compareBatches(left: EngineOpBatch, right: EngineOpBatch): number {
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
    opIndex,
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

function assertNever(value: never): never {
  throw new Error(`Unhandled engine op kind: ${JSON.stringify(value)}`);
}

function entityKeyForOp(op: EngineOp): string {
  switch (op.kind) {
    case "upsertWorkbook":
      return "workbook";
    case "setWorkbookMetadata":
      return `workbook-meta:${op.key}`;
    case "setCalculationSettings":
      return "workbook-calc";
    case "setVolatileContext":
      return "workbook-volatile";
    case "upsertSheet":
    case "deleteSheet":
      return `sheet:${op.name}`;
    case "renameSheet":
      return `sheet:${op.oldName}`;
    case "insertRows":
    case "deleteRows":
    case "moveRows":
      return `row-structure:${op.sheetName}`;
    case "insertColumns":
    case "deleteColumns":
    case "moveColumns":
      return `column-structure:${op.sheetName}`;
    case "updateRowMetadata":
      return `row-meta:${op.sheetName}:${op.start}:${op.count}`;
    case "updateColumnMetadata":
      return `column-meta:${op.sheetName}:${op.start}:${op.count}`;
    case "setFreezePane":
    case "clearFreezePane":
      return `freeze:${op.sheetName}`;
    case "setFilter":
    case "clearFilter":
      return `filter:${op.sheetName}:${op.range.startAddress}:${op.range.endAddress}`;
    case "setSort":
    case "clearSort":
      return `sort:${op.sheetName}:${op.range.startAddress}:${op.range.endAddress}`;
    case "setDataValidation":
      return `validation:${op.validation.range.sheetName}:${op.validation.range.startAddress}:${op.validation.range.endAddress}`;
    case "clearDataValidation":
      return `validation:${op.sheetName}:${op.range.startAddress}:${op.range.endAddress}`;
    case "setCellValue":
    case "setCellFormula":
    case "clearCell":
      return `cell:${op.sheetName}!${op.address}`;
    case "setCellFormat":
      return `format:${op.sheetName}!${op.address}`;
    case "upsertCellStyle":
      return `style:${op.style.id}`;
    case "upsertCellNumberFormat":
      return `number-format:${op.format.id}`;
    case "setStyleRange":
      return `style-range:${op.range.sheetName}:${op.range.startAddress}:${op.range.endAddress}`;
    case "setFormatRange":
      return `format-range:${op.range.sheetName}:${op.range.startAddress}:${op.range.endAddress}`;
    case "upsertDefinedName":
    case "deleteDefinedName":
      return `defined-name:${normalizedDefinedName(op.name)}`;
    case "upsertTable":
      return `table:${normalizedDefinedName(op.table.name)}`;
    case "deleteTable":
      return `table:${normalizedDefinedName(op.name)}`;
    case "upsertSpillRange":
    case "deleteSpillRange":
      return `spill:${op.sheetName}!${op.address}`;
    case "upsertPivotTable":
    case "deletePivotTable":
      return pivotEntityKey(op.sheetName, op.address);
  }
  return assertNever(op);
}

function sheetDeleteBarrierForOp(
  op: EngineOp,
  latestSheetDeletes: Map<string, OpOrder>,
): OpOrder | undefined {
  switch (op.kind) {
    case "upsertWorkbook":
    case "setWorkbookMetadata":
    case "setCalculationSettings":
    case "setVolatileContext":
    case "deleteSheet":
    case "upsertDefinedName":
    case "deleteDefinedName":
    case "upsertTable":
    case "deleteTable":
      return undefined;
    case "upsertSheet":
      return latestSheetDeletes.get(op.name);
    case "renameSheet":
      return latestSheetDeletes.get(op.oldName);
    case "insertRows":
    case "deleteRows":
    case "moveRows":
    case "insertColumns":
    case "deleteColumns":
    case "moveColumns":
      return latestSheetDeletes.get(op.sheetName);
    case "updateRowMetadata":
    case "updateColumnMetadata":
    case "setFreezePane":
    case "clearFreezePane":
    case "setFilter":
    case "clearFilter":
    case "setSort":
    case "clearSort":
    case "clearDataValidation":
    case "setCellValue":
    case "setCellFormula":
    case "setCellFormat":
    case "clearCell":
    case "upsertSpillRange":
    case "deleteSpillRange":
    case "deletePivotTable":
      return latestSheetDeletes.get(op.sheetName);
    case "setStyleRange":
    case "setFormatRange":
      return latestSheetDeletes.get(op.range.sheetName);
    case "upsertCellNumberFormat":
    case "upsertCellStyle":
      return undefined;
    case "setDataValidation":
      return latestSheetDeletes.get(op.validation.range.sheetName);
    case "upsertPivotTable":
      return latestSheetDeletes.get(op.sheetName) ?? latestSheetDeletes.get(op.source.sheetName);
  }
  return assertNever(op);
}

export function compactLog(batches: EngineOpBatch[]): EngineOpBatch[] {
  const deduped = new Map<OpId, EngineOpBatch>();
  batches.forEach((batch) => {
    deduped.set(batch.id, batch);
  });
  const ordered = [...deduped.values()].toSorted(compareBatches);
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
    .map((batch) => {
      const ops = batch.ops.filter((op, opIndex) => {
        const order = batchOpOrder(batch, opIndex);
        const sheetDeleteBarrier = sheetDeleteBarrierForOp(op, latestSheetDeletes);
        if (sheetDeleteBarrier && compareOpOrder(order, sheetDeleteBarrier) <= 0) {
          return false;
        }

        const latestOrder = latestByEntity.get(entityKeyForOp(op));
        return latestOrder !== undefined && compareOpOrder(order, latestOrder) === 0;
      });
      return {
        id: batch.id,
        replicaId: batch.replicaId,
        clock: batch.clock,
        ops,
      };
    })
    .filter((batch) => batch.ops.length > 0);
}

export function mergeBatches(batches: EngineOpBatch[]): EngineOpBatch[] {
  const deduped = new Map<OpId, EngineOpBatch>();
  batches.forEach((batch) => {
    deduped.set(batch.id, batch);
  });
  return [...deduped.values()].toSorted(compareBatches);
}
