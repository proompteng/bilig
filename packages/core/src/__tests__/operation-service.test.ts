import { Effect } from "effect";
import { describe, expect, it, vi } from "vitest";
import { ValueTag } from "@bilig/protocol";
import { createBatch } from "../replica-state.js";
import { SpreadsheetEngine } from "../engine.js";
import type { EngineOperationService } from "../engine/services/operation-service.js";
import { cellMutationRefToEngineOp, type EngineCellMutationRef } from "../cell-mutations-at.js";

function isEngineOperationService(value: unknown): value is EngineOperationService {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  return (
    typeof Reflect.get(value, "applyBatch") === "function" &&
    typeof Reflect.get(value, "applyDerivedOp") === "function"
  );
}

function getOperationService(engine: SpreadsheetEngine): EngineOperationService {
  const runtime = Reflect.get(engine, "runtime");
  if (typeof runtime !== "object" || runtime === null) {
    throw new TypeError("Expected engine runtime");
  }
  const operations = Reflect.get(runtime, "operations");
  if (!isEngineOperationService(operations)) {
    throw new TypeError("Expected engine operation service");
  }
  return operations;
}

function hasRuntimeStateMetrics(value: unknown): value is {
  getLastMetrics(): unknown;
  setLastMetrics(metrics: unknown): void;
} {
  if (typeof value !== "object" || value === null) {
    return false;
  }
  const getLastMetrics = Reflect.get(value, "getLastMetrics");
  const setLastMetrics = Reflect.get(value, "setLastMetrics");
  return typeof getLastMetrics === "function" && typeof setLastMetrics === "function";
}

function getRuntimeState(engine: SpreadsheetEngine): {
  getLastMetrics(): unknown;
  setLastMetrics(metrics: unknown): void;
} {
  const state = Reflect.get(engine, "state");
  if (!hasRuntimeStateMetrics(state)) {
    throw new TypeError("Expected runtime state metric accessors");
  }
  return state;
}

function expectBatch<Batch>(batch: Batch | undefined): Batch {
  expect(batch).toBeDefined();
  return batch;
}

function getReplicaState(engine: SpreadsheetEngine) {
  const replicaState = Reflect.get(engine, "replicaState");
  if (typeof replicaState !== "object" || replicaState === null) {
    throw new TypeError("Expected engine replica state");
  }
  return replicaState;
}

describe("EngineOperationService", () => {
  it("applies remote rename batches through the service and keeps the selection on the renamed sheet", async () => {
    const primary = new SpreadsheetEngine({ workbookName: "operation-rename", replicaId: "a" });
    const replica = new SpreadsheetEngine({ workbookName: "operation-rename", replicaId: "b" });
    await Promise.all([primary.ready(), replica.ready()]);

    const outbound: Parameters<SpreadsheetEngine["applyRemoteBatch"]>[0][] = [];
    primary.subscribeBatches((batch) => outbound.push(batch));

    primary.createSheet("Old");
    const createdSheetBatch = expectBatch(outbound.at(-1));
    replica.applyRemoteBatch(createdSheetBatch);
    replica.setSelection("Old", "B2");

    primary.renameSheet("Old", "New");
    const renameBatch = expectBatch(outbound.at(-1));

    Effect.runSync(getOperationService(replica).applyBatch(renameBatch, "remote"));

    expect(replica.getSelectionState()).toMatchObject({
      sheetName: "New",
      address: "B2",
      anchorAddress: "B2",
    });
  });

  it("rejects stale remote cell replays behind sheet tombstones through the service", async () => {
    const primary = new SpreadsheetEngine({ workbookName: "operation-tombstone", replicaId: "a" });
    const replica = new SpreadsheetEngine({ workbookName: "operation-tombstone", replicaId: "b" });
    await Promise.all([primary.ready(), replica.ready()]);

    const outbound: Parameters<SpreadsheetEngine["applyRemoteBatch"]>[0][] = [];
    primary.subscribeBatches((batch) => outbound.push(batch));

    primary.createSheet("Sheet1");
    const createdSheetBatch = expectBatch(outbound.at(-1));

    primary.setCellValue("Sheet1", "A1", 7);
    const valueBatch = expectBatch(outbound.at(-1));

    primary.deleteSheet("Sheet1");
    const deleteBatch = expectBatch(outbound.at(-1));

    replica.applyRemoteBatch(createdSheetBatch);
    replica.applyRemoteBatch(deleteBatch);

    const restored = new SpreadsheetEngine({ workbookName: "restored", replicaId: "b" });
    await restored.ready();
    restored.importSnapshot(replica.exportSnapshot());
    restored.importReplicaSnapshot(replica.exportReplicaSnapshot());

    Effect.runSync(getOperationService(restored).applyBatch(valueBatch, "remote"));

    expect(restored.getCellValue("Sheet1", "A1")).toEqual({ tag: ValueTag.Empty });
  });

  it("does not rewrite last metrics once per formula during snapshot restore", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "operation-restore-metrics" });
    await engine.ready();

    const state = getRuntimeState(engine);
    const setLastMetricsSpy = vi.spyOn(state, "setLastMetrics");
    setLastMetricsSpy.mockClear();

    engine.importSnapshot({
      version: 1,
      workbook: { name: "operation-restore-metrics" },
      sheets: [
        {
          id: 1,
          name: "Sheet1",
          order: 0,
          cells: [
            { address: "A1", value: 1 },
            { address: "B1", formula: "A1*2" },
            { address: "C1", formula: "B1+1" },
          ],
        },
      ],
    });

    expect(setLastMetricsSpy).toHaveBeenCalledTimes(3);
  });

  it("applies local cell mutation refs through the service", async () => {
    const engine = new SpreadsheetEngine({ workbookName: "operation-local-refs", replicaId: "a" });
    await engine.ready();
    engine.createSheet("Sheet1");
    engine.setCellValue("Sheet1", "A1", 1);
    const sheetId = engine.workbook.getSheet("Sheet1")!.id;
    const refs: EngineCellMutationRef[] = [
      {
        sheetId,
        mutation: { kind: "setCellValue", row: 0, col: 0, value: 10 },
      },
      {
        sheetId,
        mutation: { kind: "setCellFormula", row: 0, col: 1, formula: "A1*2" },
      },
      {
        sheetId,
        mutation: { kind: "setCellFormula", row: 0, col: 2, formula: "SUM(" },
      },
      {
        sheetId,
        mutation: { kind: "clearCell", row: 3, col: 3 },
      },
      {
        sheetId,
        mutation: { kind: "clearCell", row: 0, col: 0 },
      },
    ];
    const forwardOps = refs.map((ref) => cellMutationRefToEngineOp(engine.workbook, ref));
    const batch = createBatch(getReplicaState(engine), forwardOps);

    Effect.runSync(getOperationService(engine).applyLocalCellMutationsAt(refs, batch, 3));

    expect(engine.getCellValue("Sheet1", "A1")).toEqual({ tag: ValueTag.Empty });
    expect(engine.getCellValue("Sheet1", "B1")).toEqual({ tag: ValueTag.Number, value: 0 });
    expect(engine.getCellValue("Sheet1", "C1")).toMatchObject({
      tag: ValueTag.Error,
      code: expect.any(Number),
    });
    expect(engine.getCellValue("Sheet1", "D4")).toEqual({ tag: ValueTag.Empty });
  });

  it("rejects local cell mutation refs for unknown sheets", async () => {
    const engine = new SpreadsheetEngine({
      workbookName: "operation-local-refs-missing",
      replicaId: "a",
    });
    await engine.ready();
    engine.createSheet("Sheet1");
    const refs: EngineCellMutationRef[] = [
      {
        sheetId: 999,
        mutation: { kind: "setCellValue", row: 0, col: 0, value: 1 },
      },
    ];
    const batch = createBatch(getReplicaState(engine), [
      { kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 1 },
    ]);

    expect(() =>
      Effect.runSync(getOperationService(engine).applyLocalCellMutationsAt(refs, batch, 1)),
    ).toThrow("Unknown sheet id: 999");
  });
});
