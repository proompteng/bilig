import { Effect } from "effect";
import { describe, expect, it, vi } from "vitest";
import { ValueTag } from "@bilig/protocol";
import { SpreadsheetEngine } from "../engine.js";
import type { EngineOperationService } from "../engine/services/operation-service.js";

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

describe("EngineOperationService", () => {
  it("applies remote rename batches through the service and keeps the selection on the renamed sheet", async () => {
    const primary = new SpreadsheetEngine({ workbookName: "operation-rename", replicaId: "a" });
    const replica = new SpreadsheetEngine({ workbookName: "operation-rename", replicaId: "b" });
    await Promise.all([primary.ready(), replica.ready()]);

    const outbound: Parameters<SpreadsheetEngine["applyRemoteBatch"]>[0][] = [];
    primary.subscribeBatches((batch) => outbound.push(batch));

    primary.createSheet("Old");
    const createBatch = expectBatch(outbound.at(-1));
    replica.applyRemoteBatch(createBatch);
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
    const createBatch = expectBatch(outbound.at(-1));

    primary.setCellValue("Sheet1", "A1", 7);
    const valueBatch = expectBatch(outbound.at(-1));

    primary.deleteSheet("Sheet1");
    const deleteBatch = expectBatch(outbound.at(-1));

    replica.applyRemoteBatch(createBatch);
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
});
