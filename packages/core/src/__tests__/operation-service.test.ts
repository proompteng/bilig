import { Effect } from "effect";
import { describe, expect, it } from "vitest";
import { ValueTag } from "@bilig/protocol";
import type { EngineOpBatch } from "@bilig/workbook-domain";
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

describe("EngineOperationService", () => {
  it("applies remote rename batches through the service and keeps the selection on the renamed sheet", async () => {
    const primary = new SpreadsheetEngine({ workbookName: "operation-rename", replicaId: "a" });
    const replica = new SpreadsheetEngine({ workbookName: "operation-rename", replicaId: "b" });
    await Promise.all([primary.ready(), replica.ready()]);

    const outbound: EngineOpBatch[] = [];
    primary.subscribeBatches((batch) => outbound.push(batch));

    primary.createSheet("Old");
    const createBatch = outbound.at(-1);
    expect(createBatch).toBeDefined();
    replica.applyRemoteBatch(createBatch!);
    replica.setSelection("Old", "B2");

    primary.renameSheet("Old", "New");
    const renameBatch = outbound.at(-1);
    expect(renameBatch).toBeDefined();

    Effect.runSync(getOperationService(replica).applyBatch(renameBatch!, "remote"));

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

    const outbound: EngineOpBatch[] = [];
    primary.subscribeBatches((batch) => outbound.push(batch));

    primary.createSheet("Sheet1");
    const createBatch = outbound.at(-1);
    expect(createBatch).toBeDefined();

    primary.setCellValue("Sheet1", "A1", 7);
    const valueBatch = outbound.at(-1);
    expect(valueBatch).toBeDefined();

    primary.deleteSheet("Sheet1");
    const deleteBatch = outbound.at(-1);
    expect(deleteBatch).toBeDefined();

    replica.applyRemoteBatch(createBatch!);
    replica.applyRemoteBatch(deleteBatch!);

    const restored = new SpreadsheetEngine({ workbookName: "restored", replicaId: "b" });
    await restored.ready();
    restored.importSnapshot(replica.exportSnapshot());
    restored.importReplicaSnapshot(replica.exportReplicaSnapshot());

    Effect.runSync(getOperationService(restored).applyBatch(valueBatch!, "remote"));

    expect(restored.getCellValue("Sheet1", "A1")).toEqual({ tag: ValueTag.Empty });
  });
});
