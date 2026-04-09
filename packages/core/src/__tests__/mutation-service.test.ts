import { Effect } from "effect";
import { describe, expect, it } from "vitest";
import { createReplicaState } from "../replica-state.js";
import { createEngineMutationService } from "../engine/services/mutation-service.js";

describe("EngineMutationService", () => {
  it("clears redo history when a new local transaction lands", () => {
    const replicaState = createReplicaState("local");
    let replayDepth = 0;
    const batches: import("@bilig/workbook-domain").EngineOpBatch[] = [];
    const service = createEngineMutationService({
      state: {
        replicaState,
        undoStack: [],
        redoStack: [{ forward: { ops: [] }, inverse: { ops: [] } }],
        getTransactionReplayDepth: () => replayDepth,
        setTransactionReplayDepth: (next) => {
          replayDepth = next;
        },
      },
      buildInverseOps: () => [{ kind: "upsertWorkbook", name: "inverse" }],
      applyBatchNow: (batch) => {
        batches.push(batch);
      },
    });

    const undoOps = Effect.runSync(
      service.executeLocal([{ kind: "upsertWorkbook", name: "forward" }]),
    );

    expect(undoOps).toEqual([{ kind: "upsertWorkbook", name: "inverse" }]);
    expect(batches).toHaveLength(1);
    expect(service).toBeDefined();
  });

  it("captures a single local transaction and clones the undo ops", () => {
    const replicaState = createReplicaState("local");
    let replayDepth = 0;
    const state = {
      replicaState,
      undoStack: [] as Array<{ forward: { ops: unknown[] }; inverse: { ops: unknown[] } }>,
      redoStack: [] as Array<{ forward: { ops: unknown[] }; inverse: { ops: unknown[] } }>,
      getTransactionReplayDepth: () => replayDepth,
      setTransactionReplayDepth: (next: number) => {
        replayDepth = next;
      },
    };
    const service = createEngineMutationService({
      state,
      buildInverseOps: () => [{ kind: "upsertWorkbook", name: "inverse" }],
      applyBatchNow: () => {},
    });

    const captured = Effect.runSync(
      service.captureUndoOps(() =>
        Effect.runSync(service.executeLocal([{ kind: "upsertWorkbook", name: "forward" }])),
      ),
    );

    expect(captured.undoOps).toEqual([{ kind: "upsertWorkbook", name: "inverse" }]);
  });

  it("drops malformed render commit records instead of forwarding partial engine ops", () => {
    const replicaState = createReplicaState("local");
    let replayDepth = 0;
    const batches: import("@bilig/workbook-domain").EngineOpBatch[] = [];
    const service = createEngineMutationService({
      state: {
        replicaState,
        undoStack: [],
        redoStack: [],
        getTransactionReplayDepth: () => replayDepth,
        setTransactionReplayDepth: (next) => {
          replayDepth = next;
        },
      },
      buildInverseOps: () => [],
      applyBatchNow: (batch) => {
        batches.push(batch);
      },
    });

    Effect.runSync(
      service.renderCommit([
        { kind: "renameSheet", oldName: "Old" },
        { kind: "upsertCell", sheetName: "Sheet1" },
        { kind: "upsertSheet", name: "Sheet1", order: 0 },
        { kind: "upsertCell", sheetName: "Sheet1", addr: "A1", value: 7, format: "0.00" },
      ]),
    );

    expect(batches).toHaveLength(1);
    expect(batches[0]?.ops).toEqual([
      { kind: "upsertSheet", name: "Sheet1", order: 0 },
      { kind: "setCellValue", sheetName: "Sheet1", address: "A1", value: 7 },
      { kind: "setCellFormat", sheetName: "Sheet1", address: "A1", format: "0.00" },
    ]);
  });
});
