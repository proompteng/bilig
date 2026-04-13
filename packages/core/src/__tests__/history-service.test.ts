import { Effect, Exit } from "effect";
import { describe, expect, it } from "vitest";
import type { TransactionLogEntry } from "../engine/runtime-state.js";
import { createEngineHistoryService } from "../engine/services/history-service.js";

describe("EngineHistoryService", () => {
  it("returns false when undo and redo stacks are empty", () => {
    const service = createEngineHistoryService({
      state: {
        undoStack: [],
        redoStack: [],
        getTransactionReplayDepth: () => 0,
        setTransactionReplayDepth: () => {},
      },
      executeTransaction: () => {},
    });

    expect(Effect.runSync(service.undo())).toBe(false);
    expect(Effect.runSync(service.redo())).toBe(false);
  });

  it("moves entries between undo and redo stacks on successful replay", () => {
    let replayDepth = 0;
    const executed: Array<"undo" | "redo"> = [];
    const entry: TransactionLogEntry = {
      forward: { kind: "ops", ops: [{ kind: "upsertWorkbook", name: "forward" }] },
      inverse: { kind: "ops", ops: [{ kind: "upsertWorkbook", name: "inverse" }] },
    };
    const undoStack: TransactionLogEntry[] = [entry];
    const redoStack: TransactionLogEntry[] = [];
    const service = createEngineHistoryService({
      state: {
        undoStack,
        redoStack,
        getTransactionReplayDepth: () => replayDepth,
        setTransactionReplayDepth: (next) => {
          replayDepth = next;
        },
      },
      executeTransaction: (transaction) => {
        executed.push(transaction === entry.inverse ? "undo" : "redo");
      },
    });

    expect(Effect.runSync(service.undo())).toBe(true);
    expect(executed).toEqual(["undo"]);
    expect(undoStack).toEqual([]);
    expect(redoStack).toEqual([entry]);
    expect(replayDepth).toBe(0);

    expect(Effect.runSync(service.redo())).toBe(true);
    expect(executed).toEqual(["undo", "redo"]);
    expect(undoStack).toEqual([entry]);
    expect(redoStack).toEqual([]);
    expect(replayDepth).toBe(0);
  });

  it("restores replay depth even when history replay throws", () => {
    let replayDepth = 0;
    const entry: TransactionLogEntry = {
      forward: { kind: "ops", ops: [{ kind: "upsertWorkbook", name: "book" }] },
      inverse: { kind: "ops", ops: [{ kind: "upsertWorkbook", name: "book" }] },
    };
    const service = createEngineHistoryService({
      state: {
        undoStack: [entry],
        redoStack: [],
        getTransactionReplayDepth: () => replayDepth,
        setTransactionReplayDepth: (next) => {
          replayDepth = next;
        },
      },
      executeTransaction: () => {
        throw new Error("boom");
      },
    });

    const exit = Effect.runSyncExit(service.undo());

    expect(Exit.isFailure(exit)).toBe(true);
    expect(replayDepth).toBe(0);
  });
});
