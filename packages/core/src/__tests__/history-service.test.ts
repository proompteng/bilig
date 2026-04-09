import { Effect, Exit } from "effect";
import { describe, expect, it } from "vitest";
import type { TransactionLogEntry } from "../engine/runtime-state.js";
import { createEngineHistoryService } from "../engine/services/history-service.js";

describe("EngineHistoryService", () => {
  it("restores replay depth even when history replay throws", () => {
    let replayDepth = 0;
    const entry: TransactionLogEntry = {
      forward: { ops: [{ kind: "upsertWorkbook", name: "book" }] },
      inverse: { ops: [{ kind: "upsertWorkbook", name: "book" }] },
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
