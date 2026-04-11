import { describe, expect, it } from "vitest";
import {
  EngineCellStateError,
  EngineFormulaBindingError,
  EngineFormulaEvaluationError,
  EngineFormulaGraphError,
  EngineHistoryError,
  EngineMaintenanceError,
  EngineMutationError,
  EnginePivotError,
  EngineRecalcError,
  EngineRuntimeScratchError,
  EngineSnapshotError,
  EngineStructureError,
  EngineSyncError,
  EngineTraversalError,
} from "../engine/errors.js";

describe("engine errors", () => {
  it("preserves tags, messages, and causes for every service error type", () => {
    const cause = new Error("root cause");
    const errors = [
      new EngineMutationError({ message: "mutation", cause }),
      new EngineRecalcError({ message: "recalc", cause }),
      new EngineSnapshotError({ message: "snapshot", cause }),
      new EngineSyncError({ message: "sync", cause }),
      new EngineHistoryError({ message: "history", cause }),
      new EnginePivotError({ message: "pivot", cause }),
      new EngineStructureError({ message: "structure", cause }),
      new EngineFormulaBindingError({ message: "binding", cause }),
      new EngineFormulaGraphError({ message: "graph", cause }),
      new EngineFormulaEvaluationError({ message: "evaluation", cause }),
      new EngineCellStateError({ message: "cell-state", cause }),
      new EngineTraversalError({ message: "traversal", cause }),
      new EngineMaintenanceError({ message: "maintenance", cause }),
      new EngineRuntimeScratchError({ message: "runtime-scratch", cause }),
    ];

    expect(errors.map((error) => error._tag)).toEqual([
      "EngineMutationError",
      "EngineRecalcError",
      "EngineSnapshotError",
      "EngineSyncError",
      "EngineHistoryError",
      "EnginePivotError",
      "EngineStructureError",
      "EngineFormulaBindingError",
      "EngineFormulaGraphError",
      "EngineFormulaEvaluationError",
      "EngineCellStateError",
      "EngineTraversalError",
      "EngineMaintenanceError",
      "EngineRuntimeScratchError",
    ]);
    for (const error of errors) {
      expect(error.message.length).toBeGreaterThan(0);
      expect(error.cause).toBe(cause);
    }
  });
});
