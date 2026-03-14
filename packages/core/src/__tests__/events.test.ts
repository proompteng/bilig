import { describe, expect, it, vi } from "vitest";
import type { EngineEvent } from "@bilig/protocol";
import { EngineEventBus } from "../events.js";

function batchEvent(changedCellIndices: Uint32Array = new Uint32Array()): EngineEvent {
  return {
    kind: "batch",
    changedCellIndices,
    metrics: {
      batchId: 1,
      changedInputCount: 0,
      dirtyFormulaCount: 0,
      wasmFormulaCount: 0,
      jsFormulaCount: 0,
      rangeNodeVisits: 0,
      recalcMs: 0,
      compileMs: 0
    }
  };
}

describe("EngineEventBus", () => {
  it("deduplicates repeated listener registrations across changed indices and addresses", () => {
    const events = new EngineEventBus();
    const listener = vi.fn();

    const unsubscribe = events.subscribeCells([4, 9], ["Sheet1!A1"], listener);
    events.emit(batchEvent(new Uint32Array([4, 9])), new Uint32Array([4, 9]), (cellIndex) =>
      cellIndex === 4 ? "Sheet1!A1" : "Sheet1!B1"
    );

    expect(listener).toHaveBeenCalledTimes(1);

    unsubscribe();
    events.emit(batchEvent(new Uint32Array([4])), new Uint32Array([4]), () => "Sheet1!A1");
    expect(listener).toHaveBeenCalledTimes(1);
  });

  it("notifies every watched listener once during broadcast invalidation", () => {
    const events = new EngineEventBus();
    const listener = vi.fn();

    events.subscribeCellIndex(2, listener);
    events.subscribeCellAddress("Sheet1!C3", listener);

    events.emitAllWatched(batchEvent());

    expect(listener).toHaveBeenCalledTimes(1);
  });
});
