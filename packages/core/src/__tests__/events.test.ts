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
      compileMs: 0,
    },
  };
}

describe("EngineEventBus", () => {
  it("tracks listener presence, handles skipped address resolution, and grows watcher ids", () => {
    const events = new EngineEventBus();
    const general = vi.fn();
    const indexed = vi.fn();
    const addressed = vi.fn();

    expect(events.hasListeners()).toBe(false);
    expect(events.hasCellListeners()).toBe(false);
    expect(events.hasAddressListeners()).toBe(false);

    const unsubscribeGeneral = events.subscribe(general);
    const unsubscribeIndex = events.subscribeCellIndex(2, indexed);
    const unsubscribeAddress = events.subscribeCellAddress("Sheet1!A1", addressed);

    expect(events.hasListeners()).toBe(true);
    expect(events.hasCellListeners()).toBe(true);
    expect(events.hasAddressListeners()).toBe(true);

    events.emit(batchEvent(), new Uint32Array());
    expect(general).toHaveBeenCalledTimes(1);
    expect(indexed).not.toHaveBeenCalled();
    expect(addressed).not.toHaveBeenCalled();

    events.emit(batchEvent(new Uint32Array([2, 3])), new Uint32Array([2, 3]), (cellIndex) =>
      cellIndex === 2 ? "" : "!deleted",
    );
    expect(indexed).toHaveBeenCalledTimes(1);
    expect(addressed).not.toHaveBeenCalled();

    unsubscribeIndex();
    unsubscribeAddress();
    unsubscribeGeneral();

    expect(events.hasListeners()).toBe(false);
    expect(events.hasCellListeners()).toBe(false);
    expect(events.hasAddressListeners()).toBe(false);

    const growing = new EngineEventBus();
    const listeners = Array.from({ length: 80 }, () => vi.fn());
    listeners.forEach((listener, index) => {
      growing.subscribeCellIndex(index, listener);
    });
    growing.emitAllWatched(batchEvent());
    listeners.forEach((listener) => {
      expect(listener).toHaveBeenCalledTimes(1);
    });
  });

  it("deduplicates repeated listener registrations across changed indices and addresses", () => {
    const events = new EngineEventBus();
    const listener = vi.fn();

    const unsubscribe = events.subscribeCells([4, 9], ["Sheet1!A1"], listener);
    events.emit(batchEvent(new Uint32Array([4, 9])), new Uint32Array([4, 9]), (cellIndex) =>
      cellIndex === 4 ? "Sheet1!A1" : "Sheet1!B1",
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
