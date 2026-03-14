import { describe, expect, it } from "vitest";
import { FormulaMode, ValueTag } from "@bilig/protocol";
import { CellStore } from "../cell-store.js";
import { WasmKernelFacade } from "../wasm-facade.js";

describe("WasmKernelFacade", () => {
  it("exposes refreshed runtime views after formula and range uploads", async () => {
    const facade = new WasmKernelFacade();
    await facade.init();

    facade.uploadFormulas({
      targets: new Uint32Array([2]),
      modes: [FormulaMode.WasmFastPath],
      programs: new Uint32Array([(4 << 24) | 0, (20 << 24) | (1 << 8) | 1, 255 << 24]),
      programOffsets: new Uint32Array([0]),
      programLengths: new Uint32Array([3]),
      constants: new Float64Array(),
      constantOffsets: new Uint32Array([0]),
      constantLengths: new Uint32Array([0])
    });

    facade.uploadRanges({
      members: new Uint32Array([0, 1]),
      offsets: new Uint32Array([0]),
      lengths: new Uint32Array([2])
    });

    expect(facade.programOffsets[0]).toBe(0);
    expect(facade.programLengths[0]).toBe(3);
    expect(facade.rangeOffsets[0]).toBe(0);
    expect(facade.rangeLengths[0]).toBe(2);
    expect(Array.from(facade.rangeMembers.slice(0, 2))).toEqual([0, 1]);
  });

  it("syncs sparse cell updates into kernel-backed views and back to the store", async () => {
    const facade = new WasmKernelFacade();
    await facade.init();

    facade.uploadFormulas({
      targets: new Uint32Array([1]),
      modes: [FormulaMode.WasmFastPath],
      programs: new Uint32Array([(3 << 24) | 0, (1 << 24) | 0, (7 << 24), 255 << 24]),
      programOffsets: new Uint32Array([0]),
      programLengths: new Uint32Array([4]),
      constants: new Float64Array([2]),
      constantOffsets: new Uint32Array([0]),
      constantLengths: new Uint32Array([1])
    });

    const store = new CellStore();
    const sourceIndex = store.allocate(0, 0, 0);
    const targetIndex = store.allocate(0, 0, 1);
    store.setValue(sourceIndex, { tag: ValueTag.Number, value: 10 });

    facade.syncFromStore(store, Uint32Array.from([sourceIndex]));
    expect(facade.tags[sourceIndex]).toBe(ValueTag.Number);
    expect(facade.numbers[sourceIndex]).toBe(10);

    facade.evalBatch(new Uint32Array([targetIndex]));
    facade.syncToStore(store, new Uint32Array([targetIndex]));

    expect(store.numbers[targetIndex]).toBe(20);
    expect(store.versions[targetIndex]).toBe(1);
  });
});
