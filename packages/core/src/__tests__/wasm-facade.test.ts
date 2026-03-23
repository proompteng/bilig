import { describe, expect, it } from "vitest";
import { BuiltinId, Opcode, ValueTag } from "@bilig/protocol";
import { CellStore } from "../cell-store.js";
import { StringPool } from "../string-pool.js";
import { WasmKernelFacade } from "../wasm-facade.js";

describe("WasmKernelFacade", () => {
  it("exposes refreshed runtime views after formula and range uploads", async () => {
    const facade = new WasmKernelFacade();
    await facade.init();

    facade.uploadFormulas({
      targets: new Uint32Array([2]),
      programs: new Uint32Array([(4 << 24) | 0, (20 << 24) | (1 << 8) | 1, 255 << 24]),
      programOffsets: new Uint32Array([0]),
      programLengths: new Uint32Array([3]),
      constants: new Float64Array(),
      constantOffsets: new Uint32Array([0]),
      constantLengths: new Uint32Array([0]),
    });

    facade.uploadRanges({
      members: new Uint32Array([0, 1]),
      offsets: new Uint32Array([0]),
      lengths: new Uint32Array([2]),
      rowCounts: new Uint32Array([2]),
      colCounts: new Uint32Array([1]),
    });

    expect(facade.programOffsets[0]).toBe(0);
    expect(facade.programLengths[0]).toBe(3);
    expect(facade.constantOffsets[0]).toBe(0);
    expect(facade.constantLengths[0]).toBe(0);
    expect(facade.rangeOffsets[0]).toBe(0);
    expect(facade.rangeLengths[0]).toBe(2);
    expect(Array.from(facade.rangeMembers.slice(0, 2))).toEqual([0, 1]);
  });

  it("syncs sparse cell updates into kernel-backed views and back to the store", async () => {
    const facade = new WasmKernelFacade();
    await facade.init();

    facade.uploadFormulas({
      targets: new Uint32Array([1]),
      programs: new Uint32Array([(3 << 24) | 0, (1 << 24) | 0, 7 << 24, 255 << 24]),
      programOffsets: new Uint32Array([0]),
      programLengths: new Uint32Array([4]),
      constants: new Float64Array([2]),
      constantOffsets: new Uint32Array([0]),
      constantLengths: new Uint32Array([1]),
    });

    const store = new CellStore();
    const sourceIndex = store.allocate(0, 0, 0);
    const targetIndex = store.allocate(0, 0, 1);
    store.setValue(sourceIndex, { tag: ValueTag.Number, value: 10 });

    facade.syncFromStore(store, Uint32Array.from([sourceIndex]));
    expect(facade.tags[sourceIndex]).toBe(ValueTag.Number);
    expect(facade.numbers[sourceIndex]).toBe(10);

    facade.evalBatch(new Uint32Array([targetIndex]));
    facade.syncToStore(store, new Uint32Array([targetIndex]), new StringPool());

    expect(store.numbers[targetIndex]).toBe(20);
    expect(store.versions[targetIndex]).toBe(1);
    expect(facade.constantOffsets[0]).toBe(0);
    expect(facade.constantLengths[0]).toBe(1);
    expect(facade.constants[0]).toBe(2);
  });
  it("uploads volatile random values for RAND evaluation", async () => {
    const facade = new WasmKernelFacade();
    await facade.init();

    facade.uploadFormulas({
      targets: new Uint32Array([0]),
      programs: new Uint32Array([
        (Opcode.CallBuiltin << 24) | ((BuiltinId.Rand << 8) | 0),
        Opcode.Ret << 24,
      ]),
      programOffsets: new Uint32Array([0]),
      programLengths: new Uint32Array([2]),
      constants: new Float64Array(),
      constantOffsets: new Uint32Array([0]),
      constantLengths: new Uint32Array([0]),
    });
    facade.uploadVolatileRandomValues(new Float64Array([0.625]));

    const store = new CellStore();
    const targetIndex = store.allocate(0, 0, 0);
    facade.syncFromStore(store, Uint32Array.from([targetIndex]));
    facade.evalBatch(new Uint32Array([targetIndex]));
    facade.syncToStore(store, new Uint32Array([targetIndex]), new StringPool());

    expect(store.getValue(targetIndex, () => undefined)).toEqual({
      tag: ValueTag.Number,
      value: 0.625,
    });
  });

  it("exposes numeric spill results for native SEQUENCE evaluation", async () => {
    const facade = new WasmKernelFacade();
    await facade.init();

    facade.uploadFormulas({
      targets: new Uint32Array([0]),
      programs: new Uint32Array([
        (Opcode.PushNumber << 24) | 0,
        (Opcode.PushNumber << 24) | 1,
        (Opcode.PushNumber << 24) | 2,
        (Opcode.PushNumber << 24) | 3,
        (Opcode.CallBuiltin << 24) | ((BuiltinId.Sequence << 8) | 4),
        Opcode.Ret << 24,
      ]),
      programOffsets: new Uint32Array([0]),
      programLengths: new Uint32Array([6]),
      constants: new Float64Array([3, 1, 1, 1]),
      constantOffsets: new Uint32Array([0]),
      constantLengths: new Uint32Array([4]),
    });

    const store = new CellStore();
    const targetIndex = store.allocate(0, 0, 0);
    facade.syncFromStore(store, Uint32Array.from([targetIndex]));
    facade.evalBatch(new Uint32Array([targetIndex]));
    facade.syncToStore(store, new Uint32Array([targetIndex]), new StringPool());

    expect(store.getValue(targetIndex, () => undefined)).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(facade.readNumericSpill(targetIndex)).toEqual({
      rows: 3,
      cols: 1,
      values: new Float64Array([1, 2, 3]),
    });
  });
});
