import { describe, expect, it } from "vitest";
import { ErrorCode, ValueTag } from "@bilig/protocol";
import { CellFlags, CellStore } from "../cell-store.js";

describe("CellStore", () => {
  it("allocates materialized empty cells and grows capacity", () => {
    const store = new CellStore(1);
    const first = store.allocate(2, 4, 6);
    const second = store.allocate(2, 5, 7);

    expect(first).toBe(0);
    expect(second).toBe(1);
    expect(store.capacity).toBeGreaterThanOrEqual(2);
    expect(store.size).toBe(2);
    expect(store.sheetIds[1]).toBe(2);
    expect(store.rows[1]).toBe(5);
    expect(store.cols[1]).toBe(7);
    expect(store.tags[0]).toBe(ValueTag.Empty);
    expect(store.errors[0]).toBe(ErrorCode.None);
    expect(store.flags[0]).toBe(CellFlags.Materialized);
    expect(store.cycleGroupIds[1]).toBe(-1);
  });

  it("roundtrips values and reset clears all runtime arrays", () => {
    const store = new CellStore(2);
    const numberIndex = store.allocate(1, 0, 0);
    const stringIndex = store.allocate(1, 0, 1);
    const errorIndex = store.allocate(1, 0, 2);

    store.setValue(numberIndex, { tag: ValueTag.Boolean, value: true });
    store.setValue(stringIndex, { tag: ValueTag.String, value: "alpha", stringId: 4 }, 4);
    store.setValue(errorIndex, { tag: ValueTag.Error, code: ErrorCode.Ref });

    expect(store.getValue(numberIndex, () => "")).toEqual({ tag: ValueTag.Boolean, value: true });
    expect(store.getValue(stringIndex, (id) => `string-${id}`)).toEqual({
      tag: ValueTag.String,
      value: "string-4",
      stringId: 4
    });
    expect(store.getValue(errorIndex, () => "")).toEqual({ tag: ValueTag.Error, code: ErrorCode.Ref });
    expect(store.versions[numberIndex]).toBe(1);
    expect(store.versions[stringIndex]).toBe(1);
    expect(store.versions[errorIndex]).toBe(1);

    store.reset();

    expect(store.size).toBe(0);
    expect([...store.tags.slice(0, store.capacity)]).toEqual(new Array(store.capacity).fill(0));
    expect([...store.stringIds.slice(0, store.capacity)]).toEqual(new Array(store.capacity).fill(0));
    expect([...store.cycleGroupIds.slice(0, store.capacity)]).toEqual(new Array(store.capacity).fill(-1));
  });
});
