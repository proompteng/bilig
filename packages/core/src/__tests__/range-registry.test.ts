import { describe, expect, it } from "vitest";
import { makeCellEntity, makeRangeEntity } from "../entity-ids.js";
import { RangeRegistry } from "../range-registry.js";

describe("RangeRegistry", () => {
  it("stores bounded range members in a shared pool with descriptor offsets", () => {
    const registry = new RangeRegistry();
    let nextCellIndex = 10;

    const registered = registry.intern(
      1,
      {
        kind: "cells",
        start: { row: 0, col: 0, text: "A1" },
        end: { row: 1, col: 1, text: "B2" },
      },
      {
        ensureCell: () => nextCellIndex++,
        forEachSheetCell: () => {},
      },
    );

    const descriptor = registry.getDescriptor(registered.rangeIndex);
    expect(descriptor.membersOffset).toBe(0);
    expect(descriptor.membersLength).toBe(4);
    expect(registry.getMembers(registered.rangeIndex)).toEqual(Uint32Array.from([10, 11, 12, 13]));
    expect(registry.getMemberPoolView()).toEqual(Uint32Array.from([10, 11, 12, 13]));
  });

  it("updates dynamic range member offsets and lengths when cells materialize later", () => {
    const registry = new RangeRegistry();
    const registered = registry.intern(
      3,
      {
        kind: "rows",
        start: { row: 1, text: "2" },
        end: { row: 1, text: "2" },
      },
      {
        ensureCell: () => {
          throw new Error("rows ranges should use sheet iteration callbacks");
        },
        forEachSheetCell: (_sheetId, fn) => {
          fn(4, 1, 0);
        },
      },
    );

    const before = registry.getDescriptor(registered.rangeIndex);
    expect(before.membersLength).toBe(1);
    expect(registry.getMembers(registered.rangeIndex)).toEqual(Uint32Array.from([4]));

    expect(registry.addDynamicMember(3, 1, 5, 9)).toEqual([registered.rangeIndex]);

    const after = registry.getDescriptor(registered.rangeIndex);
    expect(after.membersOffset).toBeGreaterThanOrEqual(before.membersOffset);
    expect(after.membersLength).toBe(2);
    expect(registry.getMembers(registered.rangeIndex)).toEqual(Uint32Array.from([4, 9]));
    expect(
      registry
        .getMemberPoolView()
        .subarray(after.membersOffset, after.membersOffset + after.membersLength),
    ).toEqual(Uint32Array.from([4, 9]));
  });

  it("clears descriptor slices when the last reference is released", () => {
    const registry = new RangeRegistry();
    const registered = registry.intern(
      1,
      {
        kind: "cells",
        start: { row: 0, col: 0, text: "A1" },
        end: { row: 0, col: 1, text: "B1" },
      },
      {
        ensureCell: ((_sheetId, _row, col) => col + 20) as (
          sheetId: number,
          row: number,
          col: number,
        ) => number,
        forEachSheetCell: () => {},
      },
    );

    const released = registry.release(registered.rangeIndex);
    const descriptor = registry.getDescriptor(registered.rangeIndex);

    expect(released.removed).toBe(true);
    expect(released.members).toEqual(Uint32Array.from([20, 21]));
    expect(descriptor.membersOffset).toBe(0);
    expect(descriptor.membersLength).toBe(0);
    expect(descriptor.refCount).toBe(0);
  });

  it("chains prefix ranges through prior range entities instead of all member cells", () => {
    const registry = new RangeRegistry();
    let nextCellIndex = 30;

    const first = registry.intern(
      1,
      {
        kind: "cells",
        start: { row: 0, col: 0, text: "A1" },
        end: { row: 1, col: 0, text: "A2" },
      },
      {
        ensureCell: () => nextCellIndex++,
        forEachSheetCell: () => {},
        isFormulaCell: (cellIndex: number) => cellIndex === 31 || cellIndex === 32,
      },
    );
    const second = registry.intern(
      1,
      {
        kind: "cells",
        start: { row: 0, col: 0, text: "A1" },
        end: { row: 2, col: 0, text: "A3" },
      },
      {
        ensureCell: () => nextCellIndex++,
        forEachSheetCell: () => {},
        isFormulaCell: (cellIndex: number) => cellIndex === 31 || cellIndex === 32,
      },
    );

    expect(registry.getMembers(second.rangeIndex)).toEqual(Uint32Array.from([30, 31, 32]));
    expect(registry.getFormulaMembers(second.rangeIndex)).toEqual(Uint32Array.from([31, 32]));
    expect(registry.getDependencySourceEntities(second.rangeIndex)).toEqual(
      Uint32Array.from([makeRangeEntity(first.rangeIndex), makeCellEntity(32)]),
    );

    const releasedFirst = registry.release(first.rangeIndex);
    expect(releasedFirst.removed).toBe(false);
    expect(registry.getDescriptor(first.rangeIndex).refCount).toBe(1);

    const releasedSecond = registry.release(second.rangeIndex);
    expect(releasedSecond.removed).toBe(true);
    expect(registry.getDescriptor(second.rangeIndex).refCount).toBe(0);

    const releasedParent = registry.release(first.rangeIndex);
    expect(releasedParent.removed).toBe(true);
    expect(registry.getDescriptor(first.rangeIndex).refCount).toBe(0);
  });

  it("refreshes range members and dependency sources after structural row remaps", () => {
    const registry = new RangeRegistry();
    let entries = [
      { cellIndex: 100, row: 0, col: 0 },
      { cellIndex: 101, row: 1, col: 0 },
      { cellIndex: 102, row: 2, col: 0 },
    ];
    const materializer = {
      ensureCell: (_sheetId: number, row: number, col: number) => {
        const entry = entries.find((current) => current.row === row && current.col === col);
        if (!entry) {
          throw new Error(`Missing materialized cell for ${row},${col}`);
        }
        return entry.cellIndex;
      },
      forEachSheetCell: (
        _sheetId: number,
        fn: (cellIndex: number, row: number, col: number) => void,
      ) => {
        entries.forEach(({ cellIndex, row, col }) => {
          fn(cellIndex, row, col);
        });
      },
      isFormulaCell: (cellIndex: number) => cellIndex === 101,
    };

    const prefix = registry.intern(
      1,
      {
        kind: "cells",
        start: { row: 0, col: 0, text: "A1" },
        end: { row: 1, col: 0, text: "A2" },
      },
      materializer,
    );
    const range = registry.intern(
      1,
      {
        kind: "cells",
        start: { row: 0, col: 0, text: "A1" },
        end: { row: 2, col: 0, text: "A3" },
      },
      materializer,
    );

    expect(registry.getMembers(range.rangeIndex)).toEqual(Uint32Array.from([100, 101, 102]));
    expect(registry.getDependencySourceEntities(range.rangeIndex)).toEqual(
      Uint32Array.from([makeRangeEntity(prefix.rangeIndex), makeCellEntity(102)]),
    );

    entries = [
      { cellIndex: 102, row: 0, col: 0 },
      { cellIndex: 100, row: 1, col: 0 },
      { cellIndex: 101, row: 2, col: 0 },
    ];

    registry.refresh(prefix.rangeIndex, materializer);
    registry.refresh(range.rangeIndex, materializer);

    expect(registry.getMembers(prefix.rangeIndex)).toEqual(Uint32Array.from([102, 100]));
    expect(registry.getFormulaMembers(prefix.rangeIndex)).toEqual(new Uint32Array());
    expect(registry.getMembers(range.rangeIndex)).toEqual(Uint32Array.from([102, 100, 101]));
    expect(registry.getFormulaMembers(range.rangeIndex)).toEqual(Uint32Array.from([101]));
    expect(registry.getDependencySourceEntities(range.rangeIndex)).toEqual(
      Uint32Array.from([makeRangeEntity(prefix.rangeIndex), makeCellEntity(101)]),
    );
  });
});
