import { describe, expect, it } from "vitest";
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
        end: { row: 1, col: 1, text: "B2" }
      },
      {
        ensureCell: () => nextCellIndex++,
        forEachSheetCell: () => {}
      }
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
        end: { row: 1, text: "2" }
      },
      {
        ensureCell: () => {
          throw new Error("rows ranges should use sheet iteration callbacks");
        },
        forEachSheetCell: (_sheetId, fn) => {
          fn(4, 1, 0);
        }
      }
    );

    const before = registry.getDescriptor(registered.rangeIndex);
    expect(before.membersLength).toBe(1);
    expect(registry.getMembers(registered.rangeIndex)).toEqual(Uint32Array.from([4]));

    expect(registry.addDynamicMember(3, 1, 5, 9)).toEqual([registered.rangeIndex]);

    const after = registry.getDescriptor(registered.rangeIndex);
    expect(after.membersOffset).toBeGreaterThanOrEqual(before.membersOffset);
    expect(after.membersLength).toBe(2);
    expect(registry.getMembers(registered.rangeIndex)).toEqual(Uint32Array.from([4, 9]));
    expect(registry.getMemberPoolView().subarray(after.membersOffset, after.membersOffset + after.membersLength)).toEqual(
      Uint32Array.from([4, 9])
    );
  });

  it("clears descriptor slices when the last reference is released", () => {
    const registry = new RangeRegistry();
    const registered = registry.intern(
      1,
      {
        kind: "cells",
        start: { row: 0, col: 0, text: "A1" },
        end: { row: 0, col: 1, text: "B1" }
      },
      {
        ensureCell: ((_sheetId, _row, col) => col + 20) as (sheetId: number, row: number, col: number) => number,
        forEachSheetCell: () => {}
      }
    );

    const released = registry.release(registered.rangeIndex);
    const descriptor = registry.getDescriptor(registered.rangeIndex);

    expect(released.removed).toBe(true);
    expect(released.members).toEqual(Uint32Array.from([20, 21]));
    expect(descriptor.membersOffset).toBe(0);
    expect(descriptor.membersLength).toBe(0);
    expect(descriptor.refCount).toBe(0);
  });
});
