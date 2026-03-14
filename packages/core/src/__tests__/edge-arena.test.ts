import { describe, expect, it } from "vitest";
import { EdgeArena } from "../edge-arena.js";

describe("EdgeArena", () => {
  it("allocates, replaces, appends uniquely, removes, and reuses freed slices", () => {
    const arena = new EdgeArena();
    const initial = arena.alloc(2);
    const filled = arena.replace(initial, Uint32Array.from([3, 7]));
    expect([...arena.read(filled)]).toEqual([3, 7]);

    const unique = arena.appendUnique(filled, 7);
    expect(unique).toEqual(filled);
    expect([...arena.read(unique)]).toEqual([3, 7]);

    const expanded = arena.appendUnique(unique, 9);
    expect([...arena.read(expanded)]).toEqual([3, 7, 9]);

    const removed = arena.removeValue(expanded, 7);
    expect([...arena.read(removed)]).toEqual([3, 9]);

    arena.free(removed);
    const reused = arena.alloc(1);
    expect([filled.ptr, removed.ptr]).toContain(reused.ptr);
  });

  it("returns empty slices for zero-length operations and resets cleanly", () => {
    const arena = new EdgeArena();
    const empty = arena.empty();

    expect(arena.alloc(0)).toEqual(empty);
    expect([...arena.read(empty)]).toEqual([]);
    expect(arena.replace(empty, [])).toEqual(empty);
    expect(arena.removeValue(empty, 1)).toEqual(empty);

    const slice = arena.replace(empty, Uint32Array.from([1, 2]));
    expect([...arena.read(slice)]).toEqual([1, 2]);

    arena.reset();

    const afterReset = arena.alloc(1);
    expect(afterReset.ptr).toBe(0);
    expect(afterReset.len).toBe(0);
  });

  it("exposes zero-copy read views for hot iteration paths", () => {
    const arena = new EdgeArena();
    const slice = arena.replace(arena.empty(), Uint32Array.from([4, 8, 15]));

    const view = arena.readView(slice);
    const poolView = arena.view();

    expect([...view]).toEqual([4, 8, 15]);
    expect([...poolView]).toEqual([4, 8, 15]);

    view[1] = 16;
    expect([...arena.read(slice)]).toEqual([4, 16, 15]);
  });
});
