import { describe, expect, it } from "vitest";
import { Float64Arena, Uint32Arena } from "../index.js";

describe("program arenas", () => {
  it("should export arenas from the package root", () => {
    // Arrange

    // Act
    const uint32Arena = new Uint32Arena();
    const float64Arena = new Float64Arena();

    // Assert
    expect(uint32Arena).toBeInstanceOf(Uint32Arena);
    expect(float64Arena).toBeInstanceOf(Float64Arena);
  });

  it("packs uint32 programs into reusable contiguous slices", () => {
    // Arrange
    const arena = new Uint32Arena();

    // Act
    const first = arena.append(Uint32Array.from([1, 2, 3]));
    const second = arena.append(Uint32Array.from([4, 5]));

    // Assert
    expect(first).toEqual({ offset: 0, length: 3 });
    expect(second).toEqual({ offset: 3, length: 2 });
    expect(arena.size).toBe(5);
    expect(Array.from(arena.view())).toEqual([1, 2, 3, 4, 5]);
  });

  it("grows and resets uint32 storage without leaking stale data into the view", () => {
    // Arrange
    const arena = new Uint32Arena();

    // Act
    arena.append(Uint32Array.from({ length: 80 }, (_, index) => index + 1));

    // Assert
    expect(arena.size).toBe(80);
    expect(arena.view()[79]).toBe(80);

    // Act
    arena.reset();

    // Assert
    expect(arena.size).toBe(0);
    expect(Array.from(arena.view())).toEqual([]);

    // Act
    const slice = arena.append(Uint32Array.from([9, 10]));

    // Assert
    expect(slice).toEqual({ offset: 0, length: 2 });
    expect(Array.from(arena.view())).toEqual([9, 10]);
  });

  it("packs float constant pools with stable offsets across growth", () => {
    // Arrange
    const arena = new Float64Arena();

    // Act
    const first = arena.append([1.25, 2.5]);
    const second = arena.append(Array.from({ length: 70 }, (_, index) => index + 0.5));

    // Assert
    expect(first).toEqual({ offset: 0, length: 2 });
    expect(second).toEqual({ offset: 2, length: 70 });
    expect(arena.size).toBe(72);
    expect(Array.from(arena.view().slice(0, 4))).toEqual([1.25, 2.5, 0.5, 1.5]);

    // Act
    arena.reset();
    const rebuilt = arena.append([42]);

    // Assert
    expect(rebuilt).toEqual({ offset: 0, length: 1 });
    expect(Array.from(arena.view())).toEqual([42]);
  });
});
