import { describe, expect, it } from "vitest";
import { createKernel } from "../index.js";

describe("wasm kernel", () => {
  it("evaluates a simple program batch", async () => {
    const kernel = await createKernel();
    kernel.init(4, 4, 4, 4, 4);
    kernel.writeCells(new Uint8Array([1, 0, 0, 0]), new Float64Array([10, 0, 0, 0]), new Uint32Array(4), new Uint16Array(4));
    kernel.uploadPrograms(
      new Uint32Array([
        (3 << 24) | 0,
        (1 << 24) | 0,
        7 << 24,
        255 << 24
      ]),
      new Uint32Array([0]),
      new Uint32Array([4]),
      new Uint32Array([1])
    );
    kernel.uploadConstants(new Float64Array([2]), new Uint32Array([0]), new Uint32Array([1]));
    kernel.evalBatch(new Uint32Array([1]));
    expect(kernel.readNumbers()[1]).toBe(20);
  });

  it("evaluates aggregate and numeric builtins", async () => {
    const kernel = await createKernel();
    kernel.init(6, 6, 2, 6, 6);
    kernel.writeCells(
      new Uint8Array([1, 1, 0, 0, 0, 0]),
      new Float64Array([2, 3, 0, 0, 0, 0]),
      new Uint32Array(6),
      new Uint16Array(6)
    );
    kernel.uploadPrograms(
      new Uint32Array([
        (3 << 24) | 0,
        (3 << 24) | 1,
        (20 << 24) | (1 << 8) | 2,
        (1 << 24) | 0,
        5 << 24,
        255 << 24
      ]),
      new Uint32Array([0]),
      new Uint32Array([6]),
      new Uint32Array([2])
    );
    kernel.uploadConstants(new Float64Array([4]), new Uint32Array([0]), new Uint32Array([1]));

    kernel.evalBatch(new Uint32Array([2]));
    expect(kernel.readNumbers()[2]).toBe(9);
  });

  it("evaluates branch programs with jump opcodes", async () => {
    const kernel = await createKernel();
    kernel.init(4, 4, 4, 4, 4);
    kernel.writeCells(
      new Uint8Array([2, 0, 0, 0]),
      new Float64Array([1, 0, 0, 0]),
      new Uint32Array(4),
      new Uint16Array(4)
    );
    kernel.uploadPrograms(
      new Uint32Array([
        (3 << 24) | 0,
        (19 << 24) | 4,
        (1 << 24) | 0,
        (18 << 24) | 5,
        (1 << 24) | 1,
        255 << 24
      ]),
      new Uint32Array([0]),
      new Uint32Array([6]),
      new Uint32Array([1])
    );
    kernel.uploadConstants(new Float64Array([10, 20]), new Uint32Array([0]), new Uint32Array([2]));

    kernel.evalBatch(new Uint32Array([1]));
    expect(kernel.readNumbers()[1]).toBe(10);

    kernel.writeCells(
      new Uint8Array([2, kernel.readTags()[1]!, 0, 0]),
      new Float64Array([0, kernel.readNumbers()[1]!, 0, 0]),
      new Uint32Array(4),
      new Uint16Array([0, kernel.readErrors()[1]!, 0, 0])
    );
    kernel.evalBatch(new Uint32Array([1]));
    expect(kernel.readNumbers()[1]).toBe(20);
  });

  it("evaluates aggregate builtins through uploaded range members", async () => {
    const kernel = await createKernel();
    kernel.init(6, 6, 1, 4, 4);
    kernel.writeCells(
      new Uint8Array([1, 1, 0, 0, 0, 0]),
      new Float64Array([2, 3, 0, 0, 0, 0]),
      new Uint32Array(6),
      new Uint16Array(6)
    );
    kernel.uploadPrograms(
      new Uint32Array([
        (4 << 24) | 0,
        (20 << 24) | (1 << 8) | 1,
        255 << 24
      ]),
      new Uint32Array([0]),
      new Uint32Array([3]),
      new Uint32Array([2])
    );
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0]), new Uint32Array([0]));
    kernel.uploadRangeMembers(new Uint32Array([0, 1]), new Uint32Array([0]), new Uint32Array([2]));

    kernel.evalBatch(new Uint32Array([2]));

    expect(kernel.readNumbers()[2]).toBe(5);
    expect(kernel.readRangeLengths()[0]).toBe(2);
    expect(kernel.readRangeMembers()[1]).toBe(1);
  });
});
