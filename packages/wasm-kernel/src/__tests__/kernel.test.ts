import { describe, expect, it } from "vitest";
import { createKernel } from "../index.js";

describe("wasm kernel", () => {
  it("evaluates a simple program batch", async () => {
    const kernel = await createKernel();
    kernel.init(4, 4, 4);
    kernel.writeCells(new Uint8Array([1, 0, 0, 0]), new Float64Array([10, 0, 0, 0]), new Uint16Array(4));
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
});
