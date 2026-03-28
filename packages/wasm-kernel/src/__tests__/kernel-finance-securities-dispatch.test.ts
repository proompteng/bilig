import { describe, expect, it } from "vitest";
import { BuiltinId, Opcode, ValueTag } from "@bilig/protocol";
import { createKernel } from "../index.js";

function encodeCall(builtinId: number, argc: number): number {
  return (Opcode.CallBuiltin << 24) | ((builtinId << 8) | argc);
}

function encodePushNumber(constantIndex: number): number {
  return (Opcode.PushNumber << 24) | constantIndex;
}

function encodeRet(): number {
  return Opcode.Ret << 24;
}

function packPrograms(programs: number[][]): {
  programs: Uint32Array;
  offsets: Uint32Array;
  lengths: Uint32Array;
} {
  const flat: number[] = [];
  const offsets: number[] = [];
  const lengths: number[] = [];
  let offset = 0;

  for (const program of programs) {
    offsets.push(offset);
    lengths.push(program.length);
    flat.push(...program);
    offset += program.length;
  }

  return {
    programs: Uint32Array.from(flat),
    offsets: Uint32Array.from(offsets),
    lengths: Uint32Array.from(lengths),
  };
}

function packConstants(constantsByProgram: number[][]): {
  constants: Float64Array;
  offsets: Uint32Array;
  lengths: Uint32Array;
} {
  const flat: number[] = [];
  const offsets: number[] = [];
  const lengths: number[] = [];
  let offset = 0;

  for (const constants of constantsByProgram) {
    offsets.push(offset);
    lengths.push(constants.length);
    flat.push(...constants);
    offset += constants.length;
  }

  return {
    constants: Float64Array.from(flat),
    offsets: Uint32Array.from(offsets),
    lengths: Uint32Array.from(lengths),
  };
}

function cellIndex(row: number, col: number, width: number): number {
  return row * width + col;
}

function expectNumberCell(
  kernel: Awaited<ReturnType<typeof createKernel>>,
  index: number,
  expected: number,
  digits = 12,
): void {
  expect(kernel.readTags()[index]).toBe(ValueTag.Number);
  expect(kernel.readNumbers()[index]).toBeCloseTo(expected, digits);
}

describe("wasm kernel finance/security dispatch", () => {
  it("keeps discounted security formulas stable across refactors", async () => {
    const kernel = await createKernel();
    const width = 8;
    kernel.init(24, 8, 6, 2, 1);
    kernel.writeCells(
      new Uint8Array(24),
      new Float64Array(24),
      new Uint32Array(24),
      new Uint16Array(24),
    );

    const packed = packPrograms([
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(0),
        encodePushNumber(3),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodePushNumber(6),
        encodeCall(BuiltinId.Disc, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(0),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(5),
        encodePushNumber(6),
        encodePushNumber(7),
        encodeCall(BuiltinId.Pricedisc, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(0),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(5),
        encodePushNumber(6),
        encodePushNumber(7),
        encodeCall(BuiltinId.Yielddisc, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(0),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(5),
        encodeCall(BuiltinId.Tbillprice, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(0),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(5),
        encodeCall(BuiltinId.Tbillyield, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(0),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BuiltinId.Date, 3),
        encodePushNumber(5),
        encodeCall(BuiltinId.Tbilleq, 3),
        encodeRet(),
      ],
    ]);
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
      ]),
    );
    const constants = packConstants([
      [2023, 1, 1, 4, 97, 100, 2],
      [2008, 2, 16, 3, 1, 0.0525, 100, 2, 99.795],
      [2008, 2, 16, 3, 1, 99.795, 100, 2],
      [2008, 3, 31, 6, 1, 0.09],
      [2008, 3, 31, 6, 1, 98.45],
      [2008, 3, 31, 6, 1, 0.0914],
    ]);
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths);
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
      ]),
    );

    expectNumberCell(kernel, cellIndex(1, 0, width), 0.12);
    expectNumberCell(kernel, cellIndex(1, 1, width), 99.79583333333333);
    expectNumberCell(kernel, cellIndex(1, 2, width), 0.05282257198685834);
    expectNumberCell(kernel, cellIndex(1, 3, width), 98.45);
    expectNumberCell(kernel, cellIndex(1, 4, width), 0.09141696292534264);
    expectNumberCell(kernel, cellIndex(1, 5, width), 0.09415149356594302);
  });

  it("keeps coupon and duration formulas stable across refactors", async () => {
    const kernel = await createKernel();
    const width = 10;
    kernel.init(30, 10, 8, 2, 1);
    kernel.writeCells(
      new Uint8Array(30),
      new Float64Array(30),
      new Uint32Array(30),
      new Uint16Array(30),
    );

    const packed = packPrograms([
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodeCall(BuiltinId.Coupdaybs, 4),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodeCall(BuiltinId.Coupdays, 4),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BuiltinId.Duration, 6),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BuiltinId.Mduration, 6),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodePushNumber(6),
        encodeCall(BuiltinId.Price, 7),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodePushNumber(6),
        encodeCall(BuiltinId.Yield, 7),
        encodeRet(),
      ],
    ]);
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
      ]),
    );
    const constants = packConstants([
      [39107, 40132, 2, 4],
      [39107, 40132, 2, 4],
      [43282, 54058, 0.08, 0.09, 2, 1],
      [39448, 42370, 0.08, 0.09, 2, 1],
      [39493, 43054, 0.0575, 0.065, 100, 2, 0],
      [39493, 42689, 0.0575, 95.04287, 100, 2, 0],
    ]);
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths);
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
      ]),
    );

    expectNumberCell(kernel, cellIndex(1, 0, width), 70);
    expectNumberCell(kernel, cellIndex(1, 1, width), 180);
    expectNumberCell(kernel, cellIndex(1, 2, width), 10.919145281591925);
    expectNumberCell(kernel, cellIndex(1, 3, width), 5.735669813918838);
    expectNumberCell(kernel, cellIndex(1, 4, width), 94.63436162132213);
    expectNumberCell(kernel, cellIndex(1, 5, width), 0.065, 7);
  });

  it("keeps intrate and received stable across refactors", async () => {
    const kernel = await createKernel();
    const width = 8;
    kernel.init(16, 6, 2, 2, 1);
    kernel.writeCells(
      new Uint8Array(16),
      new Float64Array(16),
      new Uint32Array(16),
      new Uint16Array(16),
    );

    const packed = packPrograms([
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BuiltinId.Intrate, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BuiltinId.Received, 5),
        encodeRet(),
      ],
    ]);
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width)]),
    );
    const constants = packConstants([
      [44927, 45017, 1000, 1030, 2],
      [44927, 45017, 1000, 0.12, 2],
    ]);
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths);
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width)]));

    expectNumberCell(kernel, cellIndex(1, 0, width), 0.12);
    expectNumberCell(kernel, cellIndex(1, 1, width), 1030.9278350515465);
  });
});
