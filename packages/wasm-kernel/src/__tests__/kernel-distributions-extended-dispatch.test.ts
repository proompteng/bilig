import { describe, expect, it } from "vitest";
import { BuiltinId, Opcode, ValueTag } from "@bilig/protocol";
import { createKernel } from "../index.js";

function encodeCall(builtinId: number, argc: number): number {
  return (Opcode.CallBuiltin << 24) | ((builtinId << 8) | argc);
}

function encodePushBoolean(value: boolean): number {
  return (Opcode.PushBoolean << 24) | (value ? 1 : 0);
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

describe("wasm kernel extended distribution dispatch", () => {
  it("keeps exponential, chi-square, beta/f, t, and discrete distribution dispatch stable", async () => {
    const kernel = await createKernel();
    const width = 12;
    kernel.init(96, 1, 39, 1, 1);
    kernel.writeCells(
      new Uint8Array(96),
      new Float64Array(96),
      new Uint32Array(96),
      new Uint16Array(96),
    );

    const packed = packPrograms([
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushBoolean(false),
        encodeCall(BuiltinId.Expondist, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushBoolean(true),
        encodeCall(BuiltinId.ExponDist, 3),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Chidist, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.ChisqDistRt, 2), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushBoolean(true),
        encodeCall(BuiltinId.ChisqDist, 3),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Chiinv, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.ChisqInvRt, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.ChisqInv, 2), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(true),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BuiltinId.BetaDist, 6),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodeCall(BuiltinId.BetaInv, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(true),
        encodeCall(BuiltinId.FDist, 4),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.FDistRt, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.FInv, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.FInvRt, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushBoolean(true),
        encodeCall(BuiltinId.TDist, 3),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.TDistRt, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.TDist2T, 2), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Tdist, 3),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.TInv, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.TInv2T, 2), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodeCall(BuiltinId.BinomDistRange, 4),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.BinomDistRange, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Critbinom, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodeCall(BuiltinId.Hypgeomdist, 4),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushBoolean(true),
        encodeCall(BuiltinId.HypgeomDist, 5),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Negbinomdist, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(true),
        encodeCall(BuiltinId.NegbinomDist, 4),
        encodeRet(),
      ],
    ]);
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(
        Array.from({ length: 27 }, (_, index) =>
          cellIndex(1 + Math.floor(index / width), index % width, width),
        ),
      ),
    );
    const constants = packConstants([
      [1, 2],
      [1, 2],
      [18.307, 10],
      [18.307, 10],
      [18.307, 10],
      [0.050001, 10],
      [0.050001, 10],
      [0.93, 1],
      [2, 8, 10, 1, 3],
      [0.6854705810117458, 8, 10, 1, 3],
      [15.2068649, 6, 4],
      [15.2068649, 6, 4],
      [0.01, 6, 4],
      [0.01, 6, 4],
      [1, 1],
      [1, 1],
      [1, 1],
      [1, 1, 2],
      [0.75, 1],
      [0.5, 1],
      [6, 0.5, 2, 4],
      [6, 0.5, 2],
      [6, 0.5, 0.7],
      [1, 4, 3, 10],
      [1, 4, 3, 10],
      [2, 3, 0.5],
      [2, 3, 0.5],
    ]);
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths);
    const outputCells = Uint32Array.from(
      Array.from({ length: 27 }, (_, index) =>
        cellIndex(1 + Math.floor(index / width), index % width, width),
      ),
    );
    kernel.evalBatch(outputCells);

    expectNumberCell(kernel, outputCells[0], 0.2706705664732254, 12);
    expectNumberCell(kernel, outputCells[1], 0.8646647167633873, 12);
    expectNumberCell(kernel, outputCells[2], 0.0500006, 6);
    expectNumberCell(kernel, outputCells[3], 0.0500006, 6);
    expectNumberCell(kernel, outputCells[4], 0.9499994, 6);
    expectNumberCell(kernel, outputCells[5], 18.306973, 6);
    expectNumberCell(kernel, outputCells[6], 18.306973, 6);
    expectNumberCell(kernel, outputCells[7], 3.2830202867594993, 12);
    expectNumberCell(kernel, outputCells[8], 0.6854705810117458, 9);
    expectNumberCell(kernel, outputCells[9], 2, 9);
    expectNumberCell(kernel, outputCells[10], 0.99, 9);
    expectNumberCell(kernel, outputCells[11], 0.01, 9);
    expectNumberCell(kernel, outputCells[12], 0.10930991466299911, 8);
    expectNumberCell(kernel, outputCells[13], 15.206864870947697, 7);
    expectNumberCell(kernel, outputCells[14], 0.75, 12);
    expectNumberCell(kernel, outputCells[15], 0.25, 12);
    expectNumberCell(kernel, outputCells[16], 0.5, 12);
    expectNumberCell(kernel, outputCells[17], 0.5, 12);
    expectNumberCell(kernel, outputCells[18], 1, 9);
    expectNumberCell(kernel, outputCells[19], 1, 9);
    expectNumberCell(kernel, outputCells[20], 0.78125, 12);
    expectNumberCell(kernel, outputCells[21], 0.234375, 12);
    expectNumberCell(kernel, outputCells[22], 4, 12);
    expectNumberCell(kernel, outputCells[23], 0.5, 12);
    expectNumberCell(kernel, outputCells[24], 2 / 3, 12);
    expectNumberCell(kernel, outputCells[25], 0.1875, 12);
    expectNumberCell(kernel, outputCells[26], 0.5, 12);
  });
});
