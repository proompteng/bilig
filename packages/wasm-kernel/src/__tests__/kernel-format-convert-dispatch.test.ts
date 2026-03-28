import { describe, expect, it } from "vitest";
import { BuiltinId, ErrorCode, Opcode, ValueTag } from "@bilig/protocol";
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

function encodePushString(stringId: number): number {
  return (Opcode.PushString << 24) | stringId;
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

describe("wasm kernel format and conversion dispatch", () => {
  it("keeps address, dollar, radix, and unit conversion dispatch stable across refactors", async () => {
    const kernel = await createKernel();
    const width = 12;
    kernel.init(96, 8, 13, 1, 1);
    kernel.uploadStrings(
      Uint32Array.from([0, 7, 11, 21, 23, 25, 26, 27, 30, 33]),
      Uint32Array.from([7, 4, 10, 2, 2, 1, 1, 3, 3, 3]),
      Uint16Array.from(
        Array.from("O'Brien00FF1111111111mikmFCDEMEURFRF", (char) => char.charCodeAt(0)),
      ),
    );
    kernel.writeCells(
      new Uint8Array(96),
      new Float64Array(96),
      new Uint32Array(96),
      new Uint16Array(96),
    );

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Address, 2), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushString(0),
        encodeCall(BuiltinId.Address, 5),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Dollar, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Dollarde, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Dollarfr, 2), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Base, 3),
        encodeRet(),
      ],
      [encodePushString(1), encodePushNumber(0), encodeCall(BuiltinId.Decimal, 2), encodeRet()],
      [encodePushString(2), encodeCall(BuiltinId.Bin2dec, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Dec2hex, 2), encodeRet()],
      [
        encodePushNumber(0),
        encodePushString(3),
        encodePushString(4),
        encodeCall(BuiltinId.Convert, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushString(5),
        encodePushString(6),
        encodeCall(BuiltinId.Convert, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushString(7),
        encodePushString(8),
        encodeCall(BuiltinId.Euroconvert, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushString(9),
        encodePushString(7),
        encodePushBoolean(true),
        encodePushNumber(1),
        encodeCall(BuiltinId.Euroconvert, 5),
        encodeRet(),
      ],
    ]);
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 13 }, (_, index) => cellIndex(1, index, width))),
    );
    const constants = packConstants([
      [12, 3],
      [2, 28, 3, 1],
      [-1234.5, 1],
      [1.08, 16],
      [1.5, 16],
      [255, 16, 4],
      [16],
      [],
      [255, 4],
      [6],
      [68],
      [1.2],
      [1, 3],
    ]);
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths);
    kernel.evalBatch(
      Uint32Array.from(Array.from({ length: 13 }, (_, index) => cellIndex(1, index, width))),
    );

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.String);
    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.String);
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.String);
    expectNumberCell(kernel, cellIndex(1, 3, width), 1.5);
    expectNumberCell(kernel, cellIndex(1, 4, width), 1.08);
    expect(kernel.readTags()[cellIndex(1, 5, width)]).toBe(ValueTag.String);
    expectNumberCell(kernel, cellIndex(1, 6, width), 255);
    expectNumberCell(kernel, cellIndex(1, 7, width), -1);
    expect(kernel.readTags()[cellIndex(1, 8, width)]).toBe(ValueTag.String);
    expectNumberCell(kernel, cellIndex(1, 9, width), 9.656064);
    expectNumberCell(kernel, cellIndex(1, 10, width), 20);
    expectNumberCell(kernel, cellIndex(1, 11, width), 0.61);
    expectNumberCell(kernel, cellIndex(1, 12, width), 0.29728616, 8);
    expect(kernel.readOutputStrings()).toEqual([
      "$C$12",
      "'O''Brien'!$AB2",
      "-$1,234.5",
      "00FF",
      "00FF",
    ]);
  });

  it("preserves error codes for invalid conversion inputs", async () => {
    const kernel = await createKernel();
    const width = 4;
    kernel.init(16, 3, 2, 1, 1);
    kernel.uploadStrings(
      Uint32Array.from([0, 2, 5]),
      Uint32Array.from([2, 3, 3]),
      Uint16Array.from(Array.from("ftsecBAD", (char) => char.charCodeAt(0))),
    );
    kernel.writeCells(
      new Uint8Array(16),
      new Float64Array(16),
      new Uint32Array(16),
      new Uint16Array(16),
    );

    const packed = packPrograms([
      [
        encodePushNumber(0),
        encodePushString(0),
        encodePushString(1),
        encodeCall(BuiltinId.Convert, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(1),
        encodePushString(2),
        encodePushString(1),
        encodeCall(BuiltinId.Euroconvert, 3),
        encodeRet(),
      ],
    ]);
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width)]),
    );
    const constants = packConstants([[2.5], [1]]);
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths);
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width)]));

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.Error);
    expect(kernel.readErrors()[cellIndex(1, 0, width)]).toBe(ErrorCode.NA);
    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Error);
    expect(kernel.readErrors()[cellIndex(1, 1, width)]).toBe(ErrorCode.Value);
  });
});
