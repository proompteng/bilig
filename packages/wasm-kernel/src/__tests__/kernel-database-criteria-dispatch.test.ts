import { describe, expect, it } from "vitest";
import { BuiltinId, ErrorCode, Opcode, ValueTag } from "@bilig/protocol";
import { createKernel } from "../index.js";

function encodeCall(builtinId: number, argc: number): number {
  return (Opcode.CallBuiltin << 24) | ((builtinId << 8) | argc);
}

function encodePushNumber(constantIndex: number): number {
  return (Opcode.PushNumber << 24) | constantIndex;
}

function encodePushString(stringId: number): number {
  return (Opcode.PushString << 24) | stringId;
}

function encodePushRange(rangeIndex: number): number {
  return (Opcode.PushRange << 24) | rangeIndex;
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

function packStrings(strings: string[]): {
  offsets: Uint32Array;
  lengths: Uint32Array;
  data: Uint16Array;
} {
  const offsets: number[] = [];
  const lengths: number[] = [];
  const data: number[] = [];
  let offset = 0;
  for (const text of strings) {
    offsets.push(offset);
    lengths.push(text.length);
    for (let index = 0; index < text.length; index += 1) {
      data.push(text.charCodeAt(index));
    }
    offset += text.length;
  }
  return {
    offsets: Uint32Array.from(offsets),
    lengths: Uint32Array.from(lengths),
    data: Uint16Array.from(data),
  };
}

function cellIndex(row: number, col: number, width: number): number {
  return row * width + col;
}

describe("wasm kernel database criteria dispatch", () => {
  it("keeps database builtins stable across refactors", async () => {
    const kernel = await createKernel();
    const width = 32;
    const strings = packStrings([
      "Tree",
      "Height",
      "Age",
      "Yield",
      "Profit",
      "Apple",
      "Pear",
      "Cherry",
    ]);

    kernel.init(128, 8, 8, 3, 8);
    kernel.uploadStrings(strings.offsets, strings.lengths, strings.data);

    const cellTags = new Uint8Array(128);
    const cellNumbers = new Float64Array(128);
    const cellStringIds = new Uint32Array(128);
    const cellErrors = new Uint16Array(128);

    const setString = (index: number, stringId: number) => {
      cellTags[index] = ValueTag.String;
      cellStringIds[index] = stringId;
    };
    const setNumber = (index: number, value: number) => {
      cellTags[index] = ValueTag.Number;
      cellNumbers[index] = value;
    };

    setString(0, 0);
    setString(1, 1);
    setString(2, 2);
    setString(3, 3);
    setString(4, 4);

    setString(5, 5);
    setNumber(6, 18);
    setNumber(7, 20);
    setNumber(8, 14);
    setNumber(9, 105);

    setString(10, 6);
    setNumber(11, 12);
    setNumber(12, 12);
    setNumber(13, 10);
    setNumber(14, 96);

    setString(15, 7);
    setNumber(16, 13);
    setNumber(17, 14);
    setNumber(18, 9);
    setNumber(19, 105);

    setString(20, 5);
    setNumber(21, 14);
    setNumber(22, 15);
    setNumber(23, 10);
    setNumber(24, 75);

    setString(25, 6);
    setNumber(26, 9);
    setNumber(27, 8);
    setNumber(28, 8);
    setNumber(29, 77);

    setString(30, 0);
    setString(31, 5);

    kernel.writeCells(cellTags, cellNumbers, cellStringIds, cellErrors);
    kernel.uploadRangeMembers(
      Uint32Array.from([
        0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24,
        25, 26, 27, 28, 29, 30, 31, 32,
      ]),
      Uint32Array.from([0, 30, 32]),
      Uint32Array.from([30, 2, 1]),
    );
    kernel.uploadRangeShapes(Uint32Array.from([6, 2, 1]), Uint32Array.from([5, 1, 1]));

    const packed = packPrograms([
      [
        encodePushRange(0),
        encodePushString(1),
        encodePushRange(1),
        encodeCall(BuiltinId.Dcount, 3),
        encodeRet(),
      ],
      [
        encodePushRange(0),
        encodePushString(0),
        encodePushRange(1),
        encodeCall(BuiltinId.Dcounta, 3),
        encodeRet(),
      ],
      [
        encodePushRange(0),
        encodePushRange(2),
        encodePushRange(1),
        encodeCall(BuiltinId.Dcount, 3),
        encodeRet(),
      ],
      [
        encodePushRange(0),
        encodePushString(2),
        encodePushRange(1),
        encodeCall(BuiltinId.Dget, 3),
        encodeRet(),
      ],
      [
        encodePushRange(0),
        encodePushString(3),
        encodePushRange(1),
        encodeCall(BuiltinId.Daverage, 3),
        encodeRet(),
      ],
      [
        encodePushRange(0),
        encodePushString(4),
        encodePushRange(1),
        encodeCall(BuiltinId.Dmax, 3),
        encodeRet(),
      ],
      [
        encodePushRange(0),
        encodePushNumber(0),
        encodePushRange(1),
        encodeCall(BuiltinId.Dmin, 3),
        encodeRet(),
      ],
      [
        encodePushRange(0),
        encodePushString(4),
        encodePushRange(1),
        encodeCall(BuiltinId.Dsum, 3),
        encodeRet(),
      ],
      [
        encodePushRange(0),
        encodePushString(2),
        encodePushRange(1),
        encodeCall(BuiltinId.Dproduct, 3),
        encodeRet(),
      ],
      [
        encodePushRange(0),
        encodePushString(1),
        encodePushRange(1),
        encodeCall(BuiltinId.Dstdev, 3),
        encodeRet(),
      ],
      [
        encodePushRange(0),
        encodePushString(1),
        encodePushRange(1),
        encodeCall(BuiltinId.Dstdevp, 3),
        encodeRet(),
      ],
      [
        encodePushRange(0),
        encodePushString(2),
        encodePushRange(1),
        encodeCall(BuiltinId.Dvar, 3),
        encodeRet(),
      ],
      [
        encodePushRange(0),
        encodePushString(2),
        encodePushRange(1),
        encodeCall(BuiltinId.Dvarp, 3),
        encodeRet(),
      ],
    ]);
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 13 }, (_, index) => cellIndex(1, index, width))),
    );
    const constants = packConstants([[], [], [], [], [], [], [2], [], [], [], [], [], []]);
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths);
    kernel.evalBatch(
      Uint32Array.from(Array.from({ length: 13 }, (_, index) => cellIndex(1, index, width))),
    );

    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBe(2);
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(2);
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(2);
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.Error);
    expect(kernel.readErrors()[cellIndex(1, 3, width)]).toBe(ErrorCode.Value);
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(12);
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(105);
    expect(kernel.readNumbers()[cellIndex(1, 6, width)]).toBe(14);
    expect(kernel.readNumbers()[cellIndex(1, 7, width)]).toBe(180);
    expect(kernel.readNumbers()[cellIndex(1, 8, width)]).toBe(300);
    expect(kernel.readNumbers()[cellIndex(1, 9, width)]).toBeCloseTo(Math.sqrt(8), 12);
    expect(kernel.readNumbers()[cellIndex(1, 10, width)]).toBe(2);
    expect(kernel.readNumbers()[cellIndex(1, 11, width)]).toBe(12.5);
    expect(kernel.readNumbers()[cellIndex(1, 12, width)]).toBe(6.25);
  });
});
