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

describe("wasm kernel aggregate and criteria dispatch", () => {
  it("reuses repeated SUM range aggregates within a batch and refreshes them across batches", async () => {
    const kernel = await createKernel();
    const width = 4;
    kernel.init(16, 2, 1, 1, 3);

    const cellTags = new Uint8Array(16);
    const cellNumbers = new Float64Array(16);
    const cellStringIds = new Uint32Array(16);
    const cellErrors = new Uint16Array(16);

    cellTags[0] = ValueTag.Number;
    cellNumbers[0] = 1;
    cellTags[1] = ValueTag.Number;
    cellNumbers[1] = 2;
    cellTags[2] = ValueTag.Number;
    cellNumbers[2] = 3;
    kernel.writeCells(cellTags, cellNumbers, cellStringIds, cellErrors);
    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 2]),
      Uint32Array.from([0]),
      Uint32Array.from([3]),
    );
    kernel.uploadRangeShapes(Uint32Array.from([3]), Uint32Array.from([1]));

    const packed = packPrograms([
      [encodePushRange(0), encodeCall(BuiltinId.Sum, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BuiltinId.Sum, 1), encodeRet()],
    ]);
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width)]),
    );
    const constants = packConstants([[], []]);
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths);

    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width)]));
    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBe(6);
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(6);

    cellNumbers[0] = 10;
    kernel.writeCells(cellTags, cellNumbers, cellStringIds, cellErrors);
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width)]));
    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBe(15);
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(15);
  });

  it("keeps aggregate and criteria builtins stable across refactors", async () => {
    const kernel = await createKernel();
    const width = 32;
    const strings = packStrings(["x", "A", "B", ">2", "Z"]);
    kernel.init(128, 5, 8, 8, 8);
    kernel.uploadStrings(strings.offsets, strings.lengths, strings.data);

    const cellTags = new Uint8Array(128);
    const cellNumbers = new Float64Array(128);
    const cellStringIds = new Uint32Array(128);
    const cellErrors = new Uint16Array(128);

    cellTags[0] = ValueTag.Number;
    cellNumbers[0] = 1;
    cellTags[1] = ValueTag.Empty;
    cellTags[2] = ValueTag.Number;
    cellNumbers[2] = 3;
    cellTags[3] = ValueTag.String;
    cellStringIds[3] = 0;
    cellTags[4] = ValueTag.Number;
    cellNumbers[4] = 5;
    cellTags[5] = ValueTag.Number;
    cellNumbers[5] = 7;

    cellTags[6] = ValueTag.String;
    cellStringIds[6] = 1;
    cellTags[7] = ValueTag.String;
    cellStringIds[7] = 2;
    cellTags[8] = ValueTag.String;
    cellStringIds[8] = 1;
    cellTags[9] = ValueTag.String;
    cellStringIds[9] = 2;
    cellTags[10] = ValueTag.String;
    cellStringIds[10] = 1;
    cellTags[11] = ValueTag.Empty;

    const numericCells = [
      [12, 10],
      [13, 20],
      [14, 30],
      [15, 40],
      [16, 50],
      [17, 60],
      [18, 24],
      [19, 36],
      [20, 6],
      [21, 8],
      [22, 14],
      [23, 1],
      [24, 3],
      [25, 9],
      [26, 1],
      [27, 2],
      [28, 4],
      [29, 1],
      [30, 3],
      [31, 5],
    ] as const;
    for (const [index, value] of numericCells) {
      cellTags[index] = ValueTag.Number;
      cellNumbers[index] = value;
    }

    kernel.writeCells(cellTags, cellNumbers, cellStringIds, cellErrors);
    kernel.uploadRangeMembers(
      Uint32Array.from([
        0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24,
        25, 26, 27, 28, 29, 30, 31,
      ]),
      Uint32Array.from([0, 6, 12, 18, 20, 23, 26, 29]),
      Uint32Array.from([6, 6, 6, 2, 3, 3, 3, 3]),
    );
    kernel.uploadRangeShapes(
      Uint32Array.from([6, 6, 6, 2, 3, 3, 3, 3]),
      Uint32Array.from([1, 1, 1, 1, 1, 1, 1, 1]),
    );

    const packed = packPrograms([
      [encodePushRange(0), encodePushNumber(0), encodeCall(BuiltinId.Sum, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BuiltinId.Avg, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BuiltinId.Min, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BuiltinId.Max, 2), encodeRet()],
      [
        encodePushRange(0),
        encodePushNumber(0),
        encodePushString(0),
        encodeCall(BuiltinId.Count, 3),
        encodeRet(),
      ],
      [encodePushRange(0), encodeCall(BuiltinId.CountA, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BuiltinId.Countblank, 1), encodeRet()],
      [encodePushRange(3), encodeCall(BuiltinId.Gcd, 1), encodeRet()],
      [encodePushRange(4), encodeCall(BuiltinId.Lcm, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BuiltinId.Product, 1), encodeRet()],
      [encodePushRange(5), encodeCall(BuiltinId.Geomean, 1), encodeRet()],
      [encodePushRange(6), encodeCall(BuiltinId.Harmean, 1), encodeRet()],
      [encodePushRange(7), encodeCall(BuiltinId.Sumsq, 1), encodeRet()],
      [encodePushRange(1), encodePushString(1), encodeCall(BuiltinId.Countif, 2), encodeRet()],
      [
        encodePushRange(1),
        encodePushString(1),
        encodePushRange(0),
        encodePushString(3),
        encodeCall(BuiltinId.Countifs, 4),
        encodeRet(),
      ],
      [
        encodePushRange(1),
        encodePushString(1),
        encodePushRange(2),
        encodeCall(BuiltinId.Sumif, 3),
        encodeRet(),
      ],
      [
        encodePushRange(2),
        encodePushRange(1),
        encodePushString(1),
        encodePushRange(0),
        encodePushString(3),
        encodeCall(BuiltinId.Sumifs, 5),
        encodeRet(),
      ],
      [
        encodePushRange(1),
        encodePushString(1),
        encodePushRange(2),
        encodeCall(BuiltinId.Averageif, 3),
        encodeRet(),
      ],
      [
        encodePushRange(2),
        encodePushRange(1),
        encodePushString(1),
        encodePushRange(0),
        encodePushString(3),
        encodeCall(BuiltinId.Averageifs, 5),
        encodeRet(),
      ],
      [
        encodePushRange(2),
        encodePushRange(1),
        encodePushString(1),
        encodePushRange(0),
        encodePushString(3),
        encodeCall(BuiltinId.Minifs, 5),
        encodeRet(),
      ],
      [
        encodePushRange(2),
        encodePushRange(1),
        encodePushString(1),
        encodePushRange(0),
        encodePushString(3),
        encodeCall(BuiltinId.Maxifs, 5),
        encodeRet(),
      ],
      [
        encodePushRange(2),
        encodePushRange(1),
        encodePushString(4),
        encodePushRange(0),
        encodePushString(3),
        encodeCall(BuiltinId.Averageifs, 5),
        encodeRet(),
      ],
    ]);
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 22 }, (_, index) => cellIndex(1, index, width))),
    );

    const constants = packConstants([
      [2],
      [2],
      [-4],
      [9],
      [9],
      [],
      [],
      [],
      [],
      [],
      [],
      [],
      [],
      [],
      [],
      [],
      [],
      [],
      [],
      [],
      [],
      [],
    ]);
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths);
    kernel.evalBatch(
      Uint32Array.from(Array.from({ length: 22 }, (_, index) => cellIndex(1, index, width))),
    );

    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBe(18);
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBeCloseTo(3, 12);
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(-4);
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(9);
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(6);
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(5);
    expect(kernel.readNumbers()[cellIndex(1, 6, width)]).toBe(1);
    expect(kernel.readNumbers()[cellIndex(1, 7, width)]).toBe(12);
    expect(kernel.readNumbers()[cellIndex(1, 8, width)]).toBe(168);
    expect(kernel.readNumbers()[cellIndex(1, 9, width)]).toBe(0);
    expect(kernel.readNumbers()[cellIndex(1, 10, width)]).toBe(3);
    expect(kernel.readNumbers()[cellIndex(1, 11, width)]).toBeCloseTo(12 / 7, 12);
    expect(kernel.readNumbers()[cellIndex(1, 12, width)]).toBe(35);
    expect(kernel.readNumbers()[cellIndex(1, 13, width)]).toBe(3);
    expect(kernel.readNumbers()[cellIndex(1, 14, width)]).toBe(2);
    expect(kernel.readNumbers()[cellIndex(1, 15, width)]).toBe(90);
    expect(kernel.readNumbers()[cellIndex(1, 16, width)]).toBe(80);
    expect(kernel.readNumbers()[cellIndex(1, 17, width)]).toBe(30);
    expect(kernel.readNumbers()[cellIndex(1, 18, width)]).toBe(40);
    expect(kernel.readNumbers()[cellIndex(1, 19, width)]).toBe(30);
    expect(kernel.readNumbers()[cellIndex(1, 20, width)]).toBe(50);
    expect(kernel.readTags()[cellIndex(1, 21, width)]).toBe(ValueTag.Error);
    expect(kernel.readErrors()[cellIndex(1, 21, width)]).toBe(ErrorCode.Div0);
  });
});
