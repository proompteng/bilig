import { describe, expect, it } from "vitest";
import { BuiltinId, Opcode, ValueTag } from "@bilig/protocol";
import { createKernel } from "../index.js";

function encodeCall(builtinId: number, argc: number): number {
  return (Opcode.CallBuiltin << 24) | ((builtinId << 8) | argc);
}

function encodePushNumber(constantIndex: number): number {
  return (Opcode.PushNumber << 24) | constantIndex;
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

describe("wasm kernel cashflow helpers", () => {
  it("keeps annuity helper behavior stable across refactors", async () => {
    const kernel = await createKernel();
    const width = 11;
    kernel.init(22, 1, 0, 11, 28);
    kernel.writeCells(
      new Uint8Array(20),
      new Float64Array(20),
      new Uint32Array(20),
      new Uint16Array(20),
    );

    const packed = packPrograms([
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Pv, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Pmt, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Nper, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Rate, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodeCall(BuiltinId.Ipmt, 4),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodeCall(BuiltinId.Ppmt, 4),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodeCall(BuiltinId.Ispmt, 4),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BuiltinId.Cumipmt, 6),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodePushNumber(4),
        encodePushNumber(5),
        encodeCall(BuiltinId.Cumprinc, 6),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodeCall(BuiltinId.Fv, 4),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodeCall(BuiltinId.Npv, 4),
        encodeRet(),
      ],
    ]);
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(0, 0, width),
        cellIndex(0, 1, width),
        cellIndex(0, 2, width),
        cellIndex(0, 3, width),
        cellIndex(0, 4, width),
        cellIndex(0, 5, width),
        cellIndex(0, 6, width),
        cellIndex(0, 7, width),
        cellIndex(0, 8, width),
        cellIndex(0, 9, width),
        cellIndex(0, 10, width),
      ]),
    );
    const constants = packConstants([
      [0.1, 2, -576.1904761904761],
      [0.1, 2, 1000],
      [0.1, -576.1904761904761, 1000],
      [48, -200, 8000],
      [0.1, 1, 2, 1000],
      [0.1, 1, 2, 1000],
      [0.1, 1, 2, 1000],
      [0.09 / 12, 30 * 12, 125000, 13, 24, 0],
      [0.09 / 12, 30 * 12, 125000, 13, 24, 0],
      [0.1, 2, -100, -1000],
      [0.1, 100, 200, 300],
    ]);
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths);
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(0, 0, width),
        cellIndex(0, 1, width),
        cellIndex(0, 2, width),
        cellIndex(0, 3, width),
        cellIndex(0, 4, width),
        cellIndex(0, 5, width),
        cellIndex(0, 6, width),
        cellIndex(0, 7, width),
        cellIndex(0, 8, width),
        cellIndex(0, 9, width),
        cellIndex(0, 10, width),
      ]),
    );

    expectNumberCell(kernel, cellIndex(0, 0, width), 1000.0000000000006);
    expectNumberCell(kernel, cellIndex(0, 1, width), -576.1904761904758);
    expectNumberCell(kernel, cellIndex(0, 2, width), 1.9999999999999982);
    expectNumberCell(kernel, cellIndex(0, 3, width), 0.007701472488246008, 12);
    expectNumberCell(kernel, cellIndex(0, 4, width), -100);
    expectNumberCell(kernel, cellIndex(0, 5, width), -476.1904761904758);
    expectNumberCell(kernel, cellIndex(0, 7, width), -11135.232130750845, 9);
    expectNumberCell(kernel, cellIndex(0, 8, width), -934.1071234208765, 9);
    expectNumberCell(kernel, cellIndex(0, 9, width), 1420);
  });

  it("keeps irregular cashflow helper behavior stable across refactors", async () => {
    const kernel = await createKernel();
    const width = 10;
    kernel.init(40, 4, 3, 3, 22);
    const cellTags = new Uint8Array(40);
    const cellNumbers = new Float64Array(40);
    const cashflows = [-120000, 39000, 30000, 21000, 37000, 46000];
    const xValues = [-10000, 2750, 4250, 3250, 2750];
    const xDates = [39448, 39508, 39751, 39859, 39904];

    cashflows.forEach((value, index) => {
      cellTags[index] = ValueTag.Number;
      cellNumbers[index] = value;
    });
    xValues.forEach((value, index) => {
      cellTags[10 + index] = ValueTag.Number;
      cellNumbers[10 + index] = value;
    });
    xDates.forEach((value, index) => {
      cellTags[20 + index] = ValueTag.Number;
      cellNumbers[20 + index] = value;
    });

    kernel.writeCells(cellTags, cellNumbers, new Uint32Array(40), new Uint16Array(40));
    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 2, 3, 4, 5, 10, 11, 12, 13, 14, 20, 21, 22, 23, 24]),
      Uint32Array.from([0, 6, 11]),
      Uint32Array.from([6, 5, 5]),
    );
    kernel.uploadRangeShapes(Uint32Array.from([6, 5, 5]), Uint32Array.from([1, 1, 1]));

    const packed = packPrograms([
      [
        encodePushRange(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodeCall(BuiltinId.Mirr, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushRange(1),
        encodePushRange(2),
        encodeCall(BuiltinId.Xnpv, 3),
        encodeRet(),
      ],
      [encodePushRange(1), encodePushRange(2), encodeCall(BuiltinId.Xirr, 2), encodeRet()],
    ]);
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(3, 0, width), cellIndex(3, 1, width), cellIndex(3, 2, width)]),
    );
    const constants = packConstants([[0.1, 0.12], [0.09], []]);
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths);
    kernel.evalBatch(
      Uint32Array.from([cellIndex(3, 0, width), cellIndex(3, 1, width), cellIndex(3, 2, width)]),
    );

    expectNumberCell(kernel, cellIndex(3, 0, width), 0.1260941303659051, 12);
    expectNumberCell(kernel, cellIndex(3, 1, width), 2086.647602031535, 9);
    expectNumberCell(kernel, cellIndex(3, 2, width), 0.37336253351883136, 12);
  });
});
