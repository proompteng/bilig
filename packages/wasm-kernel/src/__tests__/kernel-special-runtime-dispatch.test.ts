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

function packConstants(groups: number[][]): {
  constants: Float64Array;
  offsets: Uint32Array;
  lengths: Uint32Array;
} {
  const values: number[] = [];
  const offsets: number[] = [];
  const lengths: number[] = [];
  let offset = 0;

  for (const group of groups) {
    offsets.push(offset);
    lengths.push(group.length);
    values.push(...group);
    offset += group.length;
  }

  return {
    constants: Float64Array.from(values),
    offsets: Uint32Array.from(offsets),
    lengths: Uint32Array.from(lengths),
  };
}

function cellIndex(row: number, col: number, width: number): number {
  return row * width + col;
}

describe("wasm kernel special runtime dispatch", () => {
  it("keeps TODAY, NOW, and RAND on the current volatile path", async () => {
    const kernel = await createKernel();
    kernel.init(4, 4, 1, 1, 1);
    kernel.writeCells(
      new Uint8Array(4),
      new Float64Array(4),
      new Uint32Array(4),
      new Uint16Array(4),
    );
    kernel.uploadPrograms(
      new Uint32Array([
        encodeCall(BuiltinId.Today, 0),
        encodeRet(),
        encodeCall(BuiltinId.Now, 0),
        encodeRet(),
        encodeCall(BuiltinId.Rand, 0),
        encodeRet(),
      ]),
      new Uint32Array([0, 2, 4]),
      new Uint32Array([2, 2, 2]),
      new Uint32Array([0, 1, 2]),
    );
    kernel.uploadConstants(
      new Float64Array(),
      new Uint32Array([0, 0, 0]),
      new Uint32Array([0, 0, 0]),
    );
    kernel.uploadVolatileNowSerial(46100.65659722222);
    kernel.uploadVolatileRandomValues(new Float64Array([0.625]));

    kernel.evalBatch(Uint32Array.from([0, 1, 2]));

    expect(kernel.readTags()[0]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[0]).toBe(46100);
    expect(kernel.readTags()[1]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[1]).toBeCloseTo(46100.65659722222, 12);
    expect(kernel.readTags()[2]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[2]).toBe(0.625);
  });

  it("keeps SUMPRODUCT on the current scalar output path", async () => {
    const kernel = await createKernel();
    const width = 8;
    kernel.init(24, 4, 1, 1, 1);
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
      ]),
      new Float64Array([1, 2, 3, 4, 5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(24),
      new Uint16Array(24),
    );
    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 2, 3, 4, 5]),
      Uint32Array.from([0, 3]),
      Uint32Array.from([3, 3]),
    );
    kernel.uploadRangeShapes(Uint32Array.from([3, 3]), Uint32Array.from([1, 1]));
    kernel.uploadPrograms(
      new Uint32Array([
        encodePushRange(0),
        encodePushRange(1),
        encodeCall(BuiltinId.Sumproduct, 2),
        encodeRet(),
      ]),
      new Uint32Array([0]),
      new Uint32Array([4]),
      Uint32Array.from([cellIndex(1, 0, width)]),
    );
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0]), new Uint32Array([0]));

    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width)]));

    expect(kernel.readTags()[cellIndex(1, 0, width)]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBe(32);
  });

  it("keeps IRR, MIRR, XNPV, and XIRR on the current finance root path", async () => {
    const kernel = await createKernel();
    const width = 8;
    const tags = new Uint8Array(40);
    const numbers = new Float64Array(40);
    for (let index = 0; index < 22; index += 1) {
      tags[index] = ValueTag.Number;
    }
    numbers.set(
      [
        -70000, 12000, 15000, 18000, 21000, 26000, -120000, 39000, 30000, 21000, 37000, 46000,
        -10000, 2750, 4250, 3250, 2750, 39448, 39508, 39751, 39859, 39904,
      ],
      0,
    );
    kernel.init(40, 4, 3, 4, 22);
    kernel.writeCells(tags, numbers, new Uint32Array(40), new Uint16Array(40));
    kernel.uploadRangeMembers(
      Uint32Array.from([
        0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21,
      ]),
      Uint32Array.from([0, 6, 12, 17]),
      Uint32Array.from([6, 6, 5, 5]),
    );
    kernel.uploadRangeShapes(Uint32Array.from([6, 6, 5, 5]), Uint32Array.from([1, 1, 1, 1]));

    const packed = packPrograms([
      [encodePushRange(0), encodeCall(BuiltinId.Irr, 1), encodeRet()],
      [
        encodePushRange(1),
        encodePushNumber(0),
        encodePushNumber(1),
        encodeCall(BuiltinId.Mirr, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushRange(2),
        encodePushRange(3),
        encodeCall(BuiltinId.Xnpv, 3),
        encodeRet(),
      ],
      [encodePushRange(2), encodePushRange(3), encodeCall(BuiltinId.Xirr, 2), encodeRet()],
    ]);
    const constants = packConstants([[], [0.1, 0.12], [0.09], []]);
    const outputCells = Uint32Array.from([
      cellIndex(3, 0, width),
      cellIndex(3, 1, width),
      cellIndex(3, 2, width),
      cellIndex(3, 3, width),
    ]);
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells);
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths);

    kernel.evalBatch(outputCells);

    expect(kernel.readTags()[cellIndex(3, 0, width)]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[cellIndex(3, 0, width)]).toBeCloseTo(0.08663094803653162, 12);
    expect(kernel.readTags()[cellIndex(3, 1, width)]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[cellIndex(3, 1, width)]).toBeCloseTo(0.1260941303659051, 12);
    expect(kernel.readTags()[cellIndex(3, 2, width)]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[cellIndex(3, 2, width)]).toBeCloseTo(2086.647602031535, 9);
    expect(kernel.readTags()[cellIndex(3, 3, width)]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[cellIndex(3, 3, width)]).toBeCloseTo(0.37336253351883136, 12);
  });
});
