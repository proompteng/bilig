import { describe, expect, it } from "vitest";
import { ErrorCode, Opcode, ValueTag } from "@bilig/protocol";
import { createKernel } from "../index.js";

const BUILTIN = {
  ISBLANK: 18,
  ISNUMBER: 19,
  ISTEXT: 20,
  DATE: 21,
  YEAR: 22,
  MONTH: 23,
  DAY: 24,
  EDATE: 25,
  EOMONTH: 26
} as const;

function encodeCall(builtinId: number, argc: number): number {
  return (Opcode.CallBuiltin << 24) | ((builtinId << 8) | argc);
}

function encodePushCell(cellIndex: number): number {
  return (Opcode.PushCell << 24) | cellIndex;
}

function encodePushRange(rangeIndex: number): number {
  return (Opcode.PushRange << 24) | rangeIndex;
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
    lengths: Uint32Array.from(lengths)
  };
}

function cellIndex(row: number, col: number, width: number): number {
  return row * width + col;
}

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
    expect(kernel.readConstantOffsets()[0]).toBe(0);
    expect(kernel.readConstantLengths()[0]).toBe(1);
    expect(kernel.readConstants()[0]).toBe(2);
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

  it("evaluates exact-safe logical info builtins with zero-arg, scalar, and range cases", async () => {
    const kernel = await createKernel();
    const width = 8;
    kernel.init(16, 8, 2, 2, 2);
    kernel.writeCells(
      new Uint8Array([0, 1, 2, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([0, 42, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array([0, 0, 0, 7, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(16)
    );
    kernel.uploadRangeMembers(new Uint32Array([0, 1]), new Uint32Array([0]), new Uint32Array([2]));

    const packed = packPrograms([
      [encodeCall(BUILTIN.ISBLANK, 0), encodeRet()],
      [encodeCall(BUILTIN.ISNUMBER, 0), encodeRet()],
      [encodeCall(BUILTIN.ISTEXT, 0), encodeRet()],
      [encodePushCell(0), encodeCall(BUILTIN.ISBLANK, 1), encodeRet()],
      [encodePushCell(1), encodeCall(BUILTIN.ISNUMBER, 1), encodeRet()],
      [encodePushCell(3), encodeCall(BUILTIN.ISTEXT, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.ISNUMBER, 1), encodeRet()]
    ]);

    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
        cellIndex(1, 7, width)
      ])
    );
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0]), new Uint32Array([0]));
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
        cellIndex(1, 7, width)
      ])
    );

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Boolean);
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(1);
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.Boolean);
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(0);
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.Boolean);
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(0);
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.Boolean);
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(1);
    expect(kernel.readTags()[cellIndex(1, 5, width)]).toBe(ValueTag.Boolean);
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(1);
    expect(kernel.readTags()[cellIndex(1, 6, width)]).toBe(ValueTag.Boolean);
    expect(kernel.readNumbers()[cellIndex(1, 6, width)]).toBe(1);
    expect(kernel.readTags()[cellIndex(1, 7, width)]).toBe(ValueTag.Error);
    expect(kernel.readErrors()[cellIndex(1, 7, width)]).toBe(ErrorCode.Value);
  });

  it("evaluates exact-safe date builtins with Excel coercion and errors", async () => {
    const kernel = await createKernel();
    const width = 10;
    kernel.init(20, 10, 5, 2, 2);
    kernel.writeCells(
      new Uint8Array([3, 2, 4, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([0, 1, 0, 45351, 45351.75, 60, 45322, 45337, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array([1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array([0, 0, ErrorCode.Ref, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0])
    );

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.DATE, 3), encodeRet()],
      [encodePushCell(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BUILTIN.DATE, 3), encodeRet()],
      [encodePushCell(3), encodeCall(BUILTIN.YEAR, 1), encodeRet()],
      [encodePushCell(4), encodeCall(BUILTIN.MONTH, 1), encodeRet()],
      [encodePushCell(5), encodeCall(BUILTIN.DAY, 1), encodeRet()],
      [encodePushCell(6), encodePushNumber(3), encodeCall(BUILTIN.EDATE, 2), encodeRet()],
      [encodePushCell(0), encodePushNumber(4), encodeCall(BUILTIN.EDATE, 2), encodeRet()],
      [encodePushCell(7), encodePushCell(1), encodeCall(BUILTIN.EOMONTH, 2), encodeRet()],
      [encodePushCell(2), encodePushNumber(4), encodeCall(BUILTIN.EOMONTH, 2), encodeRet()]
    ]);

    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
        cellIndex(1, 7, width),
        cellIndex(1, 8, width),
        cellIndex(1, 9, width)
      ])
    );
    kernel.uploadConstants(new Float64Array([2024, 2, 29, 1.9, 1]), new Uint32Array([0]), new Uint32Array([5]));
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
        cellIndex(1, 7, width),
        cellIndex(1, 8, width),
        cellIndex(1, 9, width)
      ])
    );

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(45351);
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.Error);
    expect(kernel.readErrors()[cellIndex(1, 2, width)]).toBe(ErrorCode.Value);
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(2024);
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(2);
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(29);
    expect(kernel.readNumbers()[cellIndex(1, 6, width)]).toBe(45351);
    expect(kernel.readTags()[cellIndex(1, 7, width)]).toBe(ValueTag.Error);
    expect(kernel.readErrors()[cellIndex(1, 7, width)]).toBe(ErrorCode.Value);
    expect(kernel.readNumbers()[cellIndex(1, 8, width)]).toBe(45382);
    expect(kernel.readTags()[cellIndex(1, 9, width)]).toBe(ValueTag.Error);
    expect(kernel.readErrors()[cellIndex(1, 9, width)]).toBe(ErrorCode.Ref);
  });

  it("evaluates logical and rounding builtins with parity-safe scalar semantics", async () => {
    const kernel = await createKernel();
    kernel.init(8, 8, 4, 4, 4);
    kernel.writeCells(
      new Uint8Array([1, 1, 4, 0, 0, 0, 0, 0]),
      new Float64Array([123.4, 1, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(8),
      new Uint16Array([0, 0, ErrorCode.Value, 0, 0, 0, 0, 0])
    );
    kernel.uploadPrograms(
      new Uint32Array([
        (3 << 24) | 0,
        (1 << 24) | 0,
        (20 << 24) | (8 << 8) | 2,
        255 << 24,

        (3 << 24) | 1,
        (2 << 24) | 1,
        (20 << 24) | (9 << 8) | 2,
        255 << 24,

        (3 << 24) | 1,
        (20 << 24) | (15 << 8) | 1,
        255 << 24,

        (3 << 24) | 2,
        (2 << 24) | 2,
        (20 << 24) | (13 << 8) | 2,
        255 << 24
      ]),
      new Uint32Array([0, 4, 8, 11]),
      new Uint32Array([4, 4, 3, 4]),
      new Uint32Array([3, 4, 5, 6])
    );
    kernel.uploadConstants(
      new Float64Array([-1, 0.5, 1]),
      new Uint32Array([0, 0, 0, 0]),
      new Uint32Array([2, 2, 0, 1])
    );

    kernel.evalBatch(new Uint32Array([3, 4, 5, 6]));

    expect(kernel.readNumbers()[3]).toBe(120);
    expect(kernel.readNumbers()[4]).toBe(1);
    expect(kernel.readTags()[5]).toBe(ValueTag.Boolean);
    expect(kernel.readNumbers()[5]).toBe(0);
    expect(kernel.readTags()[6]).toBe(ValueTag.Error);
    expect(kernel.readErrors()[6]).toBe(ErrorCode.Value);
  });
});
