import { describe, expect, it } from "vitest";
import { BuiltinId, ErrorCode, Opcode, ValueTag } from "@bilig/protocol";
import { createKernel } from "../index.js";

const BUILTIN = {
  CONCAT: BuiltinId.Concat,
  LEN: BuiltinId.Len,
  ISBLANK: BuiltinId.IsBlank,
  ISNUMBER: BuiltinId.IsNumber,
  ISTEXT: BuiltinId.IsText,
  DATE: BuiltinId.Date,
  YEAR: BuiltinId.Year,
  MONTH: BuiltinId.Month,
  DAY: BuiltinId.Day,
  EDATE: BuiltinId.Edate,
  EOMONTH: BuiltinId.Eomonth,
  EXACT: BuiltinId.Exact,
  INT: BuiltinId.Int,
  ROUNDUP: BuiltinId.RoundUp,
  ROUNDDOWN: BuiltinId.RoundDown,
  TIME: BuiltinId.Time,
  HOUR: BuiltinId.Hour,
  MINUTE: BuiltinId.Minute,
  SECOND: BuiltinId.Second,
  WEEKDAY: BuiltinId.Weekday,
  DAYS: BuiltinId.Days,
  WORKDAY: BuiltinId.Workday,
  NETWORKDAYS: BuiltinId.Networkdays,
  WEEKNUM: BuiltinId.Weeknum,
  TODAY: BuiltinId.Today,
  NOW: BuiltinId.Now,
  RAND: BuiltinId.Rand,
  LEFT: BuiltinId.Left,
  RIGHT: BuiltinId.Right,
  MID: BuiltinId.Mid,
  TRIM: BuiltinId.Trim,
  UPPER: BuiltinId.Upper,
  LOWER: BuiltinId.Lower,
  REPLACE: BuiltinId.Replace,
  SUBSTITUTE: BuiltinId.Substitute,
  REPT: BuiltinId.Rept,
  FIND: BuiltinId.Find,
  SEARCH: BuiltinId.Search,
  VALUE: BuiltinId.Value,
  NA: BuiltinId.Na,
  IFERROR: BuiltinId.Iferror,
  IFNA: BuiltinId.Ifna,
  COUNTIF: BuiltinId.Countif,
  COUNTIFS: BuiltinId.Countifs,
  SUMIF: BuiltinId.Sumif,
  SUMIFS: BuiltinId.Sumifs,
  AVERAGEIF: BuiltinId.Averageif,
  AVERAGEIFS: BuiltinId.Averageifs,
  SUMPRODUCT: BuiltinId.Sumproduct,
  MATCH: BuiltinId.Match,
  INDEX: BuiltinId.Index,
  VLOOKUP: BuiltinId.Vlookup,
  HLOOKUP: BuiltinId.Hlookup,
  XMATCH: BuiltinId.Xmatch,
  XLOOKUP: BuiltinId.Xlookup,
  OFFSET: BuiltinId.Offset,
  TAKE: BuiltinId.Take,
  DROP: BuiltinId.Drop,
  CHOOSECOLS: BuiltinId.Choosecols,
  CHOOSEROWS: BuiltinId.Chooserows,
  SORT: BuiltinId.Sort,
  SORTBY: BuiltinId.Sortby,
  TOCOL: BuiltinId.Tocol,
  TOROW: BuiltinId.Torow,
  WRAPROWS: BuiltinId.Wraprows,
  WRAPCOLS: BuiltinId.Wrapcols,
} as const;

function asciiCodes(text: string): Uint16Array {
  return Uint16Array.from(Array.from(text, (char) => char.charCodeAt(0)));
}

function encodeCall(builtinId: number, argc: number): number {
  return (Opcode.CallBuiltin << 24) | ((builtinId << 8) | argc);
}

function encodePushCell(cellOffset: number): number {
  return (Opcode.PushCell << 24) | cellOffset;
}

function encodePushRange(rangeIndex: number): number {
  return (Opcode.PushRange << 24) | rangeIndex;
}

function encodePushNumber(constantIndex: number): number {
  return (Opcode.PushNumber << 24) | constantIndex;
}

function encodePushString(stringId: number): number {
  return (Opcode.PushString << 24) | stringId;
}

function encodeBinary(opcode: Opcode): number {
  return opcode << 24;
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

function cellIndex(row: number, col: number, width: number): number {
  return row * width + col;
}

describe("wasm kernel", () => {
  it("evaluates a simple program batch", async () => {
    const kernel = await createKernel();
    kernel.init(4, 4, 4, 4, 4);
    kernel.writeCells(
      new Uint8Array([1, 0, 0, 0]),
      new Float64Array([10, 0, 0, 0]),
      new Uint32Array(4),
      new Uint16Array(4),
    );
    kernel.uploadPrograms(
      new Uint32Array([(3 << 24) | 0, (1 << 24) | 0, 7 << 24, 255 << 24]),
      new Uint32Array([0]),
      new Uint32Array([4]),
      new Uint32Array([1]),
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
      new Uint16Array(6),
    );
    kernel.uploadPrograms(
      new Uint32Array([
        (3 << 24) | 0,
        (3 << 24) | 1,
        (20 << 24) | (1 << 8) | 2,
        (1 << 24) | 0,
        5 << 24,
        255 << 24,
      ]),
      new Uint32Array([0]),
      new Uint32Array([6]),
      new Uint32Array([2]),
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
      new Uint16Array(4),
    );
    kernel.uploadPrograms(
      new Uint32Array([
        (3 << 24) | 0,
        (19 << 24) | 4,
        (1 << 24) | 0,
        (18 << 24) | 5,
        (1 << 24) | 1,
        255 << 24,
      ]),
      new Uint32Array([0]),
      new Uint32Array([6]),
      new Uint32Array([1]),
    );
    kernel.uploadConstants(new Float64Array([10, 20]), new Uint32Array([0]), new Uint32Array([2]));

    kernel.evalBatch(new Uint32Array([1]));
    expect(kernel.readNumbers()[1]).toBe(10);

    const tags = kernel.readTags();
    const numbers = kernel.readNumbers();
    const errors = kernel.readErrors();
    kernel.writeCells(
      new Uint8Array([2, tags[1], 0, 0]),
      new Float64Array([0, numbers[1], 0, 0]),
      new Uint32Array(4),
      new Uint16Array([0, errors[1], 0, 0]),
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
      new Uint16Array(6),
    );
    kernel.uploadPrograms(
      new Uint32Array([(4 << 24) | 0, (20 << 24) | (1 << 8) | 1, 255 << 24]),
      new Uint32Array([0]),
      new Uint32Array([3]),
      new Uint32Array([2]),
    );
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0]), new Uint32Array([0]));
    kernel.uploadRangeMembers(new Uint32Array([0, 1]), new Uint32Array([0]), new Uint32Array([2]));
    kernel.uploadRangeShapes(new Uint32Array([2]), new Uint32Array([1]));

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
      new Uint16Array(16),
    );
    kernel.uploadRangeMembers(new Uint32Array([0, 1]), new Uint32Array([0]), new Uint32Array([2]));

    const packed = packPrograms([
      [encodeCall(BUILTIN.ISBLANK, 0), encodeRet()],
      [encodeCall(BUILTIN.ISNUMBER, 0), encodeRet()],
      [encodeCall(BUILTIN.ISTEXT, 0), encodeRet()],
      [encodePushCell(0), encodeCall(BUILTIN.ISBLANK, 1), encodeRet()],
      [encodePushCell(1), encodeCall(BUILTIN.ISNUMBER, 1), encodeRet()],
      [encodePushCell(3), encodeCall(BUILTIN.ISTEXT, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.ISNUMBER, 1), encodeRet()],
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
      ]),
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
        cellIndex(1, 7, width),
      ]),
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

  it("evaluates LEN with scalar coercion and range rejection", async () => {
    const kernel = await createKernel();
    const width = 8;
    kernel.init(16, 8, 1, 1, 2);
    kernel.uploadStringLengths(Uint32Array.from([0, 0, 0, 0, 0, 0, 0, 0, 0, 5]));
    kernel.writeCells(
      new Uint8Array([0, 2, 1, 3, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([0, 1, 123.45, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array([0, 0, 0, 9, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array([0, 0, 0, 0, ErrorCode.Ref, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
    );
    kernel.uploadRangeMembers(new Uint32Array([0, 1]), new Uint32Array([0]), new Uint32Array([2]));

    const packed = packPrograms([
      [encodePushCell(0), encodeCall(BUILTIN.LEN, 1), encodeRet()],
      [encodePushCell(1), encodeCall(BUILTIN.LEN, 1), encodeRet()],
      [encodePushCell(2), encodeCall(BUILTIN.LEN, 1), encodeRet()],
      [encodePushCell(3), encodeCall(BUILTIN.LEN, 1), encodeRet()],
      [encodePushCell(4), encodeCall(BUILTIN.LEN, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.LEN, 1), encodeRet()],
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
      ]),
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
      ]),
    );

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(0);
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(4);
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(6);
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(5);
    expect(kernel.readTags()[cellIndex(1, 5, width)]).toBe(ValueTag.Error);
    expect(kernel.readErrors()[cellIndex(1, 5, width)]).toBe(ErrorCode.Ref);
    expect(kernel.readTags()[cellIndex(1, 6, width)]).toBe(ValueTag.Error);
    expect(kernel.readErrors()[cellIndex(1, 6, width)]).toBe(ErrorCode.Value);
  });

  it("evaluates EXACT and numeric rounding builtins on the wasm path", async () => {
    const kernel = await createKernel();
    const width = 8;
    kernel.init(24, 8, 2, 1, 2);
    kernel.uploadStrings(
      Uint32Array.from([0, 0, 5, 10]),
      Uint32Array.from([0, 5, 5, 5]),
      asciiCodes("AlphaAlphaalpha"),
    );
    kernel.writeCells(
      new Uint8Array([3, 3, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([
        0, 0, -3.145, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
      ]),
      new Uint32Array([1, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(24),
    );
    const packed = packPrograms([
      [encodePushCell(0), encodePushCell(1), encodeCall(BUILTIN.EXACT, 2), encodeRet()],
      [encodePushCell(2), encodeCall(BUILTIN.INT, 1), encodeRet()],
      [encodePushCell(2), encodePushNumber(0), encodeCall(BUILTIN.ROUNDUP, 2), encodeRet()],
      [encodePushCell(2), encodePushNumber(0), encodeCall(BUILTIN.ROUNDDOWN, 2), encodeRet()],
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
      ]),
    );
    kernel.uploadConstants(
      new Float64Array([2]),
      new Uint32Array([0, 0, 0, 0]),
      new Uint32Array([0, 0, 1, 1]),
    );
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
      ]),
    );

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Boolean);
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(0);
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(-4);
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(-3.15);
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(-3.14);
  });

  it("evaluates string literals and CONCAT through the wasm path", async () => {
    const kernel = await createKernel();
    const width = 8;
    kernel.init(16, 4, 1, 1, 1);
    kernel.uploadStrings(
      Uint32Array.from([0, 0, 2]),
      Uint32Array.from([0, 2, 3]),
      asciiCodes("xyfoo"),
    );
    kernel.writeCells(
      new Uint8Array([ValueTag.String, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array(16),
      new Uint32Array([2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(16),
    );
    const packed = packPrograms([
      [encodePushString(1), encodeRet()],
      [encodePushString(1), encodePushCell(0), encodeCall(BUILTIN.CONCAT, 2), encodeRet()],
    ]);
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width)]),
    );
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0, 0]), new Uint32Array([0, 0]));
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width)]));

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.String);
    expect(kernel.readStringIds()[cellIndex(1, 1, width)]).toBe(1);
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.String);
    expect(kernel.readOutputStrings()).toEqual(["xyfoo"]);
  });

  it("evaluates binary text comparison and concat operators on the wasm path", async () => {
    const kernel = await createKernel();
    const width = 8;
    kernel.init(24, 4, 1, 1, 1);
    kernel.uploadStrings(
      Uint32Array.from([0, 0, 5, 10, 11]),
      Uint32Array.from([0, 5, 5, 1, 1]),
      asciiCodes("helloHELLObA"),
    );
    kernel.writeCells(
      new Uint8Array([
        ValueTag.String,
        ValueTag.String,
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
        0,
        0,
        0,
        0,
      ]),
      new Float64Array(24),
      new Uint32Array([1, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(24),
    );
    const packed = packPrograms([
      [encodePushCell(0), encodePushCell(1), encodeBinary(Opcode.Eq), encodeRet()],
      [encodePushString(3), encodePushString(4), encodeBinary(Opcode.Gt), encodeRet()],
      [encodePushCell(0), encodePushString(4), encodeBinary(Opcode.Concat), encodeRet()],
    ]);
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]),
    );
    kernel.uploadConstants(
      new Float64Array(),
      new Uint32Array([0, 0, 0]),
      new Uint32Array([0, 0, 0]),
    );
    kernel.evalBatch(
      Uint32Array.from([cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]),
    );

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Boolean);
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(1);
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.Boolean);
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(1);
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.String);
    expect(kernel.readOutputStrings()).toEqual(["helloA"]);
  });

  it("evaluates text slicing, casing, and search builtins on the wasm path", async () => {
    const kernel = await createKernel();
    const width = 10;
    kernel.init(40, 8, 2, 1, 1);
    kernel.uploadStrings(
      Uint32Array.from([0, 0, 5, 21, 26, 30, 38, 40]),
      Uint32Array.from([0, 5, 16, 5, 4, 8, 2, 2]),
      asciiCodes("Alpha  alpha   beta  alphaBETAalphabetphP*"),
    );
    kernel.writeCells(
      new Uint8Array([
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
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
      new Float64Array(40),
      new Uint32Array([
        1, 2, 3, 4, 5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
        0, 0, 0, 0, 0, 0, 0, 0, 0,
      ]),
      new Uint16Array(40),
    );
    const packed = packPrograms([
      [encodePushCell(0), encodePushNumber(0), encodeCall(BUILTIN.LEFT, 2), encodeRet()],
      [encodePushCell(0), encodeCall(BUILTIN.RIGHT, 1), encodeRet()],
      [
        encodePushCell(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodeCall(BUILTIN.MID, 3),
        encodeRet(),
      ],
      [encodePushCell(1), encodeCall(BUILTIN.TRIM, 1), encodeRet()],
      [encodePushCell(2), encodeCall(BUILTIN.UPPER, 1), encodeRet()],
      [encodePushCell(3), encodeCall(BUILTIN.LOWER, 1), encodeRet()],
      [encodePushString(6), encodePushCell(4), encodeCall(BUILTIN.FIND, 2), encodeRet()],
      [encodePushString(7), encodePushCell(4), encodeCall(BUILTIN.SEARCH, 2), encodeRet()],
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
      ]),
    );
    kernel.uploadConstants(
      new Float64Array([2, 2, 3]),
      new Uint32Array([0, 1, 1, 3, 3, 3, 3]),
      new Uint32Array([1, 0, 2, 0, 0, 0, 0]),
    );
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
      ]),
    );

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.String);
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.String);
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.String);
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.String);
    expect(kernel.readTags()[cellIndex(1, 5, width)]).toBe(ValueTag.String);
    expect(kernel.readTags()[cellIndex(1, 6, width)]).toBe(ValueTag.String);
    expect(kernel.readTags()[cellIndex(1, 7, width)]).toBe(ValueTag.Number);
    expect(kernel.readTags()[cellIndex(1, 8, width)]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[cellIndex(1, 7, width)]).toBe(3);
    expect(kernel.readNumbers()[cellIndex(1, 8, width)]).toBe(3);
    expect(kernel.readOutputStrings()).toEqual(["Al", "a", "lph", "alpha beta", "ALPHA", "beta"]);
  });

  it("evaluates REPLACE, SUBSTITUTE, and REPT on the wasm path", async () => {
    const kernel = await createKernel();
    const width = 10;
    kernel.init(24, 6, 1, 1, 1);
    kernel.uploadStrings(
      Uint32Array.from([0, 0, 8, 9, 15, 17, 19]),
      Uint32Array.from([0, 8, 1, 6, 2, 2, 2]),
      asciiCodes("alphabetZbananaanooxo"),
    );
    kernel.writeCells(
      new Uint8Array([
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
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
        0,
        0,
        0,
      ]),
      new Float64Array(24),
      new Uint32Array([1, 3, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(24),
    );
    const packed = packPrograms([
      [
        encodePushCell(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushString(2),
        encodeCall(BUILTIN.REPLACE, 4),
        encodeRet(),
      ],
      [
        encodePushCell(1),
        encodePushString(4),
        encodePushString(5),
        encodeCall(BUILTIN.SUBSTITUTE, 3),
        encodeRet(),
      ],
      [
        encodePushCell(1),
        encodePushString(4),
        encodePushString(5),
        encodePushNumber(0),
        encodeCall(BUILTIN.SUBSTITUTE, 4),
        encodeRet(),
      ],
      [encodePushCell(2), encodePushNumber(0), encodeCall(BUILTIN.REPT, 2), encodeRet()],
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
      ]),
    );
    kernel.uploadConstants(
      new Float64Array([3, 2, 2, 3]),
      new Uint32Array([0, 0, 2, 3]),
      new Uint32Array([2, 0, 1, 1]),
    );
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
      ]),
    );

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.String);
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.String);
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.String);
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.String);
    expect(kernel.readOutputStrings()).toEqual(["alZabet", "booooa", "banooa", "xoxoxo"]);
  });

  it("evaluates VALUE for dynamic scalar inputs on the wasm path", async () => {
    const kernel = await createKernel();
    const width = 8;
    kernel.init(24, 6, 1, 1, 1);
    kernel.uploadStrings(
      Uint32Array.from([0, 0, 4, 16]),
      Uint32Array.from([0, 4, 12, 3]),
      asciiCodes("42.5  -17.25e1  not"),
    );
    kernel.writeCells(
      new Uint8Array([
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.Boolean,
        ValueTag.Empty,
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
        0,
      ]),
      new Float64Array([0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array([1, 2, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array(24),
    );
    const packed = packPrograms([
      [encodePushCell(0), encodeCall(BUILTIN.VALUE, 1), encodeRet()],
      [encodePushCell(1), encodeCall(BUILTIN.VALUE, 1), encodeRet()],
      [encodePushCell(2), encodeCall(BUILTIN.VALUE, 1), encodeRet()],
      [encodePushCell(3), encodeCall(BUILTIN.VALUE, 1), encodeRet()],
      [encodePushCell(4), encodeCall(BUILTIN.VALUE, 1), encodeRet()],
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
      ]),
    );
    kernel.uploadConstants(
      new Float64Array(),
      new Uint32Array([0, 0, 0, 0, 0]),
      new Uint32Array([0, 0, 0, 0, 0]),
    );
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
      ]),
    );

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(42.5);
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(-172.5);
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.Error);
    expect(kernel.readErrors()[cellIndex(1, 3, width)]).toBe(ErrorCode.Value);
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(1);
    expect(kernel.readTags()[cellIndex(1, 5, width)]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(0);
  });

  it("evaluates IFERROR, IFNA, and NA on the wasm path", async () => {
    const kernel = await createKernel();
    const width = 8;
    kernel.init(24, 6, 1, 1, 1);
    kernel.uploadStrings(
      Uint32Array.from([0, 0, 8]),
      Uint32Array.from([0, 8, 7]),
      asciiCodes("fallbackmissing"),
    );
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Error,
        ValueTag.Error,
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
        0,
        0,
        0,
      ]),
      new Float64Array([
        ErrorCode.Div0,
        ErrorCode.Ref,
        7,
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
        0,
        0,
        0,
      ]),
      new Uint32Array(24),
      new Uint16Array([
        ErrorCode.Div0,
        ErrorCode.Ref,
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
        0,
        0,
        0,
        0,
      ]),
    );
    const packed = packPrograms([
      [encodePushCell(0), encodePushString(1), encodeCall(BUILTIN.IFERROR, 2), encodeRet()],
      [encodeCall(BUILTIN.NA, 0), encodePushString(2), encodeCall(BUILTIN.IFNA, 2), encodeRet()],
      [encodePushCell(1), encodePushString(2), encodeCall(BUILTIN.IFNA, 2), encodeRet()],
      [encodePushCell(2), encodePushString(1), encodeCall(BUILTIN.IFERROR, 2), encodeRet()],
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
      ]),
    );
    kernel.uploadConstants(
      new Float64Array(),
      new Uint32Array([0, 0, 0, 0]),
      new Uint32Array([0, 0, 0, 0]),
    );
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
      ]),
    );

    expect(kernel.readTags()[cellIndex(1, 1, width)]).toBe(ValueTag.String);
    expect(kernel.readStringIds()[cellIndex(1, 1, width)]).toBe(1);
    expect(kernel.readTags()[cellIndex(1, 2, width)]).toBe(ValueTag.String);
    expect(kernel.readStringIds()[cellIndex(1, 2, width)]).toBe(2);
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.Error);
    expect(kernel.readErrors()[cellIndex(1, 3, width)]).toBe(ErrorCode.Ref);
    expect(kernel.readTags()[cellIndex(1, 4, width)]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(7);
  });

  it("evaluates conditional aggregates and SUMPRODUCT on the wasm path", async () => {
    const kernel = await createKernel();
    const width = 10;
    kernel.init(32, 8, 5, 1, 2);
    kernel.uploadStrings(
      Uint32Array.from([0, 0, 2, 3]),
      Uint32Array.from([0, 2, 1, 1]),
      asciiCodes(">0xy"),
    );
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
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
      ]),
      new Float64Array([
        2, 4, -1, 6, 0, 0, 0, 0, 10, 20, 30, 40, 1, 2, 3, 4, 5, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
        0, 0, 0,
      ]),
      new Uint32Array([
        0, 0, 0, 0, 2, 2, 3, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
        0,
      ]),
      new Uint16Array(32),
    );
    kernel.uploadRangeMembers(
      new Uint32Array([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]),
      Uint32Array.from([0, 4, 8, 12, 15]),
      Uint32Array.from([4, 4, 4, 3, 3]),
    );
    kernel.uploadRangeShapes(Uint32Array.from([4, 4, 4, 3, 3]), Uint32Array.from([1, 1, 1, 1, 1]));

    const packed = packPrograms([
      [encodePushRange(0), encodePushString(1), encodeCall(BUILTIN.COUNTIF, 2), encodeRet()],
      [
        encodePushRange(0),
        encodePushString(1),
        encodePushRange(1),
        encodePushString(2),
        encodeCall(BUILTIN.COUNTIFS, 4),
        encodeRet(),
      ],
      [
        encodePushRange(0),
        encodePushString(1),
        encodePushRange(2),
        encodeCall(BUILTIN.SUMIF, 3),
        encodeRet(),
      ],
      [
        encodePushRange(2),
        encodePushRange(0),
        encodePushString(1),
        encodePushRange(1),
        encodePushString(2),
        encodeCall(BUILTIN.SUMIFS, 5),
        encodeRet(),
      ],
      [encodePushRange(0), encodePushString(1), encodeCall(BUILTIN.AVERAGEIF, 2), encodeRet()],
      [
        encodePushRange(2),
        encodePushRange(0),
        encodePushString(1),
        encodePushRange(1),
        encodePushString(2),
        encodeCall(BUILTIN.AVERAGEIFS, 5),
        encodeRet(),
      ],
      [encodePushRange(3), encodePushRange(4), encodeCall(BUILTIN.SUMPRODUCT, 2), encodeRet()],
    ]);
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(3, 1, width),
        cellIndex(3, 2, width),
        cellIndex(3, 3, width),
        cellIndex(3, 4, width),
        cellIndex(3, 5, width),
        cellIndex(3, 6, width),
        cellIndex(3, 7, width),
      ]),
    );
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0]), new Uint32Array([0]));
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(3, 1, width),
        cellIndex(3, 2, width),
        cellIndex(3, 3, width),
        cellIndex(3, 4, width),
        cellIndex(3, 5, width),
        cellIndex(3, 6, width),
        cellIndex(3, 7, width),
      ]),
    );

    expect(kernel.readTags()[cellIndex(3, 1, width)]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[cellIndex(3, 1, width)]).toBe(3);
    expect(kernel.readTags()[cellIndex(3, 2, width)]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[cellIndex(3, 2, width)]).toBe(3);
    expect(kernel.readTags()[cellIndex(3, 3, width)]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[cellIndex(3, 3, width)]).toBe(70);
    expect(kernel.readTags()[cellIndex(3, 4, width)]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[cellIndex(3, 4, width)]).toBe(70);
    expect(kernel.readTags()[cellIndex(3, 5, width)]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[cellIndex(3, 5, width)]).toBe(4);
    expect(kernel.readTags()[cellIndex(3, 6, width)]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[cellIndex(3, 6, width)]).toBeCloseTo((10 + 20 + 40) / 3);
    expect(kernel.readTags()[cellIndex(3, 7, width)]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[cellIndex(3, 7, width)]).toBe(32);
  });

  it("evaluates INDEX, VLOOKUP, and HLOOKUP on the wasm path", async () => {
    const kernel = await createKernel();
    const width = 8;
    kernel.init(32, 6, 3, 3, 12);
    kernel.uploadStrings(
      Uint32Array.from([0, 4, 9, 11, 13]),
      Uint32Array.from([4, 5, 2, 2, 2]),
      asciiCodes("pearappleQ1Q2Q3"),
    );
    kernel.writeCells(
      new Uint8Array([
        ValueTag.String,
        ValueTag.Number,
        0,
        0,
        ValueTag.String,
        ValueTag.Number,
        0,
        0,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        0,
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
      ]),
      new Float64Array([
        0, 10, 0, 0, 0, 20, 0, 0, 0, 0, 0, 0, 100, 200, 300, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
        0, 0, 0, 0,
      ]),
      new Uint32Array([
        0, 0, 0, 0, 1, 0, 0, 0, 2, 3, 4, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
        0,
      ]),
      new Uint16Array(32),
    );
    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 4, 5, 0, 1, 4, 5, 8, 9, 10, 12, 13, 14]),
      Uint32Array.from([0, 4, 8]),
      Uint32Array.from([4, 4, 6]),
    );
    kernel.uploadRangeShapes(Uint32Array.from([2, 2, 2]), Uint32Array.from([2, 2, 3]));

    const packed = packPrograms([
      [
        encodePushRange(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodeCall(BUILTIN.INDEX, 3),
        encodeRet(),
      ],
      [
        encodePushString(1),
        encodePushRange(1),
        encodePushNumber(0),
        encodePushNumber(1),
        encodeCall(BUILTIN.VLOOKUP, 4),
        encodeRet(),
      ],
      [
        encodePushString(4),
        encodePushRange(2),
        encodePushNumber(0),
        encodePushNumber(1),
        encodeCall(BUILTIN.HLOOKUP, 4),
        encodeRet(),
      ],
    ]);
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(2, 0, width), cellIndex(2, 1, width), cellIndex(2, 2, width)]),
    );
    kernel.uploadConstants(
      new Float64Array([2, 2, 2, 0, 2, 0]),
      new Uint32Array([0, 2, 4]),
      new Uint32Array([2, 2, 2]),
    );

    kernel.evalBatch(
      Uint32Array.from([cellIndex(2, 0, width), cellIndex(2, 1, width), cellIndex(2, 2, width)]),
    );

    expect(kernel.readNumbers()[cellIndex(2, 0, width)]).toBe(20);
    expect(kernel.readNumbers()[cellIndex(2, 1, width)]).toBe(20);
    expect(kernel.readNumbers()[cellIndex(2, 2, width)]).toBe(300);
  });

  it("evaluates vector lookup builtins on the wasm path", async () => {
    const kernel = await createKernel();
    const width = 10;
    kernel.init(40, 8, 6, 1, 2);
    kernel.uploadStrings(
      Uint32Array.from([0, 0, 5, 9, 14, 18, 26]),
      Uint32Array.from([0, 5, 4, 5, 4, 8, 8]),
      asciiCodes("applepearpearplumfallbacknotfound"),
    );
    kernel.writeCells(
      new Uint8Array([
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.Number,
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
      new Float64Array([
        0, 0, 0, 0, 10, 20, 30, 40, 1, 3, 5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
      ]),
      new Uint32Array([
        1, 2, 2, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
        0, 0, 0, 0, 0, 0, 0, 0, 0,
      ]),
      new Uint16Array(40),
    );
    kernel.uploadRangeMembers(
      new Uint32Array([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10]),
      Uint32Array.from([0, 4, 8]),
      Uint32Array.from([4, 4, 3]),
    );
    kernel.uploadRangeShapes(Uint32Array.from([4, 4, 3]), Uint32Array.from([1, 1, 1]));

    const packed = packPrograms([
      [
        encodePushString(2),
        encodePushRange(0),
        encodePushNumber(0),
        encodeCall(BUILTIN.MATCH, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(1),
        encodePushRange(2),
        encodePushNumber(2),
        encodeCall(BUILTIN.MATCH, 3),
        encodeRet(),
      ],
      [
        encodePushString(2),
        encodePushRange(0),
        encodePushNumber(0),
        encodePushNumber(3),
        encodeCall(BUILTIN.XMATCH, 4),
        encodeRet(),
      ],
      [
        encodePushString(2),
        encodePushRange(0),
        encodePushRange(1),
        encodeCall(BUILTIN.XLOOKUP, 3),
        encodeRet(),
      ],
      [
        encodePushString(6),
        encodePushRange(0),
        encodePushRange(1),
        encodePushString(5),
        encodeCall(BUILTIN.XLOOKUP, 4),
        encodeRet(),
      ],
    ]);
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(4, 1, width),
        cellIndex(4, 2, width),
        cellIndex(4, 3, width),
        cellIndex(4, 4, width),
        cellIndex(4, 5, width),
      ]),
    );
    kernel.uploadConstants(
      new Float64Array([0, 4, 1, -1]),
      new Uint32Array([0, 0, 0, 0, 0]),
      new Uint32Array([1, 2, 2, 0, 0]),
    );
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(4, 1, width),
        cellIndex(4, 2, width),
        cellIndex(4, 3, width),
        cellIndex(4, 4, width),
        cellIndex(4, 5, width),
      ]),
    );

    expect(kernel.readTags()[cellIndex(4, 1, width)]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[cellIndex(4, 1, width)]).toBe(2);
    expect(kernel.readTags()[cellIndex(4, 2, width)]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[cellIndex(4, 2, width)]).toBe(2);
    expect(kernel.readTags()[cellIndex(4, 3, width)]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[cellIndex(4, 3, width)]).toBe(3);
    expect(kernel.readTags()[cellIndex(4, 4, width)]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[cellIndex(4, 4, width)]).toBe(20);
    expect(kernel.readTags()[cellIndex(4, 5, width)]).toBe(ValueTag.String);
    expect(kernel.readStringIds()[cellIndex(4, 5, width)]).toBe(5);
  });

  it("evaluates exact-safe date builtins with Excel coercion and errors", async () => {
    const kernel = await createKernel();
    const width = 10;
    kernel.init(20, 10, 5, 2, 2);
    kernel.writeCells(
      new Uint8Array([3, 2, 4, 1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([
        0, 1, 0, 45351, 45351.75, 60, 45322, 45337, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
      ]),
      new Uint32Array([1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint16Array([0, 0, ErrorCode.Ref, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
    );

    const packed = packPrograms([
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BUILTIN.DATE, 3),
        encodeRet(),
      ],
      [
        encodePushCell(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BUILTIN.DATE, 3),
        encodeRet(),
      ],
      [encodePushCell(3), encodeCall(BUILTIN.YEAR, 1), encodeRet()],
      [encodePushCell(4), encodeCall(BUILTIN.MONTH, 1), encodeRet()],
      [encodePushCell(5), encodeCall(BUILTIN.DAY, 1), encodeRet()],
      [encodePushCell(6), encodePushNumber(3), encodeCall(BUILTIN.EDATE, 2), encodeRet()],
      [encodePushCell(0), encodePushNumber(4), encodeCall(BUILTIN.EDATE, 2), encodeRet()],
      [encodePushCell(7), encodePushCell(1), encodeCall(BUILTIN.EOMONTH, 2), encodeRet()],
      [encodePushCell(2), encodePushNumber(4), encodeCall(BUILTIN.EOMONTH, 2), encodeRet()],
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
        cellIndex(1, 9, width),
      ]),
    );
    kernel.uploadConstants(
      new Float64Array([2024, 2, 29, 1.9, 1]),
      new Uint32Array([0]),
      new Uint32Array([5]),
    );
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
        cellIndex(1, 9, width),
      ]),
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

  it("evaluates numeric-only dynamic-array builtins on the wasm path", async () => {
    const kernel = await createKernel();
    kernel.init(24, 11, 1, 1, 1);
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
        ValueTag.Number,
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
      ]),
      new Float64Array([1, 2, 3, 4, 5, 6, 4, 3, 2, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(24),
      new Uint16Array(24),
    );
    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 2, 3, 4, 5, 6, 7, 8, 9]),
      Uint32Array.from([0, 6]),
      Uint32Array.from([6, 4]),
    );
    kernel.uploadRangeShapes(Uint32Array.from([2, 4]), Uint32Array.from([3, 1]));

    const packed = packPrograms([
      [
        encodePushRange(0),
        encodePushNumber(0),
        encodePushNumber(0),
        encodePushNumber(2),
        encodePushNumber(1),
        encodeCall(BUILTIN.OFFSET, 5),
        encodeRet(),
      ],
      [encodePushRange(0), encodePushNumber(2), encodeCall(BUILTIN.TAKE, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(1), encodeCall(BUILTIN.DROP, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(2), encodeCall(BUILTIN.CHOOSECOLS, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(2), encodeCall(BUILTIN.CHOOSEROWS, 2), encodeRet()],
      [
        encodePushRange(0),
        encodePushNumber(2),
        encodePushNumber(4),
        encodeCall(BUILTIN.SORT, 3),
        encodeRet(),
      ],
      [encodePushRange(1), encodePushRange(1), encodeCall(BUILTIN.SORTBY, 2), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.TOCOL, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BUILTIN.TOROW, 1), encodeRet()],
      [encodePushRange(0), encodePushNumber(2), encodeCall(BUILTIN.WRAPROWS, 2), encodeRet()],
      [encodePushRange(0), encodePushNumber(2), encodeCall(BUILTIN.WRAPCOLS, 2), encodeRet()],
    ]);
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22]),
    );
    kernel.uploadConstants(
      new Float64Array([0, 1, 2, 3, -1]),
      new Uint32Array([0, 0, 0, 0, 0]),
      new Uint32Array([1, 1, 1, 1, 1]),
    );
    kernel.evalBatch(Uint32Array.from([12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22]));

    expect(kernel.readTags()[12]).toBe(ValueTag.Number);
    expect(kernel.readTags()[13]).toBe(ValueTag.Number);
    expect(kernel.readTags()[14]).toBe(ValueTag.Number);
    expect(kernel.readTags()[15]).toBe(ValueTag.Number);
    expect(kernel.readTags()[16]).toBe(ValueTag.Number);
    expect(kernel.readTags()[17]).toBe(ValueTag.Number);
    expect(kernel.readTags()[18]).toBe(ValueTag.Number);
    expect(kernel.readTags()[19]).toBe(ValueTag.Number);
    expect(kernel.readTags()[20]).toBe(ValueTag.Number);
    expect(kernel.readTags()[21]).toBe(ValueTag.Number);
    expect(kernel.readTags()[22]).toBe(ValueTag.Number);

    expect(kernel.readSpillRows()[12]).toBe(2);
    expect(kernel.readSpillCols()[12]).toBe(1);
    expect(kernel.readSpillLengths()[12]).toBe(2);
    expect(
      Array.from(
        kernel
          .readSpillNumbers()
          .slice(kernel.readSpillOffsets()[12], kernel.readSpillOffsets()[12] + 2),
      ),
    ).toEqual([1, 4]);

    expect(kernel.readSpillRows()[13]).toBe(2);
    expect(kernel.readSpillCols()[13]).toBe(3);
    expect(kernel.readSpillLengths()[13]).toBe(6);
    expect(
      Array.from(
        kernel
          .readSpillNumbers()
          .slice(kernel.readSpillOffsets()[13], kernel.readSpillOffsets()[13] + 6),
      ),
    ).toEqual([1, 2, 3, 4, 5, 6]);
    expect(kernel.readSpillRows()[14]).toBe(1);
    expect(kernel.readSpillCols()[14]).toBe(3);
    expect(kernel.readSpillLengths()[14]).toBe(3);
    expect(
      Array.from(
        kernel
          .readSpillNumbers()
          .slice(kernel.readSpillOffsets()[14], kernel.readSpillOffsets()[14] + 3),
      ),
    ).toEqual([4, 5, 6]);
    expect(kernel.readSpillRows()[15]).toBe(2);
    expect(kernel.readSpillCols()[15]).toBe(1);
    expect(kernel.readSpillLengths()[15]).toBe(2);
    expect(
      Array.from(
        kernel
          .readSpillNumbers()
          .slice(kernel.readSpillOffsets()[15], kernel.readSpillOffsets()[15] + 2),
      ),
    ).toEqual([2, 5]);
    expect(kernel.readSpillRows()[16]).toBe(1);
    expect(kernel.readSpillCols()[16]).toBe(3);
    expect(kernel.readSpillLengths()[16]).toBe(3);
    expect(
      Array.from(
        kernel
          .readSpillNumbers()
          .slice(kernel.readSpillOffsets()[16], kernel.readSpillOffsets()[16] + 3),
      ),
    ).toEqual([4, 5, 6]);
    expect(kernel.readSpillRows()[17]).toBe(2);
    expect(kernel.readSpillCols()[17]).toBe(3);
    expect(kernel.readSpillLengths()[17]).toBe(6);
    expect(
      Array.from(
        kernel
          .readSpillNumbers()
          .slice(kernel.readSpillOffsets()[17], kernel.readSpillOffsets()[17] + 6),
      ),
    ).toEqual([1, 1, 1, 4, 4, 4]);
    expect(kernel.readSpillRows()[18]).toBe(4);
    expect(kernel.readSpillCols()[18]).toBe(1);
    expect(kernel.readSpillLengths()[18]).toBe(4);
    expect(
      Array.from(
        kernel
          .readSpillNumbers()
          .slice(kernel.readSpillOffsets()[18], kernel.readSpillOffsets()[18] + 4),
      ),
    ).toEqual([4, 4, 4, 4]);
    expect(kernel.readSpillRows()[19]).toBe(6);
    expect(kernel.readSpillCols()[19]).toBe(1);
    expect(kernel.readSpillLengths()[19]).toBe(6);
    expect(
      Array.from(
        kernel
          .readSpillNumbers()
          .slice(kernel.readSpillOffsets()[19], kernel.readSpillOffsets()[19] + 6),
      ),
    ).toEqual([1, 4, 2, 5, 3, 6]);
    expect(kernel.readSpillRows()[20]).toBe(1);
    expect(kernel.readSpillCols()[20]).toBe(6);
    expect(kernel.readSpillLengths()[20]).toBe(6);
    expect(
      Array.from(
        kernel
          .readSpillNumbers()
          .slice(kernel.readSpillOffsets()[20], kernel.readSpillOffsets()[20] + 6),
      ),
    ).toEqual([1, 2, 3, 4, 5, 6]);
    expect(kernel.readSpillRows()[21]).toBe(3);
    expect(kernel.readSpillCols()[21]).toBe(2);
    expect(kernel.readSpillLengths()[21]).toBe(6);
    expect(
      Array.from(
        kernel
          .readSpillNumbers()
          .slice(kernel.readSpillOffsets()[21], kernel.readSpillOffsets()[21] + 6),
      ),
    ).toEqual([1, 2, 3, 4, 5, 6]);
    expect(kernel.readSpillRows()[22]).toBe(2);
    expect(kernel.readSpillCols()[22]).toBe(3);
    expect(kernel.readSpillLengths()[22]).toBe(6);
    expect(
      Array.from(
        kernel
          .readSpillNumbers()
          .slice(kernel.readSpillOffsets()[22], kernel.readSpillOffsets()[22] + 6),
      ),
    ).toEqual([1, 2, 3, 4, 5, 6]);
  });

  it("evaluates RAND from the uploaded recalc random sequence on the wasm path", async () => {
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
        encodeCall(BUILTIN.RAND, 0),
        encodeRet(),
        encodeCall(BUILTIN.RAND, 0),
        encodeCall(BUILTIN.RAND, 0),
        encodeBinary(Opcode.Add),
        encodeRet(),
      ]),
      new Uint32Array([0, 2]),
      new Uint32Array([2, 4]),
      new Uint32Array([0, 1]),
    );
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0, 0]), new Uint32Array([0, 0]));
    kernel.uploadVolatileRandomValues(new Float64Array([0.625, 0.125, 0.875]));

    kernel.evalBatch(new Uint32Array([0, 1]));

    expect(kernel.readTags()[0]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[0]).toBe(0.625);
    expect(kernel.readTags()[1]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[1]).toBeCloseTo(1, 12);
  });

  it("returns numeric spill descriptors for SEQUENCE on the wasm path", async () => {
    const kernel = await createKernel();
    kernel.init(4, 4, 4, 1, 1);
    kernel.writeCells(
      new Uint8Array(4),
      new Float64Array(4),
      new Uint32Array(4),
      new Uint16Array(4),
    );
    kernel.uploadPrograms(
      new Uint32Array([
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodeCall(BuiltinId.Sequence, 4),
        encodeRet(),
      ]),
      new Uint32Array([0]),
      new Uint32Array([6]),
      new Uint32Array([0]),
    );
    kernel.uploadConstants(
      new Float64Array([3, 1, 1, 1]),
      new Uint32Array([0]),
      new Uint32Array([4]),
    );

    kernel.evalBatch(new Uint32Array([0]));

    expect(kernel.readTags()[0]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[0]).toBe(1);
    expect(kernel.readSpillRows()[0]).toBe(3);
    expect(kernel.readSpillCols()[0]).toBe(1);
    expect(kernel.readSpillOffsets()[0]).toBe(0);
    expect(kernel.readSpillLengths()[0]).toBe(3);
    expect(Array.from(kernel.readSpillNumbers().slice(0, kernel.getSpillValueCount()))).toEqual([
      1, 2, 3,
    ]);
  });

  it("evaluates numeric aggregate builtins over native SEQUENCE arrays on the wasm path", async () => {
    const kernel = await createKernel();
    kernel.init(12, 4, 24, 1, 1);
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
        ValueTag.Empty,
      ]),
      new Float64Array([3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(12),
      new Uint16Array(12),
    );
    kernel.uploadPrograms(
      new Uint32Array([
        encodePushCell(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Sequence, 4),
        encodeCall(BuiltinId.Sum, 1),
        encodeRet(),
        encodePushCell(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Sequence, 4),
        encodeCall(BuiltinId.Avg, 1),
        encodeRet(),
        encodePushCell(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Sequence, 4),
        encodeCall(BuiltinId.Min, 1),
        encodeRet(),
        encodePushCell(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Sequence, 4),
        encodeCall(BuiltinId.Max, 1),
        encodeRet(),
        encodePushCell(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Sequence, 4),
        encodeCall(BuiltinId.Count, 1),
        encodeRet(),
        encodePushCell(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Sequence, 4),
        encodeCall(BuiltinId.CountA, 1),
        encodeRet(),
      ]),
      new Uint32Array([0, 7, 14, 21, 28, 35]),
      new Uint32Array([7, 7, 7, 7, 7, 7]),
      new Uint32Array([1, 2, 3, 4, 5, 6]),
    );
    kernel.uploadConstants(
      new Float64Array([1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]),
      new Uint32Array([0, 3, 6, 9, 12, 15]),
      new Uint32Array([3, 3, 3, 3, 3, 3]),
    );

    kernel.evalBatch(new Uint32Array([1, 2, 3, 4, 5, 6]));

    expect(kernel.readTags()[1]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[1]).toBe(6);
    expect(kernel.readNumbers()[2]).toBe(2);
    expect(kernel.readNumbers()[3]).toBe(1);
    expect(kernel.readNumbers()[4]).toBe(3);
    expect(kernel.readNumbers()[5]).toBe(3);
    expect(kernel.readNumbers()[6]).toBe(3);
  });

  it("evaluates TODAY and NOW from the uploaded recalc timestamp on the wasm path", async () => {
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
        encodeCall(BUILTIN.TODAY, 0),
        encodeRet(),
        encodeCall(BUILTIN.NOW, 0),
        encodeRet(),
      ]),
      new Uint32Array([0, 2]),
      new Uint32Array([2, 2]),
      new Uint32Array([0, 1]),
    );
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0, 0]), new Uint32Array([0, 0]));
    kernel.uploadVolatileNowSerial(46100.65659722222);

    kernel.evalBatch(new Uint32Array([0, 1]));

    expect(kernel.readTags()[0]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[0]).toBe(46100);
    expect(kernel.readTags()[1]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[1]).toBeCloseTo(46100.65659722222, 12);
  });

  it("evaluates TIME, HOUR, MINUTE, SECOND, and WEEKDAY on the wasm path", async () => {
    const kernel = await createKernel();
    const width = 10;
    kernel.init(24, 8, 5, 1, 1);
    kernel.writeCells(
      new Uint8Array([1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([
        0.5208333333333334, 0.5208449074074074, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
        0, 0, 0, 0, 0,
      ]),
      new Uint32Array(24),
      new Uint16Array(24),
    );

    const packed = packPrograms([
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BUILTIN.TIME, 3),
        encodeRet(),
      ],
      [encodePushCell(0), encodeCall(BUILTIN.HOUR, 1), encodeRet()],
      [encodePushCell(0), encodeCall(BUILTIN.MINUTE, 1), encodeRet()],
      [encodePushCell(1), encodeCall(BUILTIN.SECOND, 1), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BUILTIN.DATE, 3),
        encodeCall(BUILTIN.WEEKDAY, 1),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BUILTIN.DATE, 3),
        encodePushNumber(3),
        encodeCall(BUILTIN.WEEKDAY, 2),
        encodeRet(),
      ],
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
      ]),
    );
    kernel.uploadConstants(
      new Float64Array([12, 30, 0, 2026, 3, 15, 2]),
      new Uint32Array([0, 0, 0, 0, 3, 3]),
      new Uint32Array([3, 0, 0, 0, 3, 4]),
    );
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
      ]),
    );

    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(0.5208333333333334);
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(12);
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(30);
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(1);
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(1);
    expect(kernel.readNumbers()[cellIndex(1, 6, width)]).toBe(7);
  });

  it("evaluates DAYS, WEEKNUM, WORKDAY, and NETWORKDAYS on the wasm path", async () => {
    const kernel = await createKernel();
    const width = 10;
    kernel.init(30, 8, 1, 1, 1);
    kernel.writeCells(
      new Uint8Array([
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
      new Float64Array([
        46097, 46101, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
        0, 0,
      ]),
      new Uint32Array(30),
      new Uint16Array(30),
    );
    kernel.uploadRangeMembers(new Uint32Array([0, 1]), new Uint32Array([0]), new Uint32Array([2]));
    kernel.uploadRangeShapes(Uint32Array.from([2]), Uint32Array.from([1]));

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.DAYS, 2), encodeRet()],
      [encodePushNumber(0), encodeCall(BUILTIN.WEEKNUM, 1), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.WEEKNUM, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.WORKDAY, 2), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushCell(0),
        encodeCall(BUILTIN.WORKDAY, 3),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BUILTIN.NETWORKDAYS, 2), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushCell(0),
        encodeCall(BUILTIN.NETWORKDAYS, 3),
        encodeRet(),
      ],
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
      ]),
    );
    kernel.uploadConstants(
      new Float64Array([
        46101, 46094, 46096, 46096, 2, 46094, 1, 46094, 1, 46094, 46101, 46094, 46101,
      ]),
      new Uint32Array([0, 2, 3, 5, 7, 9, 11]),
      new Uint32Array([2, 1, 2, 2, 2, 2, 2]),
    );
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
        cellIndex(1, 6, width),
        cellIndex(1, 7, width),
      ]),
    );

    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(7);
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(12);
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(11);
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(46097);
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(46098);
    expect(kernel.readNumbers()[cellIndex(1, 6, width)]).toBe(6);
    expect(kernel.readNumbers()[cellIndex(1, 7, width)]).toBe(5);
  });

  it("evaluates logical and rounding builtins with parity-safe scalar semantics", async () => {
    const kernel = await createKernel();
    kernel.init(8, 8, 4, 4, 4);
    kernel.writeCells(
      new Uint8Array([1, 1, 4, 0, 0, 0, 0, 0]),
      new Float64Array([123.4, 1, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(8),
      new Uint16Array([0, 0, ErrorCode.Value, 0, 0, 0, 0, 0]),
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
        255 << 24,
      ]),
      new Uint32Array([0, 4, 8, 11]),
      new Uint32Array([4, 4, 3, 4]),
      new Uint32Array([3, 4, 5, 6]),
    );
    kernel.uploadConstants(
      new Float64Array([-1, 0.5, 1]),
      new Uint32Array([0, 0, 0, 0]),
      new Uint32Array([2, 2, 0, 1]),
    );

    kernel.evalBatch(new Uint32Array([3, 4, 5, 6]));

    expect(kernel.readNumbers()[3]).toBe(120);
    expect(kernel.readNumbers()[4]).toBe(1);
    expect(kernel.readTags()[5]).toBe(ValueTag.Boolean);
    expect(kernel.readNumbers()[5]).toBe(0);
    expect(kernel.readTags()[6]).toBe(ValueTag.Error);
    expect(kernel.readErrors()[6]).toBe(ErrorCode.Value);
  });

  it("materializes pivots using the actual source width", async () => {
    const kernel = await createKernel();
    kernel.init(16, 1, 1, 1, 16);

    const strings = [
      "",
      "Region",
      "Notes",
      "Product",
      "Sales",
      "East",
      "Widget",
      "West",
      "Gizmo",
      "priority",
    ];
    const offsets = new Uint32Array(strings.length);
    const lengths = new Uint32Array(strings.length);
    const data: number[] = [];
    let offset = 0;
    strings.forEach((text, index) => {
      offsets[index] = offset;
      lengths[index] = text.length;
      for (const char of text) {
        data.push(char.charCodeAt(0));
      }
      offset += text.length;
    });
    kernel.uploadStrings(offsets, lengths, Uint16Array.from(data));

    kernel.writeCells(
      new Uint8Array([
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.Number,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.Number,
        ValueTag.String,
        ValueTag.String,
        ValueTag.String,
        ValueTag.Number,
      ]),
      new Float64Array([0, 0, 0, 0, 0, 0, 0, 10, 0, 0, 0, 7, 0, 0, 0, 5]),
      new Uint32Array([1, 2, 3, 4, 5, 9, 6, 0, 7, 9, 6, 0, 5, 9, 8, 0]),
      new Uint16Array(16),
    );
    kernel.uploadRangeMembers(
      Uint32Array.from(Array.from({ length: 16 }, (_, index) => index)),
      new Uint32Array([0]),
      new Uint32Array([16]),
    );
    kernel.uploadRangeShapes(new Uint32Array([4]), new Uint32Array([4]));

    const materialized = kernel.materializePivotTable(
      0,
      4,
      Uint32Array.from([0]),
      Uint32Array.from([3, 2]),
      Uint8Array.from([1, 2]),
    );

    expect(materialized.rows).toBe(3);
    expect(materialized.cols).toBe(3);
    expect(materialized.tags[0]).toBe(ValueTag.String);
    expect(materialized.stringIds[0]).toBe(1);
    expect(materialized.tags[1]).toBe(ValueTag.String);
    expect(materialized.stringIds[1]).toBe(4);
    expect(materialized.tags[2]).toBe(ValueTag.String);
    expect(materialized.stringIds[2]).toBe(3);
    expect(materialized.tags[3]).toBe(ValueTag.String);
    expect(materialized.stringIds[3]).toBe(5);
    expect(materialized.tags[4]).toBe(ValueTag.Number);
    expect(materialized.numbers[4]).toBe(15);
    expect(materialized.tags[5]).toBe(ValueTag.Number);
    expect(materialized.numbers[5]).toBe(2);
    expect(materialized.tags[6]).toBe(ValueTag.String);
    expect(materialized.stringIds[6]).toBe(7);
    expect(materialized.tags[7]).toBe(ValueTag.Number);
    expect(materialized.numbers[7]).toBe(7);
    expect(materialized.tags[8]).toBe(ValueTag.Number);
    expect(materialized.numbers[8]).toBe(1);
  });
});
