import { describe, expect, it } from "vitest";
import { BuiltinId, ErrorCode, Opcode, ValueTag, type CellValue } from "@bilig/protocol";
import { createKernel, type KernelInstance } from "../index.js";

const OUTPUT_STRING_BASE = 2147483648;

function encodeCall(builtinId: number, argc: number): number {
  return (Opcode.CallBuiltin << 24) | ((builtinId << 8) | argc);
}

function encodePushNumber(constantIndex: number): number {
  return (Opcode.PushNumber << 24) | constantIndex;
}

function encodePushString(stringId: number): number {
  return (Opcode.PushString << 24) | stringId;
}

function encodePushBoolean(value: boolean): number {
  return (Opcode.PushBoolean << 24) | (value ? 1 : 0);
}

function encodePushError(code: ErrorCode): number {
  return (Opcode.PushError << 24) | code;
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

function uploadPooledStrings(kernel: KernelInstance, strings: readonly string[]): void {
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
}

function decodeStringValue(
  rawValue: number,
  pooledStrings: readonly string[],
  outputStrings: readonly string[],
): string {
  const outputIndex = rawValue >= OUTPUT_STRING_BASE ? rawValue - OUTPUT_STRING_BASE : -1;
  return outputIndex >= 0 ? (outputStrings[outputIndex] ?? "") : (pooledStrings[rawValue] ?? "");
}

function toErrorCode(rawValue: number): ErrorCode {
  switch (rawValue) {
    case 0:
      return ErrorCode.None;
    case 1:
      return ErrorCode.Div0;
    case 2:
      return ErrorCode.Ref;
    case 3:
      return ErrorCode.Value;
    case 4:
      return ErrorCode.Name;
    case 5:
      return ErrorCode.NA;
    case 6:
      return ErrorCode.Cycle;
    case 7:
      return ErrorCode.Spill;
    case 8:
      return ErrorCode.Blocked;
    default:
      throw new Error(`Unexpected error code: ${rawValue}`);
  }
}

function readScalarValue(
  kernel: KernelInstance,
  cellIndex: number,
  pooledStrings: readonly string[],
): CellValue {
  const tag = kernel.readTags()[cellIndex] ?? ValueTag.Empty;
  if (tag === ValueTag.Empty) {
    return { tag };
  }
  if (tag === ValueTag.Number || tag === ValueTag.Boolean) {
    return { tag, value: kernel.readNumbers()[cellIndex] ?? 0 };
  }
  if (tag === ValueTag.String) {
    return {
      tag,
      value: decodeStringValue(
        kernel.readStringIds()[cellIndex] ?? 0,
        pooledStrings,
        kernel.readOutputStrings(),
      ),
      stringId: 0,
    };
  }
  if (tag === ValueTag.Error) {
    return { tag, code: toErrorCode(kernel.readErrors()[cellIndex] ?? 0) };
  }
  throw new Error(`Unexpected scalar tag: ${tag}`);
}

function readSpillValues(
  kernel: KernelInstance,
  ownerCellIndex: number,
  pooledStrings: readonly string[],
): CellValue[] {
  const offset = kernel.readSpillOffsets()[ownerCellIndex] ?? 0;
  const length = kernel.readSpillLengths()[ownerCellIndex] ?? 0;
  const tags = kernel.readSpillTags();
  const values = kernel.readSpillNumbers();
  const outputStrings = kernel.readOutputStrings();
  return Array.from({ length }, (_, index) => {
    const tag = tags[offset + index] ?? ValueTag.Empty;
    const rawValue = values[offset + index] ?? 0;
    if (tag === ValueTag.Number) {
      return { tag, value: rawValue };
    }
    if (tag === ValueTag.String) {
      return {
        tag,
        value: decodeStringValue(rawValue, pooledStrings, outputStrings),
        stringId: 0,
      };
    }
    if (tag === ValueTag.Error) {
      return { tag, code: toErrorCode(rawValue) };
    }
    return { tag: ValueTag.Empty };
  });
}

describe("wasm kernel array/info dispatch", () => {
  it("keeps array, text, and info builtin dispatch stable across refactors", async () => {
    const kernel = await createKernel();
    const ownerBase = 64;
    const pooledStrings = ["fallback", "na-value", "x", "-", "a", "b", "3"];

    kernel.init(96, pooledStrings.length, 4, 16, 32);
    uploadPooledStrings(kernel, pooledStrings);

    const cellTags = new Uint8Array(96);
    const cellNumbers = new Float64Array(96);
    const cellStringIds = new Uint32Array(96);
    const cellErrors = new Uint16Array(96);

    cellTags[0] = ValueTag.Number;
    cellNumbers[0] = 1;
    cellTags[1] = ValueTag.Number;
    cellNumbers[1] = 2;
    cellTags[16] = ValueTag.Number;
    cellNumbers[16] = 3;
    cellTags[17] = ValueTag.Number;
    cellNumbers[17] = 4;

    cellTags[2] = ValueTag.Number;
    cellNumbers[2] = 5;
    cellTags[18] = ValueTag.Number;
    cellNumbers[18] = 6;

    cellTags[3] = ValueTag.Number;
    cellNumbers[3] = 7;
    cellTags[4] = ValueTag.Number;
    cellNumbers[4] = 8;

    cellTags[5] = ValueTag.String;
    cellStringIds[5] = 4;
    cellTags[21] = ValueTag.Number;
    cellNumbers[21] = 2;
    cellTags[22] = ValueTag.String;
    cellStringIds[22] = 5;

    kernel.writeCells(cellTags, cellNumbers, cellStringIds, cellErrors);
    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 16, 17, 2, 18, 3, 4, 5, 6, 21, 22]),
      Uint32Array.from([0, 4, 6, 8]),
      Uint32Array.from([4, 2, 2, 4]),
    );
    kernel.uploadRangeShapes(Uint32Array.from([2, 2, 1, 2]), Uint32Array.from([2, 1, 2, 2]));

    const packed = packPrograms([
      [encodePushRange(0), encodeCall(BuiltinId.Areas, 1), encodeRet()],
      [encodePushRange(0), encodePushNumber(0), encodeCall(BuiltinId.Arraytotext, 2), encodeRet()],
      [encodePushRange(0), encodeCall(BuiltinId.Columns, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BuiltinId.Rows, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BuiltinId.Transpose, 1), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BuiltinId.Hstack, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(2), encodeCall(BuiltinId.Vstack, 2), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodeCall(BuiltinId.Choose, 4),
        encodeRet(),
      ],
      [
        encodePushString(3),
        encodePushBoolean(true),
        encodePushRange(3),
        encodeCall(BuiltinId.Textjoin, 3),
        encodeRet(),
      ],
      [encodeCall(BuiltinId.Na, 0), encodeRet()],
      [
        encodePushError(ErrorCode.Ref),
        encodePushString(0),
        encodeCall(BuiltinId.Iferror, 2),
        encodeRet(),
      ],
      [
        encodeCall(BuiltinId.Na, 0),
        encodePushString(1),
        encodeCall(BuiltinId.Ifna, 2),
        encodeRet(),
      ],
      [encodePushString(2), encodeCall(BuiltinId.T, 1), encodeRet()],
      [encodePushBoolean(true), encodeCall(BuiltinId.N, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BuiltinId.Type, 1), encodeRet()],
      [encodePushString(6), encodePushNumber(0), encodeCall(BuiltinId.Delta, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Gestep, 2), encodeRet()],
    ]);
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 17 }, (_, index) => ownerBase + index)),
    );
    const constants = packConstants([
      [],
      [1],
      [],
      [],
      [],
      [],
      [],
      [2, 10, 20, 30],
      [],
      [],
      [],
      [],
      [],
      [],
      [],
      [3],
      [4, 5],
    ]);
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths);
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 17 }, (_, index) => ownerBase + index)));

    expect(readScalarValue(kernel, ownerBase, pooledStrings)).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(readScalarValue(kernel, ownerBase + 1, pooledStrings)).toEqual({
      tag: ValueTag.String,
      value: "{1, 2;3, 4}",
      stringId: 0,
    });
    expect(readScalarValue(kernel, ownerBase + 2, pooledStrings)).toEqual({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(readScalarValue(kernel, ownerBase + 3, pooledStrings)).toEqual({
      tag: ValueTag.Number,
      value: 2,
    });
    expect(readSpillValues(kernel, ownerBase + 4, pooledStrings)).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 4 },
    ]);
    expect(readSpillValues(kernel, ownerBase + 5, pooledStrings)).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 5 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 4 },
      { tag: ValueTag.Number, value: 6 },
    ]);
    expect(readSpillValues(kernel, ownerBase + 6, pooledStrings)).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 4 },
      { tag: ValueTag.Number, value: 7 },
      { tag: ValueTag.Number, value: 8 },
    ]);
    expect(readScalarValue(kernel, ownerBase + 7, pooledStrings)).toEqual({
      tag: ValueTag.Number,
      value: 20,
    });
    expect(readScalarValue(kernel, ownerBase + 8, pooledStrings)).toEqual({
      tag: ValueTag.String,
      value: "a-2.0-b",
      stringId: 0,
    });
    expect(readScalarValue(kernel, ownerBase + 9, pooledStrings)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.NA,
    });
    expect(readScalarValue(kernel, ownerBase + 10, pooledStrings)).toEqual({
      tag: ValueTag.String,
      value: "fallback",
      stringId: 0,
    });
    expect(readScalarValue(kernel, ownerBase + 11, pooledStrings)).toEqual({
      tag: ValueTag.String,
      value: "na-value",
      stringId: 0,
    });
    expect(readScalarValue(kernel, ownerBase + 12, pooledStrings)).toEqual({
      tag: ValueTag.String,
      value: "x",
      stringId: 0,
    });
    expect(readScalarValue(kernel, ownerBase + 13, pooledStrings)).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(readScalarValue(kernel, ownerBase + 14, pooledStrings)).toEqual({
      tag: ValueTag.Error,
      code: ErrorCode.Value,
    });
    expect(readScalarValue(kernel, ownerBase + 15, pooledStrings)).toEqual({
      tag: ValueTag.Number,
      value: 1,
    });
    expect(readScalarValue(kernel, ownerBase + 16, pooledStrings)).toEqual({
      tag: ValueTag.Number,
      value: 0,
    });
  });
});
