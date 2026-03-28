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
  return rawValue as ErrorCode;
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
    if (tag == ValueTag.Number) {
      return { tag, value: rawValue };
    }
    if (tag == ValueTag.String) {
      return {
        tag,
        value: decodeStringValue(rawValue, pooledStrings, outputStrings),
        stringId: 0,
      };
    }
    if (tag == ValueTag.Error) {
      return { tag, code: toErrorCode(rawValue) };
    }
    return { tag: ValueTag.Empty };
  });
}

describe("wasm kernel builtin helper seams", () => {
  it("keeps string, spill, and coercion helper behavior stable", async () => {
    const kernel = await createKernel();
    const width = 8;
    kernel.init(32, 7, 16, 1, 1);
    const pooledStrings = ["fallback", "na-value", "Sheet 1", "x"];
    uploadPooledStrings(kernel, pooledStrings);
    kernel.writeCells(
      new Uint8Array(32),
      new Float64Array(32),
      new Uint32Array(32),
      new Uint16Array(32),
    );

    const packed = packPrograms([
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodeCall(BuiltinId.Sequence, 4),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushBoolean(true),
        encodePushString(2),
        encodeCall(BuiltinId.Address, 5),
        encodeRet(),
      ],
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
      [
        encodePushBoolean(true),
        encodePushBoolean(false),
        encodePushBoolean(true),
        encodeCall(BuiltinId.Xor, 3),
        encodeRet(),
      ],
      [
        encodePushBoolean(true),
        encodePushBoolean(true),
        encodePushNumber(0),
        encodeCall(BuiltinId.And, 3),
        encodeRet(),
      ],
      [
        encodePushBoolean(false),
        encodePushBoolean(false),
        encodePushNumber(0),
        encodeCall(BuiltinId.Or, 3),
        encodeRet(),
      ],
    ]);
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([width, width + 1, width + 2, width + 3, width + 4, width + 5, width + 6]),
    );

    const constants = packConstants([[2, 3, 10, 2], [2, 3, 1], [], [], [], [0], [5]]);
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths);
    kernel.evalBatch(
      Uint32Array.from([width, width + 1, width + 2, width + 3, width + 4, width + 5, width + 6]),
    );

    expect(readSpillValues(kernel, width, pooledStrings)).toEqual([
      { tag: ValueTag.Number, value: 10 },
      { tag: ValueTag.Number, value: 12 },
      { tag: ValueTag.Number, value: 14 },
      { tag: ValueTag.Number, value: 16 },
      { tag: ValueTag.Number, value: 18 },
      { tag: ValueTag.Number, value: 20 },
    ]);

    expect(kernel.readTags()[width + 1]).toBe(ValueTag.String);
    expect(
      decodeStringValue(
        kernel.readStringIds()[width + 1] ?? 0,
        pooledStrings,
        kernel.readOutputStrings(),
      ),
    ).toBe("'Sheet 1'!$C$2");

    expect(kernel.readTags()[width + 2]).toBe(ValueTag.String);
    expect(
      decodeStringValue(
        kernel.readStringIds()[width + 2] ?? 0,
        pooledStrings,
        kernel.readOutputStrings(),
      ),
    ).toBe("fallback");

    expect(kernel.readTags()[width + 3]).toBe(ValueTag.String);
    expect(
      decodeStringValue(
        kernel.readStringIds()[width + 3] ?? 0,
        pooledStrings,
        kernel.readOutputStrings(),
      ),
    ).toBe("na-value");

    expect(kernel.readTags()[width + 4]).toBe(ValueTag.Boolean);
    expect(kernel.readNumbers()[width + 4]).toBe(0);
    expect(kernel.readTags()[width + 5]).toBe(ValueTag.Boolean);
    expect(kernel.readNumbers()[width + 5]).toBe(0);
    expect(kernel.readTags()[width + 6]).toBe(ValueTag.Boolean);
    expect(kernel.readNumbers()[width + 6]).toBe(1);
  });

  it("keeps statistical argument collection and error propagation stable", async () => {
    const kernel = await createKernel();
    const width = 8;
    kernel.init(24, 3, 4, 1, 1);
    const pooledStrings = ["x"];
    uploadPooledStrings(kernel, pooledStrings);
    kernel.writeCells(
      new Uint8Array(24),
      new Float64Array(24),
      new Uint32Array(24),
      new Uint16Array(24),
    );

    const packed = packPrograms([
      [
        encodePushNumber(0),
        encodePushBoolean(true),
        encodePushString(0),
        encodeCall(BuiltinId.Stdeva, 3),
        encodeRet(),
      ],
      [
        encodePushError(ErrorCode.Ref),
        encodePushNumber(1),
        encodeCall(BuiltinId.Stdeva, 2),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushBoolean(true),
        encodePushString(0),
        encodeCall(BuiltinId.Vara, 3),
        encodeRet(),
      ],
    ]);
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([width, width + 1, width + 2]),
    );
    const constants = packConstants([[2], [1], [2]]);
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths);
    kernel.evalBatch(Uint32Array.from([width, width + 1, width + 2]));

    expect(kernel.readTags()[width]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[width]).toBeCloseTo(1, 12);

    expect(kernel.readTags()[width + 1]).toBe(ValueTag.Error);
    expect(kernel.readErrors()[width + 1]).toBe(ErrorCode.Ref);

    expect(kernel.readTags()[width + 2]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[width + 2]).toBeCloseTo(1, 12);
  });
});
