import { describe, expect, it } from "vitest";
import { BuiltinId, ErrorCode, Opcode, ValueTag, type CellValue } from "@bilig/protocol";
import { createKernel, type KernelInstance } from "../index.js";

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

function decodeErrorCode(rawCode: number): ErrorCode {
  switch (rawCode) {
    case 1:
      return ErrorCode.Null;
    case 2:
      return ErrorCode.Div0;
    case 3:
      return ErrorCode.Value;
    case 4:
      return ErrorCode.Ref;
    case 5:
      return ErrorCode.Name;
    case 6:
      return ErrorCode.Num;
    case 7:
      return ErrorCode.NA;
    case 8:
      return ErrorCode.Blocked;
    default:
      throw new Error(`Unexpected error code: ${rawCode}`);
  }
}

function readSpillValues(kernel: KernelInstance, ownerCellIndex: number): CellValue[] {
  const offset = kernel.readSpillOffsets()[ownerCellIndex] ?? 0;
  const length = kernel.readSpillLengths()[ownerCellIndex] ?? 0;
  const tags = kernel.readSpillTags();
  const values = kernel.readSpillNumbers();
  return Array.from({ length }, (_, index) => {
    const tag = tags[offset + index] ?? ValueTag.Empty;
    const rawValue = values[offset + index] ?? 0;
    if (tag == ValueTag.Number) {
      return { tag, value: rawValue };
    }
    if (tag == ValueTag.Empty) {
      return { tag };
    }
    if (tag == ValueTag.Error) {
      return { tag, code: decodeErrorCode(rawValue) };
    }
    throw new Error(`Unexpected spill tag: ${tag}`);
  });
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

function expectNumberSpill(
  kernel: KernelInstance,
  ownerCellIndex: number,
  expected: readonly number[],
  digits = 12,
): void {
  const spill = readSpillValues(kernel, ownerCellIndex);
  expect(spill).toHaveLength(expected.length);
  for (let index = 0; index < expected.length; index += 1) {
    const entry = spill[index];
    expect(entry).toMatchObject({ tag: ValueTag.Number });
    if (!entry || !("value" in entry) || typeof entry.value != "number") {
      throw new Error("Expected numeric spill entry");
    }
    expect(entry.value).toBeCloseTo(expected[index] ?? 0, digits);
  }
}

describe("wasm kernel regression dispatch slab", () => {
  it("keeps paired regression and fitted-array builtins stable across refactors", async () => {
    const kernel = await createKernel();
    kernel.init(80, 4, 0, 9, 24);

    const cellTags = new Uint8Array(80);
    const cellNumbers = new Float64Array(80);
    [3, 5, 7].forEach((value, index) => {
      cellTags[index] = ValueTag.Number;
      cellNumbers[index] = value;
    });
    [1, 2, 3].forEach((value, index) => {
      cellTags[3 + index] = ValueTag.Number;
      cellNumbers[3 + index] = value;
    });
    [4, 5].forEach((value, index) => {
      cellTags[6 + index] = ValueTag.Number;
      cellNumbers[6 + index] = value;
    });
    [6, 18, 54].forEach((value, index) => {
      cellTags[8 + index] = ValueTag.Number;
      cellNumbers[8 + index] = value;
    });
    [1, 2, 3].forEach((value, index) => {
      cellTags[11 + index] = ValueTag.Number;
      cellNumbers[11 + index] = value;
    });
    [4, 5].forEach((value, index) => {
      cellTags[14 + index] = ValueTag.Number;
      cellNumbers[14 + index] = value;
    });

    kernel.writeCells(cellTags, cellNumbers, new Uint32Array(80), new Uint16Array(80));
    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]),
      Uint32Array.from([0, 3, 6, 8, 11, 14]),
      Uint32Array.from([3, 3, 2, 3, 3, 2]),
    );
    kernel.uploadRangeShapes(
      Uint32Array.from([3, 3, 2, 3, 3, 2]),
      Uint32Array.from([1, 1, 1, 1, 1, 1]),
    );

    const packed = packPrograms([
      [encodePushRange(0), encodePushRange(1), encodeCall(BuiltinId.Correl, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BuiltinId.CovarianceP, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BuiltinId.Slope, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BuiltinId.Intercept, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BuiltinId.Rsq, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BuiltinId.Steyx, 2), encodeRet()],
      [
        encodePushNumber(0),
        encodePushRange(0),
        encodePushRange(1),
        encodeCall(BuiltinId.Forecast, 3),
        encodeRet(),
      ],
      [
        encodePushRange(0),
        encodePushRange(1),
        encodePushRange(2),
        encodeCall(BuiltinId.Trend, 3),
        encodeRet(),
      ],
      [
        encodePushRange(3),
        encodePushRange(4),
        encodePushRange(5),
        encodeCall(BuiltinId.Growth, 3),
        encodeRet(),
      ],
      [encodePushRange(0), encodePushRange(1), encodeCall(BuiltinId.Linest, 2), encodeRet()],
      [encodePushRange(3), encodePushRange(4), encodeCall(BuiltinId.Logest, 2), encodeRet()],
    ]);
    const constants = packConstants([[4], [], [], [], [], [], [4], [], [], [], []]);
    const outputCells = Uint32Array.from([16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26]);

    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells);
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths);
    kernel.evalBatch(outputCells);

    expectNumberCell(kernel, 16, 1);
    expectNumberCell(kernel, 17, 4 / 3, 12);
    expectNumberCell(kernel, 18, 2);
    expectNumberCell(kernel, 19, 1);
    expectNumberCell(kernel, 20, 1);
    expectNumberCell(kernel, 21, 0);
    expectNumberCell(kernel, 22, 9);
    expectNumberSpill(kernel, 23, [9, 11]);
    expectNumberSpill(kernel, 24, [162, 486]);
    expectNumberSpill(kernel, 25, [2, 1]);
    expectNumberSpill(kernel, 26, [3, 2]);
  });
});
