import { describe, expect, it } from "vitest";
import { BuiltinId, Opcode, ValueTag, type CellValue } from "@bilig/protocol";
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

function cellIndex(row: number, col: number, width: number): number {
  return row * width + col;
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
    throw new Error(`Unexpected spill tag: ${tag}`);
  });
}

describe("wasm kernel array dispatch slab", () => {
  it("keeps sequence, filter, axis aggregates, folds, and makearray stable", async () => {
    const kernel = await createKernel();
    const width = 10;
    kernel.init(100, 6, 8, 8, 32);

    const cellTags = new Uint8Array(80);
    const cellNumbers = new Float64Array(80);
    [1, 2, 3, 4, 5, 6, 1, 0, 1].forEach((value, index) => {
      cellTags[index] = ValueTag.Number;
      cellNumbers[index] = value;
    });
    kernel.writeCells(cellTags, cellNumbers, new Uint32Array(80), new Uint16Array(80));

    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 2, 3, 4, 5, 6, 7, 8]),
      Uint32Array.from([0, 6]),
      Uint32Array.from([6, 3]),
    );
    kernel.uploadRangeShapes(Uint32Array.from([2, 1]), Uint32Array.from([3, 3]));

    const packed = packPrograms([
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodeCall(BuiltinId.Sequence, 4),
        encodeRet(),
      ],
      [encodePushRange(0), encodePushRange(1), encodeCall(BuiltinId.Filter, 2), encodeRet()],
      [
        encodePushNumber(0),
        encodePushRange(0),
        encodeCall(BuiltinId.ByrowAggregate, 2),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushRange(0), encodeCall(BuiltinId.ReduceSum, 2), encodeRet()],
      [encodePushNumber(0), encodePushRange(0), encodeCall(BuiltinId.ScanSum, 2), encodeRet()],
    ]);
    const constants = packConstants([[2, 3, 10, 2], [], [2], [0], [0]]);
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(3, 0, width),
        cellIndex(3, 5, width),
        cellIndex(5, 0, width),
        cellIndex(5, 5, width),
        cellIndex(7, 0, width),
      ]),
    );
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths);
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(3, 0, width),
        cellIndex(3, 5, width),
        cellIndex(5, 0, width),
        cellIndex(5, 5, width),
        cellIndex(7, 0, width),
      ]),
    );

    expect(readSpillValues(kernel, cellIndex(3, 0, width))).toEqual([
      { tag: ValueTag.Number, value: 10 },
      { tag: ValueTag.Number, value: 12 },
      { tag: ValueTag.Number, value: 14 },
      { tag: ValueTag.Number, value: 16 },
      { tag: ValueTag.Number, value: 18 },
      { tag: ValueTag.Number, value: 20 },
    ]);
    expect(readSpillValues(kernel, cellIndex(3, 5, width))).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 4 },
      { tag: ValueTag.Number, value: 6 },
    ]);
    expect(readSpillValues(kernel, cellIndex(5, 0, width))).toEqual([
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 5 },
    ]);

    expect(kernel.readTags()[cellIndex(5, 5, width)]).toBe(ValueTag.Number);
    expect(kernel.readNumbers()[cellIndex(5, 5, width)]).toBe(21);

    expect(readSpillValues(kernel, cellIndex(7, 0, width))).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 6 },
      { tag: ValueTag.Number, value: 10 },
      { tag: ValueTag.Number, value: 15 },
      { tag: ValueTag.Number, value: 21 },
    ]);
  });
});
