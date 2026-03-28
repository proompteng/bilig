import { describe, expect, it } from "vitest";
import { BuiltinId, Opcode, ValueTag, type CellValue } from "@bilig/protocol";
import { createKernel, type KernelInstance } from "../index.js";

function asciiCodes(text: string): Uint16Array {
  return Uint16Array.from(Array.from(text, (char) => char.charCodeAt(0)));
}

function encodeCall(builtinId: number, argc: number): number {
  return (Opcode.CallBuiltin << 24) | ((builtinId << 8) | argc);
}

function encodePushCell(cellOffset: number): number {
  return (Opcode.PushCell << 24) | cellOffset;
}

function encodePushNumber(constantIndex: number): number {
  return (Opcode.PushNumber << 24) | constantIndex;
}

function encodePushString(stringId: number): number {
  return (Opcode.PushString << 24) | stringId;
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
  const outputStrings = kernel.readOutputStrings();
  return Array.from({ length }, (_, index) => {
    const tag = tags[offset + index] ?? ValueTag.Empty;
    const rawValue = values[offset + index] ?? 0;
    if (tag == ValueTag.String) {
      return {
        tag,
        value: outputStrings[rawValue - 2147483648] ?? "",
        stringId: 0,
      };
    }
    if (tag == ValueTag.Number) {
      return { tag, value: rawValue };
    }
    return { tag: ValueTag.Empty };
  });
}

describe("wasm kernel text helper seams", () => {
  it("keeps text replacement, search-mode, and split helpers stable", async () => {
    const kernel = await createKernel();
    const width = 10;
    kernel.init(32, 7, 4, 1, 1);
    kernel.uploadStrings(
      Uint32Array.from([0, 0, 8, 9, 15, 17, 19, 21, 37, 38]),
      Uint32Array.from([0, 8, 1, 6, 2, 2, 2, 16, 1, 1]),
      asciiCodes("alphabetZbananaanooxoalpha-beta-gamma-|"),
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
        0,
        0,
        0,
        0,
        0,
        0,
        0,
        0,
      ]),
      new Float64Array(32),
      new Uint32Array([
        1, 3, 6, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
        0,
      ]),
      new Uint16Array(32),
    );

    const packed = packPrograms([
      [
        encodePushCell(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushString(2),
        encodeCall(BuiltinId.Replace, 4),
        encodeRet(),
      ],
      [
        encodePushCell(1),
        encodePushString(4),
        encodePushString(5),
        encodePushNumber(0),
        encodeCall(BuiltinId.Substitute, 4),
        encodeRet(),
      ],
      [encodePushCell(2), encodePushNumber(0), encodeCall(BuiltinId.Rept, 2), encodeRet()],
      [encodePushString(7), encodePushString(8), encodeCall(BuiltinId.Textbefore, 2), encodeRet()],
      [
        encodePushString(7),
        encodePushString(8),
        encodePushNumber(0),
        encodeCall(BuiltinId.Textafter, 3),
        encodeRet(),
      ],
      [
        encodePushString(7),
        encodePushString(8),
        encodePushString(9),
        encodeCall(BuiltinId.Textsplit, 3),
        encodeRet(),
      ],
    ]);
    const constants = packConstants([[3, 2], [2], [3], [], [-1], []]);
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
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths);
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

    expect(kernel.readOutputStrings()).toEqual([
      "alZabet",
      "banooa",
      "xoxoxo",
      "alpha",
      "gamma",
      "alpha",
      "beta",
      "gamma",
    ]);
    expect(readSpillValues(kernel, cellIndex(1, 6, width))).toEqual([
      { tag: ValueTag.String, value: "alpha", stringId: 0 },
      { tag: ValueTag.String, value: "beta", stringId: 0 },
      { tag: ValueTag.String, value: "gamma", stringId: 0 },
    ]);
  });
});
