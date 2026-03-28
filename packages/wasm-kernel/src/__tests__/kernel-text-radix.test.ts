import { describe, expect, it } from "vitest";
import { BuiltinId, Opcode, ValueTag } from "@bilig/protocol";
import { createKernel } from "../index.js";

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

function readStringCell(
  kernel: Awaited<ReturnType<typeof createKernel>>,
  index: number,
  pooledStrings: readonly string[],
): string {
  expect(kernel.readTags()[index]).toBe(ValueTag.String);
  const raw = kernel.readStringIds()[index] ?? 0;
  const outputIndex = raw >= OUTPUT_STRING_BASE ? raw - OUTPUT_STRING_BASE : -1;
  return outputIndex >= 0
    ? (kernel.readOutputStrings()[outputIndex] ?? "")
    : (pooledStrings[raw] ?? "");
}

describe("wasm kernel text/radix helpers", () => {
  it("keeps byte-oriented text helpers stable across refactors", async () => {
    const kernel = await createKernel();
    const width = 10;
    kernel.init(20, 8, 4, 1, 1);
    const pooledStrings = ["abcdef", "alphabet", "ph", "d", "Z", "é"];
    kernel.uploadStrings(
      Uint32Array.from([0, 6, 14, 16, 17, 18]),
      Uint32Array.from([6, 8, 2, 1, 1, 1]),
      Uint16Array.from(Array.from("abcdefalphabetphdZé", (char) => char.charCodeAt(0))),
    );
    kernel.writeCells(
      new Uint8Array(20),
      new Float64Array(20),
      new Uint32Array(20),
      new Uint16Array(20),
    );

    const packed = packPrograms([
      [encodePushString(5), encodeCall(BuiltinId.Lenb, 1), encodeRet()],
      [encodePushString(0), encodePushNumber(0), encodeCall(BuiltinId.Leftb, 2), encodeRet()],
      [
        encodePushString(0),
        encodePushNumber(0),
        encodePushNumber(1),
        encodeCall(BuiltinId.Midb, 3),
        encodeRet(),
      ],
      [encodePushString(0), encodePushNumber(0), encodeCall(BuiltinId.Rightb, 2), encodeRet()],
      [
        encodePushString(3),
        encodePushString(0),
        encodePushNumber(0),
        encodeCall(BuiltinId.Findb, 3),
        encodeRet(),
      ],
      [encodePushString(2), encodePushString(1), encodeCall(BuiltinId.Searchb, 2), encodeRet()],
      [
        encodePushString(1),
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushString(4),
        encodeCall(BuiltinId.Replaceb, 4),
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
    const constants = packConstants([[], [2], [3, 2], [3], [3], [], [3, 2]]);
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths);
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

    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(2);
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(4);
    expect(kernel.readNumbers()[cellIndex(1, 6, width)]).toBe(3);
    expect(readStringCell(kernel, cellIndex(1, 2, width), pooledStrings)).toBe("ab");
    expect(readStringCell(kernel, cellIndex(1, 3, width), pooledStrings)).toBe("cd");
    expect(readStringCell(kernel, cellIndex(1, 4, width), pooledStrings)).toBe("def");
    expect(readStringCell(kernel, cellIndex(1, 7, width), pooledStrings)).toBe("alZabet");
  });

  it("keeps radix conversion helpers stable across refactors", async () => {
    const kernel = await createKernel();
    const width = 8;
    kernel.init(16, 5, 2, 1, 1);
    const pooledStrings = ["1010", "1F"];
    kernel.uploadStrings(
      Uint32Array.from([0, 4]),
      Uint32Array.from([4, 2]),
      Uint16Array.from(Array.from("10101F", (char) => char.charCodeAt(0))),
    );
    kernel.writeCells(
      new Uint8Array(16),
      new Float64Array(16),
      new Uint32Array(16),
      new Uint16Array(16),
    );

    const packed = packPrograms([
      [encodePushString(0), encodeCall(BuiltinId.Bin2dec, 1), encodeRet()],
      [encodePushString(1), encodeCall(BuiltinId.Hex2dec, 1), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Base, 3),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Dec2hex, 2), encodeRet()],
      [encodePushString(1), encodePushNumber(0), encodeCall(BuiltinId.Decimal, 2), encodeRet()],
    ]);
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
      ]),
    );
    const constants = packConstants([[255], [], [255, 16, 4], [255, 4], [16]]);
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths);
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
      ]),
    );

    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBe(10);
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(31);
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(31);
    expect(readStringCell(kernel, cellIndex(1, 2, width), pooledStrings)).toBe("00FF");
    expect(readStringCell(kernel, cellIndex(1, 3, width), pooledStrings)).toBe("00FF");
  });
});
