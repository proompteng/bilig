import { describe, expect, it } from "vitest";
import { BuiltinId, ErrorCode, Opcode, ValueTag } from "@bilig/protocol";
import { createKernel } from "../index.js";

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

describe("wasm kernel convert helpers", () => {
  it("keeps unit conversion behavior stable across refactors", async () => {
    const kernel = await createKernel();
    const width = 8;
    kernel.init(24, 5, 5, 1, 1);
    const strings = ["m", "ft", "km", "m", "F", "C", "m", "s", "km", "mi"];
    const offsets: number[] = [];
    const lengths: number[] = [];
    const data: number[] = [];
    let cursor = 0;
    for (const value of strings) {
      offsets.push(cursor);
      lengths.push(value.length);
      data.push(...Array.from(value, (char) => char.charCodeAt(0)));
      cursor += value.length;
    }
    kernel.uploadStrings(
      Uint32Array.from(offsets),
      Uint32Array.from(lengths),
      Uint16Array.from(data),
    );
    kernel.writeCells(
      new Uint8Array(24),
      new Float64Array(24),
      new Uint32Array(24),
      new Uint16Array(24),
    );

    const packed = packPrograms([
      [
        encodePushNumber(0),
        encodePushString(0),
        encodePushString(1),
        encodeCall(BuiltinId.Convert, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushString(2),
        encodePushString(3),
        encodeCall(BuiltinId.Convert, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushString(4),
        encodePushString(5),
        encodeCall(BuiltinId.Convert, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushString(6),
        encodePushString(7),
        encodeCall(BuiltinId.Convert, 3),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushString(8),
        encodePushString(9),
        encodeCall(BuiltinId.Convert, 3),
        encodeRet(),
      ],
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
    const constants = packConstants([[1], [1], [212], [1], [1]]);
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

    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBeCloseTo(3.280839895013, 12);
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(1000);
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBeCloseTo(100, 12);
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.Error);
    expect(kernel.readErrors()[cellIndex(1, 3, width)]).toBe(ErrorCode.NA);
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBeCloseTo(0.621371192237, 12);
  });
});
