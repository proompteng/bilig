import { describe, expect, it } from 'vitest'
import { BuiltinId, Opcode, ValueTag } from '@bilig/protocol'
import { createKernel } from '../index.js'

function encodeCall(builtinId: number, argc: number): number {
  return (Opcode.CallBuiltin << 24) | ((builtinId << 8) | argc)
}

function encodePushNumber(constantIndex: number): number {
  return (Opcode.PushNumber << 24) | constantIndex
}

function encodePushString(stringId: number): number {
  return (Opcode.PushString << 24) | stringId
}

function encodePushBoolean(value: boolean): number {
  return (Opcode.PushBoolean << 24) | (value ? 1 : 0)
}

function encodeRet(): number {
  return Opcode.Ret << 24
}

function packPrograms(programs: number[][]): {
  programs: Uint32Array
  offsets: Uint32Array
  lengths: Uint32Array
} {
  const flat: number[] = []
  const offsets: number[] = []
  const lengths: number[] = []
  let offset = 0
  for (const program of programs) {
    offsets.push(offset)
    lengths.push(program.length)
    flat.push(...program)
    offset += program.length
  }
  return {
    programs: Uint32Array.from(flat),
    offsets: Uint32Array.from(offsets),
    lengths: Uint32Array.from(lengths),
  }
}

function packConstants(constantsByProgram: number[][]): {
  constants: Float64Array
  offsets: Uint32Array
  lengths: Uint32Array
} {
  const flat: number[] = []
  const offsets: number[] = []
  const lengths: number[] = []
  let offset = 0
  for (const constants of constantsByProgram) {
    offsets.push(offset)
    lengths.push(constants.length)
    flat.push(...constants)
    offset += constants.length
  }
  return {
    constants: Float64Array.from(flat),
    offsets: Uint32Array.from(offsets),
    lengths: Uint32Array.from(lengths),
  }
}

function cellIndex(row: number, col: number, width: number): number {
  return row * width + col
}

describe('wasm kernel logic and info dispatch', () => {
  it('keeps bitwise, logical, and type-info builtins stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 32
    const pooledStrings = ['a', 'b', 'z', 'x']
    kernel.init(96, pooledStrings.length, 8, 1, 1)
    kernel.uploadStrings(
      Uint32Array.from([0, 1, 2, 3]),
      Uint32Array.from([1, 1, 1, 1]),
      Uint16Array.from(Array.from('abzx', (char) => char.charCodeAt(0))),
    )
    kernel.writeCells(new Uint8Array(96), new Float64Array(96), new Uint32Array(96), new Uint16Array(96))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Bitand, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Bitor, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Bitxor, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Bitlshift, 2), encodeRet()],
      [encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Bitrshift, 2), encodeRet()],
      [encodePushBoolean(true), encodePushBoolean(true), encodePushNumber(0), encodeCall(BuiltinId.And, 3), encodeRet()],
      [encodePushBoolean(false), encodePushNumber(0), encodePushBoolean(true), encodeCall(BuiltinId.Or, 3), encodeRet()],
      [encodePushBoolean(true), encodePushBoolean(false), encodePushBoolean(true), encodeCall(BuiltinId.Xor, 3), encodeRet()],
      [encodePushBoolean(false), encodeCall(BuiltinId.Not, 1), encodeRet()],
      [
        encodePushBoolean(false),
        encodePushNumber(0),
        encodePushBoolean(true),
        encodePushNumber(1),
        encodeCall(BuiltinId.Ifs, 4),
        encodeRet(),
      ],
      [
        encodePushString(1),
        encodePushString(0),
        encodePushNumber(0),
        encodePushString(1),
        encodePushNumber(1),
        encodePushNumber(2),
        encodeCall(BuiltinId.Switch, 6),
        encodeRet(),
      ],
      [encodePushString(2), encodePushString(0), encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Switch, 4), encodeRet()],
      [encodeCall(BuiltinId.IsBlank, 0), encodeRet()],
      [encodePushString(3), encodeCall(BuiltinId.IsBlank, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.IsNumber, 1), encodeRet()],
      [encodePushString(3), encodeCall(BuiltinId.IsNumber, 1), encodeRet()],
      [encodePushString(3), encodeCall(BuiltinId.IsText, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.IsText, 1), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 18 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([
      [6, 3],
      [6, 3],
      [6, 3],
      [3, 2],
      [8, 2],
      [1],
      [0],
      [],
      [],
      [1, 2],
      [1, 2, 9],
      [1, 9],
      [],
      [],
      [5],
      [],
      [],
      [5],
    ])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 18 }, (_, index) => cellIndex(1, index, width))))

    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBe(2)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(7)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(5)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(12)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(2)
    expect(kernel.readTags()[cellIndex(1, 5, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(1)
    expect(kernel.readTags()[cellIndex(1, 6, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 6, width)]).toBe(1)
    expect(kernel.readTags()[cellIndex(1, 7, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 7, width)]).toBe(0)
    expect(kernel.readTags()[cellIndex(1, 8, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 8, width)]).toBe(1)
    expect(kernel.readNumbers()[cellIndex(1, 9, width)]).toBe(2)
    expect(kernel.readNumbers()[cellIndex(1, 10, width)]).toBe(2)
    expect(kernel.readNumbers()[cellIndex(1, 11, width)]).toBe(9)
    expect(kernel.readTags()[cellIndex(1, 12, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 12, width)]).toBe(1)
    expect(kernel.readTags()[cellIndex(1, 13, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 13, width)]).toBe(0)
    expect(kernel.readTags()[cellIndex(1, 14, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 14, width)]).toBe(1)
    expect(kernel.readTags()[cellIndex(1, 15, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 15, width)]).toBe(0)
    expect(kernel.readTags()[cellIndex(1, 16, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 16, width)]).toBe(1)
    expect(kernel.readTags()[cellIndex(1, 17, width)]).toBe(ValueTag.Boolean)
    expect(kernel.readNumbers()[cellIndex(1, 17, width)]).toBe(0)
  })
})
