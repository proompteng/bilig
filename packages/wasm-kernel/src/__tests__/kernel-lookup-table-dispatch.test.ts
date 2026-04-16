import { describe, expect, it } from 'vitest'
import { BuiltinId, Opcode, ValueTag } from '@bilig/protocol'
import { createKernel } from '../index.js'

function encodeCall(builtinId: number, argc: number): number {
  return (Opcode.CallBuiltin << 24) | ((builtinId << 8) | argc)
}

function encodePushNumber(constantIndex: number): number {
  return (Opcode.PushNumber << 24) | constantIndex
}

function encodePushRange(rangeIndex: number): number {
  return (Opcode.PushRange << 24) | rangeIndex
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

describe('wasm kernel lookup table dispatch', () => {
  it('keeps index, vlookup, and hlookup stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 16
    kernel.init(64, 0, 8, 4, 8)

    const cellTags = new Uint8Array(64)
    const cellNumbers = new Float64Array(64)
    const values = [1, 10, 2, 20, 5, 6, 7, 1, 10, 2, 20, 3, 30, 1, 2, 3, 10, 20, 30]
    values.forEach((value, index) => {
      cellTags[index] = ValueTag.Number
      cellNumbers[index] = value
    })
    kernel.writeCells(cellTags, cellNumbers, new Uint32Array(64), new Uint16Array(64))

    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18]),
      Uint32Array.from([0, 4, 7, 13]),
      Uint32Array.from([4, 3, 6, 6]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([2, 1, 3, 2]), Uint32Array.from([2, 3, 2, 3]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.Index, 3), encodeRet()],
      [encodePushRange(1), encodePushNumber(0), encodeCall(BuiltinId.Index, 2), encodeRet()],
      [
        encodePushNumber(0),
        encodePushRange(2),
        encodePushNumber(1),
        encodePushBoolean(false),
        encodeCall(BuiltinId.Vlookup, 4),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushRange(2),
        encodePushNumber(1),
        encodePushBoolean(true),
        encodeCall(BuiltinId.Vlookup, 4),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushRange(3),
        encodePushNumber(1),
        encodePushBoolean(false),
        encodeCall(BuiltinId.Hlookup, 4),
        encodeRet(),
      ],
      [
        encodePushNumber(0),
        encodePushRange(3),
        encodePushNumber(1),
        encodePushBoolean(true),
        encodeCall(BuiltinId.Hlookup, 4),
        encodeRet(),
      ],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from(Array.from({ length: 6 }, (_, index) => cellIndex(1, index, width))),
    )
    const constants = packConstants([[2, 2], [2], [2, 2], [2.5, 2], [2, 2], [2.5, 2]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from(Array.from({ length: 6 }, (_, index) => cellIndex(1, index, width))))

    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBe(20)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(6)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(20)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(20)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(6)
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(6)
  })
})
