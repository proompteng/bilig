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

describe('wasm kernel numeric and calendar helper seams', () => {
  it('keeps core numeric helper behavior stable', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 5, 4, 4, 4)
    kernel.writeCells(
      new Uint8Array([ValueTag.Number, ValueTag.Number, ValueTag.Number, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Float64Array([3.2, 8, 12, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(16),
      new Uint16Array(16),
    )
    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 2]), Uint32Array.from([0]), Uint32Array.from([3]))
    kernel.uploadRangeShapes(Uint32Array.from([3]), Uint32Array.from([1]))

    const packed = packPrograms([
      [encodePushNumber(0), encodeCall(BuiltinId.Fact, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Factdouble, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BuiltinId.Gcd, 1), encodeRet()],
      [encodePushRange(0), encodeCall(BuiltinId.Lcm, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Even, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Odd, 1), encodeRet()],
    ])
    const constants = packConstants([[5], [6], [], [], [-3.2], [-3.2]])
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
        cellIndex(1, 5, width),
      ]),
    )
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(1, 0, width),
        cellIndex(1, 1, width),
        cellIndex(1, 2, width),
        cellIndex(1, 3, width),
        cellIndex(1, 4, width),
        cellIndex(1, 5, width),
      ]),
    )

    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBe(120)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(48)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(1)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(24)
    expect(kernel.readNumbers()[cellIndex(1, 4, width)]).toBe(-4)
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(-5)
  })

  it('keeps weekend-mask and workday helper behavior stable', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(16, 4, 10, 4, 4)
    kernel.writeCells(new Uint8Array(16), new Float64Array(16), new Uint32Array(16), new Uint16Array(16))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.WorkdayIntl, 3), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodeCall(BuiltinId.WorkdayIntl, 4),
        encodeRet(),
      ],
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.NetworkdaysIntl, 3), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushNumber(2),
        encodePushNumber(3),
        encodeCall(BuiltinId.NetworkdaysIntl, 4),
        encodeRet(),
      ],
    ])
    const constants = packConstants([
      [46094, 1, 7],
      [46094, 2, 7, 46096],
      [46094, 46098, 7],
      [46094, 46098, 7, 46096],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]),
    )
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]))

    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBe(46097)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBe(46099)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(3)
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBe(2)
  })
})
