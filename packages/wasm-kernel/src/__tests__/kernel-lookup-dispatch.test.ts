import { describe, expect, it } from 'vitest'
import { BuiltinId, ErrorCode, Opcode, ValueTag } from '@bilig/protocol'
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

function encodeBinary(opcode: Opcode): number {
  return opcode << 24
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

function expectNumberCell(kernel: Awaited<ReturnType<typeof createKernel>>, index: number, expected: number, digits = 12): void {
  expect(kernel.readTags()[index]).toBe(ValueTag.Number)
  expect(kernel.readNumbers()[index]).toBeCloseTo(expected, digits)
}

function expectErrorCell(kernel: Awaited<ReturnType<typeof createKernel>>, index: number, expected: ErrorCode): void {
  expect(kernel.readTags()[index]).toBe(ValueTag.Error)
  expect(kernel.readErrors()[index]).toBe(expected)
}

describe('wasm kernel lookup dispatch slab', () => {
  it('keeps match-family lookup behavior stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 12
    kernel.init(120, 1, 6, 8, 40)

    const cellTags = new Uint8Array(120)
    const cellNumbers = new Float64Array(120)
    ;[10, 20, 20, 30].forEach((value, index) => {
      cellTags[index] = ValueTag.Number
      cellNumbers[index] = value
    })
    ;[100, 200, 250, 300].forEach((value, index) => {
      cellTags[10 + index] = ValueTag.Number
      cellNumbers[10 + index] = value
    })
    ;[10, 20, 30].forEach((value, index) => {
      cellTags[20 + index] = ValueTag.Number
      cellNumbers[20 + index] = value
    })
    ;[1, 2, 3].forEach((value, index) => {
      cellTags[30 + index] = ValueTag.Number
      cellNumbers[30 + index] = value
    })
    cellTags[40] = ValueTag.Number
    cellNumbers[40] = 2374.28
    cellTags[50] = ValueTag.Number
    cellNumbers[50] = 700
    kernel.writeCells(cellTags, cellNumbers, new Uint32Array(120), new Uint16Array(120))

    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 2, 3, 10, 11, 12, 13, 20, 21, 22, 30, 31, 32, 40, 50]),
      Uint32Array.from([0, 4, 8, 11, 14, 15]),
      Uint32Array.from([4, 4, 3, 3, 1, 1]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([4, 4, 3, 3, 1, 1]), Uint32Array.from([1, 1, 1, 1, 1, 1]))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushRange(0), encodePushNumber(1), encodeCall(BuiltinId.Match, 3), encodeRet()],
      [encodePushNumber(0), encodePushRange(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Xmatch, 4), encodeRet()],
      [encodePushNumber(0), encodePushRange(0), encodePushRange(1), encodeCall(BuiltinId.Xlookup, 3), encodeRet()],
      [encodePushNumber(0), encodePushRange(0), encodePushRange(1), encodePushNumber(1), encodeCall(BuiltinId.Xlookup, 4), encodeRet()],
      [encodePushNumber(0), encodePushRange(2), encodePushRange(3), encodeCall(BuiltinId.Lookup, 3), encodeRet()],
      [encodePushNumber(0), encodePushRange(4), encodePushRange(5), encodeCall(BuiltinId.Xlookup, 3), encodeRet()],
    ])
    const constants = packConstants([[20, 0], [20, 0, -1], [20], [25, 999], [25], [2374.2799999999997]])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(5, 0, width),
        cellIndex(5, 1, width),
        cellIndex(5, 2, width),
        cellIndex(5, 3, width),
        cellIndex(5, 4, width),
        cellIndex(5, 5, width),
      ]),
    )
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(5, 0, width),
        cellIndex(5, 1, width),
        cellIndex(5, 2, width),
        cellIndex(5, 3, width),
        cellIndex(5, 4, width),
        cellIndex(5, 5, width),
      ]),
    )

    expectNumberCell(kernel, cellIndex(5, 0, width), 2)
    expectNumberCell(kernel, cellIndex(5, 1, width), 3)
    expectNumberCell(kernel, cellIndex(5, 2, width), 200)
    expectNumberCell(kernel, cellIndex(5, 3, width), 999)
    expectNumberCell(kernel, cellIndex(5, 4, width), 2)
    expectNumberCell(kernel, cellIndex(5, 5, width), 700)
  })

  it('skips lookup-vector error sentinels for approximate LOOKUP', async () => {
    const kernel = await createKernel()
    const width = 12
    kernel.init(120, 2, 2, 3, 32)

    const cellTags = new Uint8Array(120)
    const cellNumbers = new Float64Array(120)
    const cellErrors = new Uint16Array(120)
    ;[0, 2, 5].forEach((index) => {
      cellTags[index] = ValueTag.Error
      cellErrors[index] = ErrorCode.Div0
    })
    ;[1, 3, 4].forEach((index) => {
      cellTags[index] = ValueTag.Number
      cellNumbers[index] = 1
    })
    ;[10, 20, 30, 40, 50, 60].forEach((value, index) => {
      cellTags[10 + index] = ValueTag.Number
      cellNumbers[10 + index] = value
    })
    ;[0, 1, 0, 1, 1, 0].forEach((value, index) => {
      cellTags[20 + index] = ValueTag.Number
      cellNumbers[20 + index] = value
    })
    kernel.writeCells(cellTags, cellNumbers, new Uint32Array(120), cellErrors)

    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 2, 3, 4, 5, 10, 11, 12, 13, 14, 15, 20, 21, 22, 23, 24, 25]),
      Uint32Array.from([0, 6, 12]),
      Uint32Array.from([6, 6, 6]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([6, 6, 6]), Uint32Array.from([1, 1, 1]))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushRange(0), encodePushRange(1), encodeCall(BuiltinId.Lookup, 3), encodeRet()],
      [
        encodePushNumber(0),
        encodePushNumber(1),
        encodePushRange(2),
        encodeBinary(Opcode.Div),
        encodePushRange(1),
        encodeCall(BuiltinId.Lookup, 3),
        encodeRet(),
      ],
    ])
    const constants = packConstants([[2], [2, 1]])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(5, 0, width), cellIndex(5, 1, width)]),
    )
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(Uint32Array.from([cellIndex(5, 0, width), cellIndex(5, 1, width)]))

    expectNumberCell(kernel, cellIndex(5, 0, width), 50)
    expectNumberCell(kernel, cellIndex(5, 1, width), 50)
  })

  it('keeps lookup dispatch validation stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 12
    kernel.init(120, 1, 7, 8, 32)

    const cellTags = new Uint8Array(120)
    const cellNumbers = new Float64Array(120)
    ;[10, 20, 30].forEach((value, index) => {
      cellTags[index] = ValueTag.Number
      cellNumbers[index] = value
    })
    ;[100, 200].forEach((value, index) => {
      cellTags[10 + index] = ValueTag.Number
      cellNumbers[10 + index] = value
    })
    ;[100, 200, 300].forEach((value, index) => {
      cellTags[20 + index] = ValueTag.Number
      cellNumbers[20 + index] = value
    })
    ;[1, 2, 3, 4].forEach((value, index) => {
      cellTags[30 + index] = ValueTag.Number
      cellNumbers[30 + index] = value
    })
    kernel.writeCells(cellTags, cellNumbers, new Uint32Array(120), new Uint16Array(120))

    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 2, 10, 11, 20, 21, 22, 30, 31, 32, 33]),
      Uint32Array.from([0, 3, 5, 8]),
      Uint32Array.from([3, 2, 3, 4]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([3, 2, 3, 2]), Uint32Array.from([1, 1, 1, 2]))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushRange(0), encodePushNumber(1), encodeCall(BuiltinId.Match, 3), encodeRet()],
      [encodePushNumber(0), encodePushRange(0), encodePushNumber(1), encodePushNumber(2), encodeCall(BuiltinId.Xmatch, 4), encodeRet()],
      [encodePushNumber(0), encodePushRange(0), encodePushRange(1), encodeCall(BuiltinId.Xlookup, 3), encodeRet()],
      [encodePushNumber(0), encodePushRange(0), encodePushRange(1), encodePushNumber(1), encodeCall(BuiltinId.Lookup, 3), encodeRet()],
      [encodePushNumber(0), encodePushRange(3), encodePushNumber(1), encodeCall(BuiltinId.Match, 3), encodeRet()],
    ])
    const constants = packConstants([[20, 2], [20, 0, 2], [20], [20], [20, 0]])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([
        cellIndex(5, 0, width),
        cellIndex(5, 1, width),
        cellIndex(5, 2, width),
        cellIndex(5, 3, width),
        cellIndex(5, 4, width),
      ]),
    )
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(5, 0, width),
        cellIndex(5, 1, width),
        cellIndex(5, 2, width),
        cellIndex(5, 3, width),
        cellIndex(5, 4, width),
      ]),
    )

    expectErrorCell(kernel, cellIndex(5, 0, width), ErrorCode.Value)
    expectErrorCell(kernel, cellIndex(5, 1, width), ErrorCode.Value)
    expectErrorCell(kernel, cellIndex(5, 2, width), ErrorCode.Value)
    expectErrorCell(kernel, cellIndex(5, 3, width), ErrorCode.Value)
    expectErrorCell(kernel, cellIndex(5, 4, width), ErrorCode.Value)
  })
})
