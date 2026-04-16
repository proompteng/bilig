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

function expectNumberCell(kernel: Awaited<ReturnType<typeof createKernel>>, index: number, expected: number, digits = 12): void {
  expect(kernel.readTags()[index]).toBe(ValueTag.Number)
  expect(kernel.readNumbers()[index]).toBeCloseTo(expected, digits)
}

describe('wasm kernel statistical test helpers', () => {
  it('keeps paired regression helpers stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(40, 4, 0, 1, 4)

    const cellTags = new Uint8Array(40)
    const cellNumbers = new Float64Array(40)
    ;[5, 8, 11].forEach((value, index) => {
      cellTags[cellIndex(index, 0, width)] = ValueTag.Number
      cellNumbers[cellIndex(index, 0, width)] = value
    })
    ;[1, 2, 3].forEach((value, index) => {
      cellTags[cellIndex(index, 1, width)] = ValueTag.Number
      cellNumbers[cellIndex(index, 1, width)] = value
    })

    kernel.writeCells(cellTags, cellNumbers, new Uint32Array(40), new Uint16Array(40))
    kernel.uploadRangeMembers(
      Uint32Array.from([
        cellIndex(0, 0, width),
        cellIndex(1, 0, width),
        cellIndex(2, 0, width),
        cellIndex(0, 1, width),
        cellIndex(1, 1, width),
        cellIndex(2, 1, width),
      ]),
      Uint32Array.from([0, 3]),
      Uint32Array.from([3, 3]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([3, 3]), Uint32Array.from([1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushRange(1), encodeCall(BuiltinId.Intercept, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BuiltinId.Slope, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BuiltinId.Rsq, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BuiltinId.Steyx, 2), encodeRet()],
    ])
    const outputCells = Uint32Array.from([cellIndex(3, 0, width), cellIndex(3, 1, width), cellIndex(3, 2, width), cellIndex(3, 3, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(new Float64Array(), new Uint32Array([0, 0, 0, 0]), new Uint32Array([0, 0, 0, 0]))
    kernel.evalBatch(outputCells)

    expectNumberCell(kernel, outputCells[0], 2)
    expectNumberCell(kernel, outputCells[1], 3)
    expectNumberCell(kernel, outputCells[2], 1)
    expectNumberCell(kernel, outputCells[3], 0)
  })

  it('keeps chi/f/z hypothesis helpers stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 12
    kernel.init(36, 8, 3, 4, 12)

    const cellTags = new Uint8Array(36)
    const cellNumbers = new Float64Array(36)
    ;[58, 35, 11, 25, 10, 23, 45.35, 47.65, 17.56, 18.44, 16.09, 16.91, 6, 7, 9, 15, 21, 20, 28, 31, 38, 40, 1, 2, 3, 4, 5].forEach(
      (value, index) => {
        cellTags[index] = ValueTag.Number
        cellNumbers[index] = value
      },
    )

    kernel.writeCells(cellTags, cellNumbers, new Uint32Array(36), new Uint16Array(36))
    kernel.uploadRangeMembers(
      Uint32Array.from([0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26]),
      Uint32Array.from([0, 6, 12, 17, 22]),
      Uint32Array.from([6, 6, 5, 5, 5]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([3, 3, 5, 5, 5]), Uint32Array.from([2, 2, 1, 1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushRange(1), encodeCall(BuiltinId.ChisqTest, 2), encodeRet()],
      [encodePushRange(2), encodePushRange(3), encodeCall(BuiltinId.Ftest, 2), encodeRet()],
      [encodePushRange(4), encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.ZTest, 3), encodeRet()],
    ])
    const outputCells = Uint32Array.from([cellIndex(3, 0, width), cellIndex(3, 1, width), cellIndex(3, 2, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    const constants = packConstants([[], [], [2, 1]])
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    expectNumberCell(kernel, outputCells[0], 0.0003082, 7)
    expectNumberCell(kernel, outputCells[1], 0.648317846786175, 12)
    expectNumberCell(kernel, outputCells[2], 0.012673617875446075, 12)
  })

  it('keeps t-test helpers stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 6
    kernel.init(width, 4, 1, 2, 2)
    kernel.writeCells(
      new Uint8Array([
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        ValueTag.Number,
        ValueTag.Number,
        0,
        0,
        0,
        0,
        ValueTag.Number,
        ValueTag.Number,
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
      new Float64Array([1, 1, 0, 0, 0, 0, 2, 3, 0, 0, 0, 0, 4, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]),
      new Uint32Array(width * 4),
      new Uint16Array(width * 4),
    )
    kernel.uploadRangeMembers(
      Uint32Array.from([
        cellIndex(0, 0, width),
        cellIndex(1, 0, width),
        cellIndex(2, 0, width),
        cellIndex(0, 1, width),
        cellIndex(1, 1, width),
        cellIndex(2, 1, width),
      ]),
      Uint32Array.from([0, 3]),
      Uint32Array.from([3, 3]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([3, 3]), Uint32Array.from([1, 1]))

    const packed = packPrograms([
      [encodePushRange(0), encodePushRange(1), encodePushNumber(0), encodePushNumber(1), encodeCall(BuiltinId.TTest, 4), encodeRet()],
    ])
    const constants = packConstants([[2, 1]])
    const outputCells = Uint32Array.from([cellIndex(3, 0, width)])
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, outputCells)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(outputCells)

    expectNumberCell(kernel, outputCells[0], 1, 12)
  })
})
