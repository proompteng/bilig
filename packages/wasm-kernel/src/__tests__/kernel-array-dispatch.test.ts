import { describe, expect, it } from 'vitest'
import { BuiltinId, Opcode, ValueTag, type CellValue } from '@bilig/protocol'
import { createKernel, type KernelInstance } from '../index.js'

const OUTPUT_STRING_BASE = 2147483648

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

function packStrings(values: readonly string[]): {
  data: Uint16Array
  lengths: Uint32Array
  offsets: Uint32Array
} {
  const data: number[] = []
  const lengths: number[] = []
  const offsets: number[] = []
  let offset = 0

  for (const value of values) {
    offsets.push(offset)
    lengths.push(value.length)
    for (let index = 0; index < value.length; index += 1) {
      data.push(value.charCodeAt(index))
    }
    offset += value.length
  }

  return {
    data: Uint16Array.from(data),
    lengths: Uint32Array.from(lengths),
    offsets: Uint32Array.from(offsets),
  }
}

function readSpillValues(kernel: KernelInstance, ownerCellIndex: number, inputStrings: readonly string[] = []): CellValue[] {
  const offset = kernel.readSpillOffsets()[ownerCellIndex] ?? 0
  const length = kernel.readSpillLengths()[ownerCellIndex] ?? 0
  const tags = kernel.readSpillTags()
  const values = kernel.readSpillNumbers()
  const outputStrings = kernel.readOutputStrings()
  return Array.from({ length }, (_, index) => {
    const tag = tags[offset + index] ?? ValueTag.Empty
    const rawValue = values[offset + index] ?? 0
    if (tag === ValueTag.Number) {
      return { tag, value: rawValue }
    }
    if (tag === ValueTag.String) {
      const stringIndex = rawValue >= OUTPUT_STRING_BASE ? rawValue - OUTPUT_STRING_BASE : rawValue
      const value = rawValue >= OUTPUT_STRING_BASE ? (outputStrings[stringIndex] ?? '') : (inputStrings[stringIndex] ?? '')
      return { tag, value, stringId: 0 }
    }
    if (tag === ValueTag.Empty) {
      return { tag }
    }
    throw new Error(`Unexpected spill tag: ${tag}`)
  })
}

describe('wasm kernel array dispatch slab', () => {
  it('keeps sequence, filter, axis aggregates, folds, and makearray stable', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(100, 6, 8, 8, 32)

    const cellTags = new Uint8Array(80)
    const cellNumbers = new Float64Array(80)
    ;[1, 2, 3, 4, 5, 6, 1, 0, 1].forEach((value, index) => {
      cellTags[index] = ValueTag.Number
      cellNumbers[index] = value
    })
    kernel.writeCells(cellTags, cellNumbers, new Uint32Array(80), new Uint16Array(80))

    kernel.uploadRangeMembers(Uint32Array.from([0, 1, 2, 3, 4, 5, 6, 7, 8]), Uint32Array.from([0, 6]), Uint32Array.from([6, 3]))
    kernel.uploadRangeShapes(Uint32Array.from([2, 1]), Uint32Array.from([3, 3]))

    const packed = packPrograms([
      [encodePushNumber(0), encodePushNumber(1), encodePushNumber(2), encodePushNumber(3), encodeCall(BuiltinId.Sequence, 4), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodeCall(BuiltinId.Filter, 2), encodeRet()],
      [encodePushNumber(0), encodePushRange(0), encodeCall(BuiltinId.ByrowAggregate, 2), encodeRet()],
      [encodePushNumber(0), encodePushRange(0), encodeCall(BuiltinId.ReduceSum, 2), encodeRet()],
      [encodePushNumber(0), encodePushRange(0), encodeCall(BuiltinId.ScanSum, 2), encodeRet()],
    ])
    const constants = packConstants([[2, 3, 10, 2], [], [2], [0], [0]])
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
    )
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(
      Uint32Array.from([
        cellIndex(3, 0, width),
        cellIndex(3, 5, width),
        cellIndex(5, 0, width),
        cellIndex(5, 5, width),
        cellIndex(7, 0, width),
      ]),
    )

    expect(readSpillValues(kernel, cellIndex(3, 0, width))).toEqual([
      { tag: ValueTag.Number, value: 10 },
      { tag: ValueTag.Number, value: 12 },
      { tag: ValueTag.Number, value: 14 },
      { tag: ValueTag.Number, value: 16 },
      { tag: ValueTag.Number, value: 18 },
      { tag: ValueTag.Number, value: 20 },
    ])
    expect(readSpillValues(kernel, cellIndex(3, 5, width))).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 4 },
      { tag: ValueTag.Number, value: 6 },
    ])
    expect(readSpillValues(kernel, cellIndex(5, 0, width))).toEqual([
      { tag: ValueTag.Number, value: 2 },
      { tag: ValueTag.Number, value: 5 },
    ])

    expect(kernel.readTags()[cellIndex(5, 5, width)]).toBe(ValueTag.Number)
    expect(kernel.readNumbers()[cellIndex(5, 5, width)]).toBe(21)

    expect(readSpillValues(kernel, cellIndex(7, 0, width))).toEqual([
      { tag: ValueTag.Number, value: 1 },
      { tag: ValueTag.Number, value: 3 },
      { tag: ValueTag.Number, value: 6 },
      { tag: ValueTag.Number, value: 10 },
      { tag: ValueTag.Number, value: 15 },
      { tag: ValueTag.Number, value: 21 },
    ])
  })

  it('keeps canonical GROUPBY and PIVOTBY spills stable on the wasm path', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(80, 4, 1, 8, 16)

    const inputStrings = ['Region', 'Product', 'Sales', 'East', 'West', 'Widget', 'Gizmo']
    const packedStrings = packStrings(inputStrings)
    kernel.uploadStringLengths(packedStrings.lengths)
    kernel.uploadStrings(packedStrings.offsets, packedStrings.lengths, packedStrings.data)

    const cellTags = new Uint8Array(50)
    const cellNumbers = new Float64Array(50)
    const cellStringIds = new Uint32Array(50)
    const cellErrors = new Uint16Array(50)
    const writeStringCell = (row: number, col: number, stringId: number) => {
      const index = cellIndex(row, col, width)
      cellTags[index] = ValueTag.String
      cellStringIds[index] = stringId
    }
    const writeNumberCell = (row: number, col: number, value: number) => {
      const index = cellIndex(row, col, width)
      cellTags[index] = ValueTag.Number
      cellNumbers[index] = value
    }

    writeStringCell(0, 0, 0)
    writeStringCell(0, 1, 1)
    writeStringCell(0, 2, 2)
    writeStringCell(1, 0, 3)
    writeStringCell(1, 1, 5)
    writeNumberCell(1, 2, 10)
    writeStringCell(2, 0, 4)
    writeStringCell(2, 1, 5)
    writeNumberCell(2, 2, 7)
    writeStringCell(3, 0, 3)
    writeStringCell(3, 1, 6)
    writeNumberCell(3, 2, 5)
    writeStringCell(4, 0, 4)
    writeStringCell(4, 1, 6)
    writeNumberCell(4, 2, 4)

    kernel.writeCells(cellTags, cellNumbers, cellStringIds, cellErrors)
    kernel.uploadRangeMembers(
      Uint32Array.from([
        cellIndex(0, 0, width),
        cellIndex(1, 0, width),
        cellIndex(2, 0, width),
        cellIndex(3, 0, width),
        cellIndex(4, 0, width),
        cellIndex(0, 1, width),
        cellIndex(1, 1, width),
        cellIndex(2, 1, width),
        cellIndex(3, 1, width),
        cellIndex(4, 1, width),
        cellIndex(0, 2, width),
        cellIndex(1, 2, width),
        cellIndex(2, 2, width),
        cellIndex(3, 2, width),
        cellIndex(4, 2, width),
      ]),
      Uint32Array.from([0, 5, 10]),
      Uint32Array.from([5, 5, 5]),
    )
    kernel.uploadRangeShapes(Uint32Array.from([5, 5, 5]), Uint32Array.from([1, 1, 1]))

    const programs = packPrograms([
      [encodePushRange(0), encodePushRange(2), encodeCall(BuiltinId.GroupbySumCanonical, 2), encodeRet()],
      [encodePushRange(0), encodePushRange(1), encodePushRange(2), encodeCall(BuiltinId.PivotbySumCanonical, 3), encodeRet()],
    ])
    kernel.uploadPrograms(
      programs.programs,
      programs.offsets,
      programs.lengths,
      Uint32Array.from([cellIndex(0, 4, width), cellIndex(0, 7, width)]),
    )
    kernel.uploadConstants(new Float64Array(), Uint32Array.from([0, 0]), Uint32Array.from([0, 0]))
    kernel.evalBatch(Uint32Array.from([cellIndex(0, 4, width), cellIndex(0, 7, width)]))

    expect(kernel.readSpillRows()[cellIndex(0, 4, width)]).toBe(4)
    expect(kernel.readSpillCols()[cellIndex(0, 4, width)]).toBe(2)
    expect(readSpillValues(kernel, cellIndex(0, 4, width), inputStrings)).toEqual([
      { tag: ValueTag.String, value: 'Region', stringId: 0 },
      { tag: ValueTag.String, value: 'Sales', stringId: 0 },
      { tag: ValueTag.String, value: 'East', stringId: 0 },
      { tag: ValueTag.Number, value: 15 },
      { tag: ValueTag.String, value: 'West', stringId: 0 },
      { tag: ValueTag.Number, value: 11 },
      { tag: ValueTag.String, value: 'Total', stringId: 0 },
      { tag: ValueTag.Number, value: 26 },
    ])

    expect(kernel.readSpillRows()[cellIndex(0, 7, width)]).toBe(4)
    expect(kernel.readSpillCols()[cellIndex(0, 7, width)]).toBe(4)
    expect(readSpillValues(kernel, cellIndex(0, 7, width), inputStrings)).toEqual([
      { tag: ValueTag.String, value: 'Region', stringId: 0 },
      { tag: ValueTag.String, value: 'Widget', stringId: 0 },
      { tag: ValueTag.String, value: 'Gizmo', stringId: 0 },
      { tag: ValueTag.String, value: 'Total', stringId: 0 },
      { tag: ValueTag.String, value: 'East', stringId: 0 },
      { tag: ValueTag.Number, value: 10 },
      { tag: ValueTag.Number, value: 5 },
      { tag: ValueTag.Number, value: 15 },
      { tag: ValueTag.String, value: 'West', stringId: 0 },
      { tag: ValueTag.Number, value: 7 },
      { tag: ValueTag.Number, value: 4 },
      { tag: ValueTag.Number, value: 11 },
      { tag: ValueTag.String, value: 'Total', stringId: 0 },
      { tag: ValueTag.Number, value: 17 },
      { tag: ValueTag.Number, value: 9 },
      { tag: ValueTag.Number, value: 26 },
    ])
  })
})
