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

function encodePushString(stringId: number): number {
  return (Opcode.PushString << 24) | stringId
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

function packStrings(values: readonly string[]): {
  offsets: Uint32Array
  lengths: Uint32Array
  data: Uint16Array
} {
  const offsets: number[] = []
  const lengths: number[] = []
  const data: number[] = []
  let cursor = 0

  for (const value of values) {
    const codes = Array.from(value, (char) => char.charCodeAt(0))
    offsets.push(cursor)
    lengths.push(codes.length)
    data.push(...codes)
    cursor += codes.length
  }

  return {
    offsets: Uint32Array.from(offsets),
    lengths: Uint32Array.from(lengths),
    data: Uint16Array.from(data),
  }
}

function cellIndex(row: number, col: number, width: number): number {
  return row * width + col
}

function readStringCell(kernel: KernelInstance, index: number, pooledStrings: readonly string[]): string {
  expect(kernel.readTags()[index]).toBe(ValueTag.String)
  const raw = kernel.readStringIds()[index] ?? 0
  const outputIndex = raw >= OUTPUT_STRING_BASE ? raw - OUTPUT_STRING_BASE : -1
  return outputIndex >= 0 ? (kernel.readOutputStrings()[outputIndex] ?? '') : (pooledStrings[raw] ?? '')
}

function readSpillValues(kernel: KernelInstance, ownerCellIndex: number, pooledStrings: readonly string[]): CellValue[] {
  const offset = kernel.readSpillOffsets()[ownerCellIndex] ?? 0
  const length = kernel.readSpillLengths()[ownerCellIndex] ?? 0
  const tags = kernel.readSpillTags()
  const values = kernel.readSpillNumbers()
  const outputStrings = kernel.readOutputStrings()
  return Array.from({ length }, (_, index) => {
    const tag = tags[offset + index] ?? ValueTag.Empty
    const rawValue = values[offset + index] ?? 0
    if (tag === ValueTag.String) {
      const outputIndex = rawValue >= OUTPUT_STRING_BASE ? rawValue - OUTPUT_STRING_BASE : -1
      return {
        tag,
        value: outputIndex >= 0 ? (outputStrings[outputIndex] ?? '') : (pooledStrings[rawValue] ?? ''),
        stringId: 0,
      }
    }
    if (tag === ValueTag.Number) {
      return { tag, value: rawValue }
    }
    return { tag: ValueTag.Empty }
  })
}

describe('wasm kernel text-formatting dispatch', () => {
  it('keeps text formatting, conversion, and split builtins stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 24
    const pooledStrings = [
      'alpha-beta-gamma',
      '-',
      '  -12.5e1 ',
      'ABC',
      'é',
      'A\u0001B',
      'ＡＢＣ　１２３',
      'ABC 123',
      'ｶﾀｶﾅ',
      '2,500.27%',
      '.',
      ',',
      'alpha',
      '#,##0.00',
      'prefix @',
      'ABCD',
    ]
    const packedStrings = packStrings(pooledStrings)
    kernel.init(64, pooledStrings.length, 8, 1, 1)
    kernel.uploadStrings(packedStrings.offsets, packedStrings.lengths, packedStrings.data)
    kernel.writeCells(new Uint8Array(64), new Float64Array(64), new Uint32Array(64), new Uint16Array(64))

    const packed = packPrograms([
      [encodePushString(0), encodePushString(1), encodeCall(BuiltinId.Textbefore, 2), encodeRet()],
      [encodePushString(0), encodePushString(1), encodePushNumber(0), encodeCall(BuiltinId.Textafter, 3), encodeRet()],
      [encodePushString(0), encodePushString(1), encodeCall(BuiltinId.Textsplit, 2), encodeRet()],
      [encodePushString(2), encodeCall(BuiltinId.Value, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Char, 1), encodeRet()],
      [encodePushString(3), encodeCall(BuiltinId.Code, 1), encodeRet()],
      [encodePushString(4), encodeCall(BuiltinId.Unicode, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Unichar, 1), encodeRet()],
      [encodePushString(5), encodeCall(BuiltinId.Clean, 1), encodeRet()],
      [encodePushString(6), encodeCall(BuiltinId.Asc, 1), encodeRet()],
      [encodePushString(7), encodeCall(BuiltinId.Jis, 1), encodeRet()],
      [encodePushString(8), encodeCall(BuiltinId.Dbcs, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Bahttext, 1), encodeRet()],
      [encodePushString(9), encodePushString(10), encodePushString(11), encodeCall(BuiltinId.Numbervalue, 3), encodeRet()],
      [encodePushString(12), encodePushNumber(0), encodeCall(BuiltinId.Valuetotext, 2), encodeRet()],
      [encodePushNumber(0), encodePushString(13), encodeCall(BuiltinId.Text, 2), encodeRet()],
      [encodePushString(12), encodePushString(14), encodeCall(BuiltinId.Text, 2), encodeRet()],
      [encodePushString(15), encodeCall(BuiltinId.Phonetic, 1), encodeRet()],
    ])
    const constants = packConstants([[], [-1], [], [], [65], [], [], [66], [], [], [], [], [21.25], [], [1], [1234.567], [], []])
    const targets = Uint32Array.from(Array.from({ length: 18 }, (_, index) => cellIndex(1, index, width)))
    kernel.uploadPrograms(packed.programs, packed.offsets, packed.lengths, targets)
    kernel.uploadConstants(constants.constants, constants.offsets, constants.lengths)
    kernel.evalBatch(targets)

    expect(readStringCell(kernel, cellIndex(1, 0, width), pooledStrings)).toBe('alpha')
    expect(readStringCell(kernel, cellIndex(1, 1, width), pooledStrings)).toBe('gamma')
    expect(kernel.readSpillRows()[cellIndex(1, 2, width)]).toBe(1)
    expect(kernel.readSpillCols()[cellIndex(1, 2, width)]).toBe(3)
    expect(readSpillValues(kernel, cellIndex(1, 2, width), pooledStrings)).toEqual([
      { tag: ValueTag.String, value: 'alpha', stringId: 0 },
      { tag: ValueTag.String, value: 'beta', stringId: 0 },
      { tag: ValueTag.String, value: 'gamma', stringId: 0 },
    ])
    expect(kernel.readNumbers()[cellIndex(1, 3, width)]).toBeCloseTo(-125, 12)
    expect(readStringCell(kernel, cellIndex(1, 4, width), pooledStrings)).toBe('A')
    expect(kernel.readNumbers()[cellIndex(1, 5, width)]).toBe(65)
    expect(kernel.readNumbers()[cellIndex(1, 6, width)]).toBe(233)
    expect(readStringCell(kernel, cellIndex(1, 7, width), pooledStrings)).toBe('B')
    expect(readStringCell(kernel, cellIndex(1, 8, width), pooledStrings)).toBe('AB')
    expect(readStringCell(kernel, cellIndex(1, 9, width), pooledStrings)).toBe('ABC 123')
    expect(readStringCell(kernel, cellIndex(1, 10, width), pooledStrings)).toBe('ＡＢＣ　１２３')
    expect(readStringCell(kernel, cellIndex(1, 11, width), pooledStrings)).toBe('カタカナ')
    expect(readStringCell(kernel, cellIndex(1, 12, width), pooledStrings)).toBe('ยี่สิบเอ็ดบาทยี่สิบห้าสตางค์')
    expect(kernel.readNumbers()[cellIndex(1, 13, width)]).toBeCloseTo(25.0027, 12)
    expect(readStringCell(kernel, cellIndex(1, 14, width), pooledStrings)).toBe('"alpha"')
    expect(readStringCell(kernel, cellIndex(1, 15, width), pooledStrings)).toBe('1,234.57')
    expect(readStringCell(kernel, cellIndex(1, 16, width), pooledStrings)).toBe('prefix alpha')
    expect(readStringCell(kernel, cellIndex(1, 17, width), pooledStrings)).toBe('ABCD')
  })
})
