import { describe, expect, it } from 'vitest'
import { BuiltinId, Opcode, ValueTag } from '@bilig/protocol'
import { createKernel } from '../index.js'

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

function cellIndex(row: number, col: number, width: number): number {
  return row * width + col
}

function readStringCell(kernel: Awaited<ReturnType<typeof createKernel>>, index: number, pooledStrings: readonly string[]): string {
  expect(kernel.readTags()[index]).toBe(ValueTag.String)
  const raw = kernel.readStringIds()[index] ?? 0
  const outputIndex = raw >= OUTPUT_STRING_BASE ? raw - OUTPUT_STRING_BASE : -1
  return outputIndex >= 0 ? (kernel.readOutputStrings()[outputIndex] ?? '') : (pooledStrings[raw] ?? '')
}

describe('wasm kernel text-special helpers', () => {
  it('keeps numeric and time text coercion stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 8
    kernel.init(24, 5, 4, 1, 1)
    kernel.uploadStrings(
      Uint32Array.from([0, 10, 17, 25]),
      Uint32Array.from([10, 7, 8, 3]),
      Uint16Array.from(Array.from('  -12.5e1 2:30 PM24:00:00bad', (char) => char.charCodeAt(0))),
    )
    kernel.writeCells(new Uint8Array(24), new Float64Array(24), new Uint32Array(24), new Uint16Array(24))

    const packed = packPrograms([
      [encodePushString(0), encodeCall(BuiltinId.Value, 1), encodeRet()],
      [encodePushString(1), encodeCall(BuiltinId.Timevalue, 1), encodeRet()],
      [encodePushString(2), encodeCall(BuiltinId.Timevalue, 1), encodeRet()],
      [encodePushString(3), encodeCall(BuiltinId.Value, 1), encodeRet()],
    ])
    kernel.uploadPrograms(
      packed.programs,
      packed.offsets,
      packed.lengths,
      Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]),
    )
    kernel.uploadConstants(new Float64Array(0), Uint32Array.from([0, 0, 0, 0]), Uint32Array.from([0, 0, 0, 0]))
    kernel.evalBatch(Uint32Array.from([cellIndex(1, 0, width), cellIndex(1, 1, width), cellIndex(1, 2, width), cellIndex(1, 3, width)]))

    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBeCloseTo(-125, 12)
    expect(kernel.readNumbers()[cellIndex(1, 1, width)]).toBeCloseTo(0.604166666667, 12)
    expect(kernel.readNumbers()[cellIndex(1, 2, width)]).toBe(0)
    expect(kernel.readTags()[cellIndex(1, 3, width)]).toBe(ValueTag.Error)
  })

  it('keeps unicode and width-conversion text helpers stable across refactors', async () => {
    const kernel = await createKernel()
    const width = 10
    kernel.init(30, 6, 2, 1, 1)
    const pooledStrings = ['é', 'A\u0007B', 'ABC 123', 'ｶﾀｶﾅ', 'カタカナ']
    kernel.uploadStrings(
      Uint32Array.from([0, 1, 4, 11, 17]),
      Uint32Array.from([1, 3, 7, 4, 4]),
      Uint16Array.from(Array.from('éA\u0007BABC 123ｶﾀｶﾅカタカナ', (char) => char.charCodeAt(0))),
    )
    kernel.writeCells(new Uint8Array(30), new Float64Array(30), new Uint32Array(30), new Uint16Array(30))

    const packed = packPrograms([
      [encodePushString(0), encodeCall(BuiltinId.Unicode, 1), encodeRet()],
      [encodePushNumber(0), encodeCall(BuiltinId.Unichar, 1), encodeRet()],
      [encodePushString(1), encodeCall(BuiltinId.Clean, 1), encodeRet()],
      [encodePushString(2), encodeCall(BuiltinId.Jis, 1), encodeRet()],
      [encodePushString(3), encodeCall(BuiltinId.Dbcs, 1), encodeRet()],
      [encodePushString(2), encodeCall(BuiltinId.Asc, 1), encodeRet()],
    ])
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
    const constants = packConstants([[], [12459], [], [], [], []])
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

    expect(kernel.readNumbers()[cellIndex(1, 0, width)]).toBe(233)
    expect(readStringCell(kernel, cellIndex(1, 1, width), pooledStrings)).toBe('カ')
    expect(readStringCell(kernel, cellIndex(1, 2, width), pooledStrings)).toBe('AB')
    expect(readStringCell(kernel, cellIndex(1, 3, width), pooledStrings)).toBe('ＡＢＣ　１２３')
    expect(readStringCell(kernel, cellIndex(1, 4, width), pooledStrings)).toBe('カタカナ')
    expect(readStringCell(kernel, cellIndex(1, 5, width), pooledStrings)).toBe('ABC 123')
  })
})
