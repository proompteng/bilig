import { describe, expect, it } from 'vitest'
import * as fc from 'fast-check'
import { runProperty } from '@bilig/test-fuzz'
import { createKernelSync } from '../index.js'

describe('wasm kernel bridge fuzz', () => {
  it('should preserve uploaded bridge buffers and monotonic capacity growth', async () => {
    await runProperty({
      suite: 'wasm-kernel/bridge/upload-roundtrip',
      arbitrary: kernelUploadSpecArbitrary,
      predicate: async (spec) => {
        const kernel = createKernelSync()
        kernel.init(spec.cellCapacity, spec.formulaCapacity, spec.constantCapacity, spec.rangeCapacity, spec.memberCapacity)

        kernel.uploadPrograms(
          new Uint32Array(spec.programs),
          new Uint32Array(spec.programOffsets),
          new Uint32Array(spec.programLengths),
          new Uint32Array(spec.programTargets),
        )
        kernel.uploadConstants(
          new Float64Array(spec.constants),
          new Uint32Array(spec.constantOffsets),
          new Uint32Array(spec.constantLengths),
        )
        kernel.uploadRangeMembers(new Uint32Array(spec.members), new Uint32Array(spec.rangeOffsets), new Uint32Array(spec.rangeLengths))
        kernel.uploadRangeShapes(new Uint32Array(spec.rangeRowCounts), new Uint32Array(spec.rangeColCounts))
        kernel.writeCells(
          new Uint8Array(spec.tags),
          new Float64Array(spec.numbers),
          new Uint32Array(spec.stringIds),
          new Uint16Array(spec.errors),
        )

        kernel.ensureCellCapacity(spec.cellCapacity + 2)
        kernel.ensureFormulaCapacity(spec.formulaCapacity + 2)
        kernel.ensureConstantCapacity(spec.constantCapacity + 2)
        kernel.ensureRangeCapacity(spec.rangeCapacity + 2)
        kernel.ensureMemberCapacity(spec.memberCapacity + 2)

        expect(kernel.getCellCapacity()).toBeGreaterThanOrEqual(spec.cellCapacity + 2)
        expect(kernel.getFormulaCapacity()).toBeGreaterThanOrEqual(spec.formulaCapacity + 2)
        expect(kernel.getConstantCapacity()).toBeGreaterThanOrEqual(spec.constantCapacity + 2)
        expect(kernel.getRangeCapacity()).toBeGreaterThanOrEqual(spec.rangeCapacity + 2)
        expect(kernel.getMemberCapacity()).toBeGreaterThanOrEqual(spec.memberCapacity + 2)

        expect(Array.from(kernel.readProgramOffsets().slice(0, spec.programOffsets.length))).toEqual(spec.programOffsets)
        expect(Array.from(kernel.readProgramLengths().slice(0, spec.programLengths.length))).toEqual(spec.programLengths)
        expect(Array.from(kernel.readConstantOffsets().slice(0, spec.constantOffsets.length))).toEqual(spec.constantOffsets)
        expect(Array.from(kernel.readConstantLengths().slice(0, spec.constantLengths.length))).toEqual(spec.constantLengths)
        expect(Array.from(kernel.readConstants().slice(0, spec.constants.length))).toEqual(spec.constants)
        expect(Array.from(kernel.readRangeOffsets().slice(0, spec.rangeOffsets.length))).toEqual(spec.rangeOffsets)
        expect(Array.from(kernel.readRangeLengths().slice(0, spec.rangeLengths.length))).toEqual(spec.rangeLengths)
        expect(Array.from(kernel.readRangeMembers().slice(0, spec.members.length))).toEqual(spec.members)
        expect(Array.from(kernel.readTags().slice(0, spec.tags.length))).toEqual(spec.tags)
        expect(Array.from(kernel.readNumbers().slice(0, spec.numbers.length))).toEqual(spec.numbers)
        expect(Array.from(kernel.readStringIds().slice(0, spec.stringIds.length))).toEqual(spec.stringIds)
        expect(Array.from(kernel.readErrors().slice(0, spec.errors.length))).toEqual(spec.errors)
      },
    })
  })
})

// Helpers

const kernelUploadSpecArbitrary = fc
  .record({
    cellCount: fc.integer({ min: 1, max: 12 }),
    formulaCount: fc.integer({ min: 1, max: 6 }),
    constantCount: fc.integer({ min: 1, max: 12 }),
    rangeCount: fc.integer({ min: 1, max: 6 }),
    memberCount: fc.integer({ min: 1, max: 16 }),
  })
  .chain(({ cellCount, formulaCount, constantCount, rangeCount, memberCount }) =>
    fc
      .record({
        cellCapacity: fc.constant(cellCount + 2),
        formulaCapacity: fc.constant(formulaCount + 2),
        constantCapacity: fc.constant(constantCount + 2),
        rangeCapacity: fc.constant(rangeCount + 2),
        memberCapacity: fc.constant(memberCount + 2),
        tags: fc.array(fc.integer({ min: 0, max: 4 }), { minLength: cellCount, maxLength: cellCount }),
        numbers: fc.array(fc.double({ noNaN: true, noDefaultInfinity: true }), { minLength: cellCount, maxLength: cellCount }),
        stringIds: fc.array(fc.integer({ min: 0, max: 20 }), { minLength: cellCount, maxLength: cellCount }),
        errors: fc.array(fc.integer({ min: 0, max: 20 }), { minLength: cellCount, maxLength: cellCount }),
        programLengths: fc.array(fc.integer({ min: 0, max: 3 }), { minLength: formulaCount, maxLength: formulaCount }),
        constantLengths: fc.array(fc.integer({ min: 0, max: 3 }), { minLength: formulaCount, maxLength: formulaCount }),
        rangeLengths: fc.array(fc.integer({ min: 0, max: 4 }), { minLength: rangeCount, maxLength: rangeCount }),
        rangeRowCounts: fc.array(fc.integer({ min: 1, max: 4 }), { minLength: rangeCount, maxLength: rangeCount }),
        rangeColCounts: fc.array(fc.integer({ min: 1, max: 4 }), { minLength: rangeCount, maxLength: rangeCount }),
      })
      .map((generated) => {
        const programOffsets = offsetsFor(generated.programLengths)
        const constantOffsets = offsetsFor(generated.constantLengths)
        const rangeOffsets = offsetsFor(generated.rangeLengths)
        return {
          cellCapacity: generated.cellCapacity,
          formulaCapacity: generated.formulaCapacity,
          constantCapacity: generated.constantCapacity,
          rangeCapacity: generated.rangeCapacity,
          memberCapacity: generated.memberCapacity,
          tags: generated.tags,
          numbers: generated.numbers,
          stringIds: generated.stringIds,
          errors: generated.errors,
          programLengths: generated.programLengths,
          constantLengths: generated.constantLengths,
          rangeLengths: generated.rangeLengths,
          rangeRowCounts: generated.rangeRowCounts,
          rangeColCounts: generated.rangeColCounts,
          programs: Array.from({ length: sum(generated.programLengths) }, (_, index) => index),
          programOffsets,
          programTargets: Array.from({ length: formulaCount }, (_, index) => index),
          constants: Array.from({ length: sum(generated.constantLengths) }, (_, index) => index + 0.5),
          constantOffsets,
          members: Array.from({ length: sum(generated.rangeLengths) }, (_, index) => index),
          rangeOffsets,
        }
      }),
  )

function offsetsFor(lengths: readonly number[]): number[] {
  const offsets: number[] = []
  let cursor = 0
  lengths.forEach((length) => {
    offsets.push(cursor)
    cursor += length
  })
  return offsets
}

function sum(values: readonly number[]): number {
  return values.reduce((total, value) => total + value, 0)
}
