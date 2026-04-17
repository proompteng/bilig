import { describe, expect, it } from 'vitest'
import fc from 'fast-check'
import type { EngineOp, EngineOpBatch } from '../index.js'
import { isEngineOp, isEngineOpBatch } from '../index.js'
import { runProperty } from '@bilig/test-fuzz'

type GuardedOp = Extract<
  EngineOp,
  { kind: 'upsertWorkbook' } | { kind: 'upsertSheet' } | { kind: 'setCellValue' } | { kind: 'setCellFormula' } | { kind: 'clearCell' }
>

describe('workbook domain guard fuzz', () => {
  it('should accept generated valid engine ops and batches from the supported subset', async () => {
    await runProperty({
      suite: 'workbook-domain/guards/valid-subset',
      arbitrary: fc.record({
        op: engineOpArbitrary,
        batch: engineOpBatchArbitrary,
      }),
      predicate: async ({ op, batch }) => {
        expect(isEngineOp(op)).toBe(true)
        expect(isEngineOpBatch(batch)).toBe(true)
      },
    })
  })

  it('should reject corrupted ops and batches from the same subset', async () => {
    await runProperty({
      suite: 'workbook-domain/guards/reject-corruption',
      arbitrary: fc.record({
        op: engineOpArbitrary,
        batch: engineOpBatchArbitrary,
      }),
      predicate: async ({ op, batch }) => {
        expect(isEngineOp(corruptOp(op))).toBe(false)
        expect(isEngineOpBatch(corruptBatch(batch))).toBe(false)
      },
    })
  })
})

// Helpers

const engineOpArbitrary = fc.oneof<GuardedOp>(
  fc.constantFrom('Book', 'Spec', 'Revenue').map((name) => ({ kind: 'upsertWorkbook', name })),
  fc
    .record({
      name: fc.constantFrom('Sheet1', 'Sheet2'),
      order: fc.integer({ min: 0, max: 4 }),
    })
    .map(({ name, order }) => ({ kind: 'upsertSheet', name, order })),
  fc
    .record({
      sheetName: fc.constantFrom('Sheet1', 'Sheet2'),
      address: fc.constantFrom('A1', 'B2', 'C3', 'D4'),
      value: fc.oneof(fc.integer({ min: -50, max: 50 }), fc.boolean(), fc.constantFrom('north', 'south'), fc.constant(null)),
    })
    .map(({ sheetName, address, value }) => ({ kind: 'setCellValue', sheetName, address, value })),
  fc
    .record({
      sheetName: fc.constantFrom('Sheet1', 'Sheet2'),
      address: fc.constantFrom('A1', 'B2', 'C3', 'D4'),
      formula: fc.constantFrom('A1+1', 'B2*2', '1+2'),
    })
    .map(({ sheetName, address, formula }) => ({ kind: 'setCellFormula', sheetName, address, formula })),
  fc
    .record({
      sheetName: fc.constantFrom('Sheet1', 'Sheet2'),
      address: fc.constantFrom('A1', 'B2', 'C3', 'D4'),
    })
    .map(({ sheetName, address }) => ({ kind: 'clearCell', sheetName, address })),
)

const engineOpBatchArbitrary = fc
  .record({
    id: fc.uuid(),
    replicaId: fc.constantFrom('replica-a', 'replica-b'),
    counter: fc.integer({ min: 1, max: 1_000 }),
    ops: fc.array(engineOpArbitrary, { minLength: 1, maxLength: 8 }),
  })
  .map(
    ({ id, replicaId, counter, ops }) =>
      ({
        id,
        replicaId,
        clock: { counter },
        ops,
      }) satisfies EngineOpBatch,
  )

function corruptOp(op: GuardedOp): unknown {
  switch (op.kind) {
    case 'upsertWorkbook':
      return { ...op, name: 7 }
    case 'upsertSheet':
      return { ...op, order: 'bad' }
    case 'setCellValue':
      return { ...op, address: null }
    case 'setCellFormula':
      return { ...op, formula: 9 }
    case 'clearCell':
      return { ...op, sheetName: null }
  }
}

function corruptBatch(batch: EngineOpBatch): unknown {
  return {
    ...batch,
    id: 7,
  }
}
