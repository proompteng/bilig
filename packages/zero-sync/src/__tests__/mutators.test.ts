import { describe, expect, it } from 'vitest'
import { applyBatchArgsSchema, renderCommitArgsSchema, updatePresenceArgsSchema } from '../mutators.js'

describe('zero sync mutator schemas', () => {
  it('accepts workbook presence updates with the current selection payload', () => {
    const result = updatePresenceArgsSchema.safeParse({
      documentId: 'doc-1',
      sessionId: 'session-1',
      presenceClientId: 'presence:self',
      sheetName: 'Sheet1',
      address: 'B2',
      selection: {
        sheetName: 'Sheet1',
        address: 'B2',
      },
    })

    expect(result.success).toBe(true)
  })

  it('rejects workbook presence updates with malformed selection payloads', () => {
    const result = updatePresenceArgsSchema.safeParse({
      documentId: 'doc-1',
      sessionId: 'session-1',
      selection: {
        sheetName: 'Sheet1',
        address: 42,
      },
    })

    expect(result.success).toBe(false)
  })

  it('accepts engine batches with valid workbook ops', () => {
    const result = applyBatchArgsSchema.safeParse({
      documentId: 'doc-1',
      batch: {
        id: 'batch-1',
        replicaId: 'replica-1',
        clock: { counter: 1 },
        ops: [{ kind: 'upsertWorkbook', name: 'Book' }],
      },
    })

    expect(result.success).toBe(true)
  })

  it('rejects engine batches with malformed workbook ops', () => {
    const result = applyBatchArgsSchema.safeParse({
      documentId: 'doc-1',
      batch: {
        id: 'batch-1',
        replicaId: 'replica-1',
        clock: { counter: 1 },
        ops: [{ kind: 'setCellValue', sheetName: 'Sheet1', address: 'A1' }],
      },
    })

    expect(result.success).toBe(false)
  })

  it('rejects render commits with malformed commit ops', () => {
    const result = renderCommitArgsSchema.safeParse({
      documentId: 'doc-1',
      ops: [{ kind: 'deleteCell', sheetName: 'Sheet1' }],
    })

    expect(result.success).toBe(false)
  })
})
