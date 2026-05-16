import { describe, expect, it } from 'vitest'
import {
  applyBatchArgsSchema,
  clearRangeStyleArgsSchema,
  mergeCellsArgsSchema,
  renderCommitArgsSchema,
  setRangeNumberFormatArgsSchema,
  setRangeStyleArgsSchema,
  unmergeCellsArgsSchema,
  updatePresenceArgsSchema,
} from '../mutators.js'

const range = {
  sheetName: 'Sheet1',
  startAddress: 'A1',
  endAddress: 'B2',
}

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

  it('accepts merge and unmerge cell mutation payloads', () => {
    const payload = {
      documentId: 'doc-1',
      range: { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'B1' },
    }

    expect(mergeCellsArgsSchema.safeParse(payload).success).toBe(true)
    expect(unmergeCellsArgsSchema.safeParse(payload).success).toBe(true)
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

  it('accepts protocol-owned style enum values', () => {
    expect(
      setRangeStyleArgsSchema.parse({
        documentId: 'doc-1',
        range,
        patch: {
          alignment: {
            horizontal: 'centerContinuous',
            vertical: 'distributed',
          },
          borders: {
            top: {
              style: 'double',
              weight: 'thick',
              color: '#111111',
            },
          },
        },
      }),
    ).toMatchObject({
      patch: {
        alignment: {
          horizontal: 'centerContinuous',
          vertical: 'distributed',
        },
        borders: {
          top: {
            style: 'double',
            weight: 'thick',
          },
        },
      },
    })
  })

  it('rejects style values outside the protocol contract', () => {
    expect(() =>
      setRangeStyleArgsSchema.parse({
        documentId: 'doc-1',
        range,
        patch: {
          borders: {
            top: {
              style: 'hairline',
            },
          },
        },
      }),
    ).toThrow()
  })

  it('accepts protocol-owned clear style field names', () => {
    expect(
      clearRangeStyleArgsSchema.parse({
        documentId: 'doc-1',
        range,
        fields: ['fontBold', 'alignmentTextRotation', 'borderLeft'],
      }),
    ).toMatchObject({
      fields: ['fontBold', 'alignmentTextRotation', 'borderLeft'],
    })
  })

  it('accepts protocol-owned number format enum values', () => {
    expect(
      setRangeNumberFormatArgsSchema.parse({
        documentId: 'doc-1',
        range,
        format: {
          kind: 'accounting',
          currency: 'USD',
          decimals: 2,
          negativeStyle: 'parentheses',
          zeroStyle: 'dash',
          dateStyle: 'iso',
        },
      }),
    ).toMatchObject({
      format: {
        kind: 'accounting',
        negativeStyle: 'parentheses',
        zeroStyle: 'dash',
        dateStyle: 'iso',
      },
    })
  })
})
