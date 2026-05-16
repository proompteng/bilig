import { describe, expect, it } from 'vitest'
import { normalizeThreadSummary, parseNumericValue } from './workbook-chat-thread-normalizers.js'

describe('workbook-chat-thread-normalizers', () => {
  it('parses only complete safe integer database values', () => {
    expect(parseNumericValue(42)).toBe(42)
    expect(parseNumericValue('42')).toBe(42)
    expect(parseNumericValue(' 42 ')).toBe(42)
    expect(parseNumericValue('-42')).toBe(-42)

    expect(parseNumericValue(42.5)).toBeNull()
    expect(parseNumericValue('42ms')).toBeNull()
    expect(parseNumericValue('42.5')).toBeNull()
    expect(parseNumericValue('')).toBeNull()
    expect(parseNumericValue('9007199254740992')).toBeNull()
  })

  it('rejects summaries with malformed numeric fields', () => {
    expect(
      normalizeThreadSummary({
        threadId: 'thr-1',
        scope: 'private',
        ownerUserId: 'alex@example.com',
        updatedAtUnixMs: '200x',
        entryCount: '3',
        reviewQueueItemCount: '0',
        latestEntryText: 'Done',
      }),
    ).toBeNull()
  })
})
