import { describe, expect, it } from 'vitest'
import {
  isWorkbookAgentUiContext,
  normalizeTimelineEntry,
  normalizeZeroWorkbookChatThread,
  parseNumericValue,
} from './workbook-chat-thread-normalizers.js'

describe('workbook-chat-thread-normalizers', () => {
  it('parses only complete safe integer database values', () => {
    expect(parseNumericValue(42)).toBe(42)
    expect(parseNumericValue('42')).toBe(42)
    expect(parseNumericValue(' 42 ')).toBe(42)

    expect(parseNumericValue(-42)).toBeNull()
    expect(parseNumericValue('-42')).toBeNull()
    expect(parseNumericValue(42.5)).toBeNull()
    expect(parseNumericValue('42ms')).toBeNull()
    expect(parseNumericValue('42.5')).toBeNull()
    expect(parseNumericValue('')).toBeNull()
    expect(parseNumericValue('9007199254740992')).toBeNull()
  })

  it('rejects summaries with malformed numeric fields', () => {
    expect(
      normalizeZeroWorkbookChatThread({
        workbookId: 'doc-1',
        threadId: 'thr-1',
        scope: 'private',
        ownerUserId: 'alex@example.com',
        executionPolicy: 'autoApplyAll',
        context: null,
        updatedAtUnixMs: 200,
        entryCount: 3,
        reviewQueueItemCount: -1,
        latestEntryText: 'Done',
      }),
    ).toBeNull()
    expect(
      normalizeZeroWorkbookChatThread({
        workbookId: 'doc-1',
        threadId: 'thr-1',
        scope: 'private',
        ownerUserId: 'alex@example.com',
        executionPolicy: 'autoApplyAll',
        context: null,
        updatedAtUnixMs: 200,
        entryCount: -1,
        reviewQueueItemCount: 0,
        latestEntryText: 'Done',
      }),
    ).toBeNull()
  })

  it('rejects unsafe workbook UI viewport context', () => {
    expect(
      isWorkbookAgentUiContext({
        selection: { sheetName: 'Sheet1', address: 'A1' },
        viewport: { rowStart: 0, rowEnd: 10, colStart: 0, colEnd: 5 },
      }),
    ).toBe(true)
    expect(
      isWorkbookAgentUiContext({
        selection: { sheetName: 'Sheet1', address: 'A1' },
        viewport: { rowStart: 10, rowEnd: 0, colStart: 0, colEnd: 5 },
      }),
    ).toBe(false)
    expect(
      isWorkbookAgentUiContext({
        selection: { sheetName: 'Sheet1', address: 'A1' },
        viewport: { rowStart: 0, rowEnd: 10.5, colStart: 0, colEnd: 5 },
      }),
    ).toBe(false)
  })

  it('rejects unsafe revision citations', () => {
    expect(
      normalizeTimelineEntry({
        entryId: 'entry-1',
        kind: 'assistant',
        citationsJson: [{ kind: 'revision', revision: 7 }],
      }),
    ).not.toBeNull()
    expect(
      normalizeTimelineEntry({
        entryId: 'entry-1',
        kind: 'assistant',
        citationsJson: [{ kind: 'revision', revision: 7.5 }],
      }),
    ).toBeNull()
    expect(
      normalizeTimelineEntry({
        entryId: 'entry-1',
        kind: 'assistant',
        citationsJson: [{ kind: 'revision', revision: Number.MAX_SAFE_INTEGER + 1 }],
      }),
    ).toBeNull()
  })
})
