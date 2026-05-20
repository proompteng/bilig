import { describe, expect, test } from 'vitest'
import { buildInternalClipboardRange, matchesInternalClipboardPaste } from '../gridInternalClipboard.js'

describe('gridInternalClipboard', () => {
  test('builds a clipboard signature and address range from copied values', () => {
    expect(
      buildInternalClipboardRange({ x: 1, y: 2, width: 2, height: 2 }, [
        ['1', '2'],
        ['3', '4'],
      ]),
    ).toEqual({
      operation: 'copy',
      sourceStartAddress: 'B3',
      sourceEndAddress: 'C4',
      signature: '1\u001f2\u001e3\u001f4',
      plainText: '1\t2\n3\t4',
      valuesOnlyPlainText: '1\t2\n3\t4',
      rowCount: 2,
      colCount: 2,
    })
  })

  test('keeps cut intent with the captured internal range', () => {
    expect(buildInternalClipboardRange({ x: 1, y: 2, width: 1, height: 1 }, [['move-me']], 'cut')).toMatchObject({
      operation: 'cut',
      sourceStartAddress: 'B3',
      sourceEndAddress: 'B3',
      plainText: 'move-me',
      valuesOnlyPlainText: 'move-me',
    })
  })

  test('keeps a separate values-only clipboard payload for resolved formula results', () => {
    expect(buildInternalClipboardRange({ x: 1, y: 2, width: 2, height: 1 }, [['3', '=B3*2']], 'copy', [['3', '6']])).toMatchObject({
      plainText: '3\t=B3*2',
      valuesOnlyPlainText: '3\t6',
    })
  })

  test('matches internal clipboard pastes by signature and rectangular shape', () => {
    const clipboard = buildInternalClipboardRange({ x: 1, y: 2, width: 2, height: 2 }, [
      ['1', '2'],
      ['3', '4'],
    ])

    expect(
      matchesInternalClipboardPaste(clipboard, [
        ['1', '2'],
        ['3', '4'],
      ]),
    ).toBe(true)

    expect(
      matchesInternalClipboardPaste(clipboard, [
        ['1', '2', '3'],
        ['4', '5', '6'],
      ]),
    ).toBe(false)
  })

  test('matches system clipboard text with trimmed trailing blank rows from an internal range', () => {
    const clipboard = buildInternalClipboardRange({ x: 1, y: 2, width: 1, height: 5 }, [[''], ['kept'], [''], [''], ['']])

    expect(matchesInternalClipboardPaste(clipboard, [[''], ['kept']])).toBe(true)
    expect(matchesInternalClipboardPaste(clipboard, [[''], ['changed']])).toBe(false)
  })
})
