import { describe, expect, test } from 'vitest'
import {
  parseClipboardContent,
  parseClipboardHtml,
  parseClipboardPlainText,
  serializeClipboardMatrix,
  serializeClipboardPlainText,
} from '../gridClipboard.js'

describe('gridClipboard', () => {
  test('serializes matrix signatures and plain text', () => {
    const values = [
      ['A', 'B'],
      ['1', '2'],
    ] as const

    expect(serializeClipboardMatrix(values)).toBe('A\u001fB\u001e1\u001f2')
    expect(serializeClipboardPlainText(values)).toBe('A\tB\n1\t2')
  })

  test('parses clipboard plain text with mixed line endings', () => {
    expect(parseClipboardPlainText('A\tB\r\n1\t2\r3\t4')).toEqual([
      ['A', 'B'],
      ['1', '2'],
      ['3', '4'],
    ])
    expect(parseClipboardPlainText('')).toEqual([])
  })

  test('parses quoted csv with commas and multiline cells', () => {
    expect(parseClipboardPlainText('"Name","Notes"\n"Alice","Line 1\nLine 2"')).toEqual([
      ['Name', 'Notes'],
      ['Alice', 'Line 1\nLine 2'],
    ])
  })

  test('parses quoted tabular plain text from spreadsheets', () => {
    expect(parseClipboardPlainText('"A\tB"\tC\n1\t"two\tcolumns"')).toEqual([
      ['A\tB', 'C'],
      ['1', 'two\tcolumns'],
    ])
  })

  test('parses html table clipboard content', () => {
    expect(parseClipboardHtml('<table><tr><td>A</td><td><b>B</b></td></tr><tr><td>1</td><td>two<br>lines</td></tr></table>')).toEqual([
      ['A', 'B'],
      ['1', 'two\nlines'],
    ])
  })

  test('prefers html clipboard data when available', () => {
    expect(parseClipboardContent('A,B\n1,2', '<table><tr><td>A</td><td>B</td></tr><tr><td>1</td><td>2</td></tr></table>')).toEqual([
      ['A', 'B'],
      ['1', '2'],
    ])
  })
})
