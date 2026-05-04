import { describe, expect, it } from 'vitest'
import { stringifyJson, textToolResult } from './workbook-agent-tool-shared.js'

describe('workbook agent tool shared helpers', () => {
  it('formats text tool results for Codex input text content', () => {
    expect(textToolResult('Completed workbook read')).toEqual({
      success: true,
      contentItems: [{ type: 'inputText', text: 'Completed workbook read' }],
    })
    expect(textToolResult('Could not resolve selector', false)).toEqual({
      success: false,
      contentItems: [{ type: 'inputText', text: 'Could not resolve selector' }],
    })
  })

  it('pretty prints JSON payloads', () => {
    expect(
      stringifyJson({
        ok: true,
        rows: [1, 2],
      }),
    ).toBe(`{
  "ok": true,
  "rows": [
    1,
    2
  ]
}`)
  })
})
