import { describe, expect, test } from 'vitest'
import { getEditorPresentation, getOverlayStyle } from '../gridPresentation.js'

describe('gridPresentation', () => {
  test('uses the cell fill and text styling for the in-cell editor', () => {
    expect(
      getEditorPresentation({
        renderCell: {
          kind: 'string',
          displayText: 'hello',
          copyText: 'hello',
          align: 'left',
          wrap: false,
          color: '#f9fafb',
          font: 'italic 700 14px "JetBrainsMono Nerd Font","JetBrains Mono",monospace',
          fontSize: 14,
          underline: true,
          stringValue: 'hello',
        },
        fillColor: '#1f2937',
      }),
    ).toEqual({
      backgroundColor: '#1f2937',
      color: '#f9fafb',
      font: 'italic 700 14px "JetBrainsMono Nerd Font","JetBrains Mono",monospace',
      fontSize: 14,
      underline: true,
    })
  })

  test('falls back to a white editor surface when the cell has no fill', () => {
    expect(
      getEditorPresentation({
        renderCell: {
          kind: 'string',
          displayText: 'hello',
          copyText: 'hello',
          align: 'left',
          wrap: false,
          color: '#202124',
          font: '400 13px "JetBrainsMono Nerd Font","JetBrains Mono",monospace',
          fontSize: 13,
          underline: false,
          stringValue: 'hello',
        },
      }),
    ).toEqual({
      backgroundColor: '#ffffff',
      color: '#202124',
      font: '400 13px "JetBrainsMono Nerd Font","JetBrains Mono",monospace',
      fontSize: 13,
      underline: false,
    })
  })

  test('matches the edited cell bounds exactly instead of expanding the overlay frame', () => {
    expect(
      getOverlayStyle(true, {
        x: 240,
        y: 96,
        width: 104,
        height: 22,
      }),
    ).toEqual({
      height: 22,
      left: 240,
      position: 'fixed',
      top: 96,
      width: 104,
    })
  })
})
