import { describe, expect, it } from 'vitest'
import { applyStylePatch, clearStyleFields, cloneCellStyleRecord, normalizeCellStylePatch } from '../engine-style-utils.js'

describe('engine style utils', () => {
  it('deep clones nested style sections', () => {
    const original = {
      id: 'style-1',
      font: { family: 'Inter', bold: true },
      borders: {
        top: { style: 'solid', weight: 'thin', color: '#111111' },
      },
    } as const

    const cloned = cloneCellStyleRecord(original)
    cloned.font!.family = 'IBM Plex Sans'
    cloned.borders!.top!.color = '#222222'

    expect(original.font.family).toBe('Inter')
    expect(original.borders.top.color).toBe('#111111')
  })

  it('normalizes sparse patches and preserves explicit clears', () => {
    expect(
      normalizeCellStylePatch({
        fill: { backgroundColor: null },
        font: { family: 'Inter', size: undefined },
        borders: {
          top: { style: 'solid', weight: 'thin', color: '#111111' },
          right: undefined,
        },
      }),
    ).toEqual({
      fill: { backgroundColor: null },
      font: { family: 'Inter' },
      borders: {
        top: { style: 'solid', weight: 'thin', color: '#111111' },
      },
    })
  })

  it('drops invalid border patches while preserving sibling sections', () => {
    expect(
      applyStylePatch(
        {
          font: { family: 'Inter', bold: true },
          borders: {
            left: { style: 'double', weight: 'medium', color: '#222222' },
          },
        },
        {
          alignment: { horizontal: 'center' },
          borders: {
            left: { style: 'solid', weight: 'thin', color: null },
          },
        },
      ),
    ).toEqual({
      font: { family: 'Inter', bold: true },
      alignment: { horizontal: 'center' },
    })
  })

  it('clears selected fields without dropping sibling values', () => {
    expect(
      clearStyleFields(
        {
          font: { family: 'Inter', bold: true },
          alignment: { horizontal: 'right', wrap: true },
          borders: {
            top: { style: 'solid', weight: 'thin', color: '#111111' },
            left: { style: 'double', weight: 'medium', color: '#222222' },
          },
        },
        ['fontBold', 'alignmentWrap', 'borderTop'],
      ),
    ).toEqual({
      font: { family: 'Inter' },
      alignment: { horizontal: 'right' },
      borders: {
        left: { style: 'double', weight: 'medium', color: '#222222' },
      },
    })
  })
})
