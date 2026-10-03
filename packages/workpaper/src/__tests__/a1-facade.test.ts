import { describe, expect, it } from 'vitest'

import { buildA1WorkPaper, restoreA1WorkPaper } from '../index.js'

describe('A1 WorkPaper facade', () => {
  it('edits a sheet-qualified input and returns formula readback restore proof', () => {
    const book = buildPricingBook()

    try {
      const proof = book.setCellAndReadback('Inputs!B2', 48, {
        readbackRange: 'Summary!B2',
      })

      expect(proof.verified).toBe(true)
      expect(proof.editedCell).toBe('Inputs!B2')
      expect(proof.readbackRange).toBe('Summary!B2')
      expect(proof.beforeReadback.displayValues).toEqual([['48000']])
      expect(proof.afterReadback.displayValues).toEqual([['57600']])
      expect(proof.afterReadback.values[0]?.[0]).toMatchObject({ value: 57600 })
      expect(proof.afterReadback.formulaDiagnostics).toEqual([])
      expect(proof.restoredReadback.displayValues).toEqual(proof.afterReadback.displayValues)
      expect(proof.persistedDocumentBytes).toBeGreaterThan(100)
      expect(proof.checks).toMatchObject({
        readbackChanged: true,
        computedReadbackChanged: true,
        editedFormulaReadbackChanged: false,
        readbackIncludesEditedCell: false,
        readbackContainsOnlyEditedCell: false,
        restoredReadbackMatchesAfter: true,
        blockingFormulaDiagnosticCount: 0,
        previousSerialized: 40,
        newSerialized: 48,
      })
    } finally {
      book.dispose()
    }
  })

  it('rejects writes outside the configured writable sheet list', () => {
    const book = buildPricingBook({ writableSheets: ['Inputs'] })

    try {
      expect(() =>
        book.setCellAndReadback('Summary!B2', '=Inputs!B2*Inputs!B3*2', {
          readbackRange: 'Summary!B2',
        }),
      ).toThrow('Sheet "Summary" is not writable')
    } finally {
      book.dispose()
    }
  })

  it('requires sheet-qualified addresses unless a default sheet is configured', () => {
    const strictBook = buildPricingBook()
    const defaultedBook = buildPricingBook({ defaultSheetName: 'Summary' })

    try {
      expect(() => strictBook.display('B2')).toThrow('Sheet-qualified WorkPaper cell address required: B2')
      expect(defaultedBook.display('B2')).toBe('48000')
    } finally {
      strictBook.dispose()
      defaultedBook.dispose()
    }
  })

  it('writes formulas through A1 addresses and preserves formula serialization', () => {
    const book = buildPricingBook()

    try {
      expect(book.validateFormula('=Inputs!B2*Inputs!B3*2')).toBe(true)
      const proof = book.setCellAndReadback('Summary!B2', '=Inputs!B2*Inputs!B3*2', {
        readbackRange: 'Summary!B2',
      })

      expect(proof.verified).toBe(true)
      expect(proof.after.formula).toBe('=Inputs!B2*Inputs!B3*2')
      expect(proof.after.serialized).toBe('=Inputs!B2*Inputs!B3*2')
      expect(proof.afterReadback.displayValues).toEqual([['96000']])
      expect(proof.restoredReadback.serialized).toEqual([['=Inputs!B2*Inputs!B3*2']])
      expect(proof.checks.editedFormulaReadbackChanged).toBe(true)
    } finally {
      book.dispose()
    }
  })

  it('restores serialized A1 WorkPaper documents', () => {
    const source = buildPricingBook()

    try {
      source.set('Inputs!B2', 60)
      const restored = restoreA1WorkPaper(source.serialize())
      try {
        expect(restored.display('Summary!B2')).toBe('72000')
        expect(restored.readCell('Inputs!B2').serialized).toBe(60)
        expect(restored.serialize()).toContain('bilig.headless.work-paper.document.v1')
      } finally {
        restored.dispose()
      }
    } finally {
      source.dispose()
    }
  })

  it('provides concise multi-cell and JSON aliases for agent workflows', () => {
    const book = buildPricingBook()

    try {
      book.setMany({
        'Inputs!B2': 50,
        'Inputs!B3': 1500,
      })

      expect(book.display('Summary!B2')).toBe('75000')
      expect(Object.keys(book.readMany(['Inputs!B2', 'Summary!B2']))).toEqual(['Inputs!B2', 'Summary!B2'])

      const proof = book.editAndReadback('Inputs!B2', 60, {
        readbackRange: 'Summary!B2',
      })
      expect(proof.verified).toBe(true)
      expect(proof.afterReadback.displayValues).toEqual([['90000']])

      const restored = book.restoreJson(book.saveJson())
      try {
        expect(restored.display('Summary!B2')).toBe('90000')
      } finally {
        restored.dispose()
      }
    } finally {
      book.dispose()
    }
  })

  it('edits multiple cells as one proof and compares typed readback values', () => {
    const book = buildPricingBook()

    try {
      const proof = book.editManyAndReadback(
        {
          'Inputs!B2': 48,
          'Inputs!B3': 1500,
        },
        { readbackRange: 'Summary!B2' },
      )

      expect(proof.verified).toBe(true)
      expect(proof.editedCells).toEqual(['Inputs!B2', 'Inputs!B3'])
      expect(proof.before['Inputs!B2']?.serialized).toBe(40)
      expect(proof.before['Inputs!B3']?.serialized).toBe(1200)
      expect(proof.after['Inputs!B2']?.serialized).toBe(48)
      expect(proof.after['Inputs!B3']?.serialized).toBe(1500)
      expect(proof.afterReadback.displayValues).toEqual([['72000']])
      expect(proof.afterReadback.values[0]?.[0]).toMatchObject({ value: 72000 })
      expect(proof.checks).toMatchObject({
        readbackChanged: true,
        computedReadbackChanged: true,
        editedFormulaReadbackChanged: false,
        readbackIncludesEditedCells: false,
        readbackContainsOnlyEditedCells: false,
        blockingFormulaDiagnosticCount: 0,
      })
      expect(proof.checks.previousSerialized).toEqual({
        'Inputs!B2': 40,
        'Inputs!B3': 1200,
      })
      expect(proof.checks.newSerialized).toEqual({
        'Inputs!B2': 48,
        'Inputs!B3': 1500,
      })
    } finally {
      book.dispose()
    }
  })

  it('does not verify proof when readback only covers edited input cells', () => {
    const book = buildPricingBook()

    try {
      const proof = book.editAndReadback('Inputs!B2', 48, {
        readbackRange: 'Inputs!B2',
      })

      expect(proof.verified).toBe(false)
      expect(proof.checks.readbackChanged).toBe(true)
      expect(proof.checks.computedReadbackChanged).toBe(false)
      expect(proof.checks.editedFormulaReadbackChanged).toBe(false)
      expect(proof.checks.readbackIncludesEditedCell).toBe(true)
      expect(proof.checks.readbackContainsOnlyEditedCell).toBe(true)
    } finally {
      book.dispose()
    }
  })

  it('keeps multi-cell edits atomic when a later target is not writable', () => {
    const book = buildPricingBook({ writableSheets: ['Inputs'] })

    try {
      expect(() =>
        book.setMany({
          'Inputs!B2': 99,
          'Summary!B2': 123,
        }),
      ).toThrow('Sheet "Summary" is not writable')
      expect(book.readCell('Inputs!B2').serialized).toBe(40)
      expect(book.display('Summary!B2')).toBe('48000')
    } finally {
      book.dispose()
    }
  })

  it('surfaces formula diagnostics and blocks provider-backed readback proofs', () => {
    const book = buildPricingBook()

    try {
      const proof = book.editAndReadback('Summary!B2', '=IMPORTRANGE("sheet","Summary!B2")', {
        readbackRange: 'Summary!B2',
      })

      expect(proof.verified).toBe(false)
      expect(proof.afterReadback.formulaDiagnostics).toEqual(
        expect.arrayContaining([
          expect.objectContaining({
            code: 'provider-backed-adapter-missing',
            functionName: 'IMPORTRANGE',
            severity: 'error',
          }),
        ]),
      )
      expect(proof.checks.blockingFormulaDiagnosticCount).toBeGreaterThan(0)
      expect(proof.checks.formulaDiagnostics).toEqual(proof.afterReadback.formulaDiagnostics)
    } finally {
      book.dispose()
    }
  })
})

function buildPricingBook(options: Parameters<typeof buildA1WorkPaper>[2] = {}) {
  return buildA1WorkPaper(
    {
      Inputs: [
        ['Metric', 'Value'],
        ['Units', 40],
        ['Price', 1200],
      ],
      Summary: [
        ['Metric', 'Value'],
        ['Revenue', '=Inputs!B2*Inputs!B3'],
      ],
    },
    undefined,
    options,
  )
}
