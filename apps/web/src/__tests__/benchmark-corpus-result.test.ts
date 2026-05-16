import { describe, expect, it } from 'vitest'
import { isInstallBenchmarkCorpusResult } from '../benchmark-corpus-result.js'

describe('benchmark corpus result validation', () => {
  const validResult = {
    id: 'dense-mixed-100k',
    materializedCellCount: 100_000,
    primaryViewport: {
      sheetName: 'Sheet1',
      rowStart: 0,
      rowEnd: 31,
      colStart: 0,
      colEnd: 15,
    },
  }

  it('accepts well-formed install results', () => {
    expect(isInstallBenchmarkCorpusResult(validResult)).toBe(true)
  })

  it('rejects malformed viewport and count payloads', () => {
    expect(isInstallBenchmarkCorpusResult({ ...validResult, id: '' })).toBe(false)
    expect(isInstallBenchmarkCorpusResult({ ...validResult, materializedCellCount: 0 })).toBe(false)
    expect(isInstallBenchmarkCorpusResult({ ...validResult, materializedCellCount: Number.NaN })).toBe(false)
    expect(
      isInstallBenchmarkCorpusResult({
        ...validResult,
        primaryViewport: { ...validResult.primaryViewport, sheetName: '' },
      }),
    ).toBe(false)
    expect(
      isInstallBenchmarkCorpusResult({
        ...validResult,
        primaryViewport: { ...validResult.primaryViewport, rowStart: -1 },
      }),
    ).toBe(false)
    expect(
      isInstallBenchmarkCorpusResult({
        ...validResult,
        primaryViewport: { ...validResult.primaryViewport, rowStart: 5, rowEnd: 4 },
      }),
    ).toBe(false)
    expect(
      isInstallBenchmarkCorpusResult({
        ...validResult,
        primaryViewport: { ...validResult.primaryViewport, colStart: 3, colEnd: 2 },
      }),
    ).toBe(false)
  })
})
