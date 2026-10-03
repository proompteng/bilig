import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { ErrorCode, ValueTag } from '@bilig/protocol'
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { recalculateXlsxFileToFile } from '../index.js'

const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

describe('@bilig/xlsx-formula-recalc native aggregates', () => {
  it('hydrates high-index shared string targets without replaying the shared-string buffer', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-shared-string-buffer-'))
    try {
      const sourcePath = join(tempDir, 'shared-string-buffer.xlsx')
      const outputPath = join(tempDir, 'shared-string-buffer.recalculated.xlsx')
      writeFileSync(sourcePath, buildHighIndexSharedStringWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        reads: ['Data!A1'],
      })

      expect(result.reads['Data!A1']).toMatchObject({ value: 'target-shared-string' })
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(0)
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('loads xlsx-fixture-corpus style dependency rows and recalculates SUM ranges through the native kernel', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-sum-range-'))
    try {
      const sourcePath = join(tempDir, 'public-sum-range.xlsx')
      const outputPath = join(tempDir, 'public-sum-range.recalculated.xlsx')
      writeFileSync(sourcePath, buildPublicSumRangeWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        reads: ['Table 4!A1'],
      })

      expect(readNumber(result.reads['Table 4!A1'])).toBe(425)
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.targetRowCount).toBe(5)
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(1)
      expect(result.diagnostics?.formulaCounts.nativeKernelFormulaCellCount).toBe(1)
      expect(result.diagnostics?.formulaCounts.nativeKernelBatchCount).toBe(1)
      const outputBytes = readFileSync(outputPath)
      const sheetXml = strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain('<c r="A1"><f>SUM(B5:B8)</f><v>425</v></c>')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('evaluates SUM and COUNTA formulas over scanned xlsx-fixture-corpus ranges', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-sum-counta-range-'))
    try {
      const sourcePath = join(tempDir, 'public-sum-counta-range.xlsx')
      const outputPath = join(tempDir, 'public-sum-counta-range.recalculated.xlsx')
      writeFileSync(sourcePath, buildPublicSumCountaRangeWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        reads: ['Data!A2', 'Data!A6'],
      })

      expect(readNumber(result.reads['Data!A2'])).toBe(12)
      expect(readNumber(result.reads['Data!A6'])).toBe(2)
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.targetRowCount).toBe(5)
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(2)
      expect(result.diagnostics?.formulaCounts.unsupportedFormulaCellCount).toBe(0)
      expect(result.diagnostics?.formulaCounts.patchedFormulaCacheCount).toBe(2)
      const outputBytes = readFileSync(outputPath)
      const sheetXml = strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain('<c r="A2"><f>SUM(B2:B5)</f><v>12</v></c>')
      expect(sheetXml).toContain('<c r="A6"><f>IF(COUNTA(B6:D6)=0,"",COUNTA(C6:E6))</f><v>2</v></c>')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('evaluates ROUND and AVERAGE formulas from public financial forecasts', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-round-average-'))
    try {
      const sourcePath = join(tempDir, 'public-round-average.xlsx')
      const outputPath = join(tempDir, 'public-round-average.recalculated.xlsx')
      writeFileSync(sourcePath, buildPublicRoundAverageWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        reads: ['Data!A2', 'Data!A3', 'Data!A4', 'Data!A5', 'Data!A6'],
      })

      expect(readNumber(result.reads['Data!A2'])).toBe(3.1)
      expect(readNumber(result.reads['Data!A3'])).toBe(-2)
      expect(readNumber(result.reads['Data!A4'])).toBe(1200)
      expect(readBoolean(result.reads['Data!A5'])).toBe(true)
      expect(readNumber(result.reads['Data!A6'])).toBe(0)
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.targetRowCount).toBe(9)
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(5)
      expect(result.diagnostics?.formulaCounts.unsupportedFormulaCellCount).toBe(0)
      expect(result.diagnostics?.formulaCounts.patchedFormulaCacheCount).toBe(5)
      const outputBytes = readFileSync(outputPath)
      const sheetXml = strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain('<c r="A2"><f>ROUND(AVERAGE(B2:B6),1)</f><v>3.1</v></c>')
      expect(sheetXml).toContain('<c r="A3"><f>ROUND(B7,0)</f><v>-2</v></c>')
      expect(sheetXml).toContain('<c r="A4"><f>ROUND(B8,-2)</f><v>1200</v></c>')
      expect(sheetXml).toContain('<c r="A5" t="b"><f>ISERROR(B9/B10)</f><v>1</v></c>')
      expect(sheetXml).toContain('<c r="A6"><f>IF(ISERROR(B9/B10),0,B9/B10)</f><v>0</v></c>')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('propagates formula error values through public financial forecast formulas', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-error-propagation-'))
    try {
      const sourcePath = join(tempDir, 'public-error-propagation.xlsx')
      const outputPath = join(tempDir, 'public-error-propagation.recalculated.xlsx')
      writeFileSync(sourcePath, buildPublicErrorPropagationWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        reads: ['Data!A2', 'Data!A3', 'Data!A4', 'Data!A5'],
      })

      expect(readErrorCode(result.reads['Data!A2'])).toBe(ErrorCode.Div0)
      expect(readErrorCode(result.reads['Data!A3'])).toBe(ErrorCode.Ref)
      expect(readErrorCode(result.reads['Data!A4'])).toBe(ErrorCode.Value)
      expect(readErrorCode(result.reads['Data!A5'])).toBe(ErrorCode.NA)
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(4)
      expect(result.diagnostics?.formulaCounts.unsupportedFormulaCellCount).toBe(0)
      expect(result.diagnostics?.formulaCounts.patchedFormulaCacheCount).toBe(4)
      const outputBytes = readFileSync(outputPath)
      const sheetXml = strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain('<c r="A2" t="e"><f>ROUND(B2/C2,1)</f><v>#DIV/0!</v></c>')
      expect(sheetXml).toContain('<c r="A3" t="e"><f>IF(B3=#REF!,"-",ROUND(B3/#REF!,1))</f><v>#REF!</v></c>')
      expect(sheetXml).toContain('<c r="A4" t="e"><f>SUM(B4:C4)</f><v>#VALUE!</v></c>')
      expect(sheetXml).toContain('<c r="A5" t="e"><f>AVERAGE(B5:C5)</f><v>#N/A</v></c>')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('caches blank direct formula references as zero on the streaming-native path', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-blank-direct-ref-'))
    try {
      const sourcePath = join(tempDir, 'public-blank-direct-ref.xlsx')
      const outputPath = join(tempDir, 'public-blank-direct-ref.recalculated.xlsx')
      writeFileSync(sourcePath, buildPublicBlankDirectReferenceWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        reads: ['Summary!A2', 'Summary!A3'],
      })

      expect(readNumber(result.reads['Summary!A2'])).toBe(0)
      expect(readNumber(result.reads['Summary!A3'])).toBe(25)
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(2)
      expect(result.diagnostics?.formulaCounts.unsupportedFormulaCellCount).toBe(0)
      expect(result.diagnostics?.formulaCounts.patchedFormulaCacheCount).toBe(2)
      const outputBytes = readFileSync(outputPath)
      const summaryXml = strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(summaryXml).toContain('<c r="A2"><f>\'Source\'!B2</f><v>0</v></c>')
      expect(summaryXml).toContain('<c r="A3"><f>A2+25</f><v>25</v></c>')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('evaluates exponentiation and REPT formulas from public financial forecasts', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-exponent-rept-'))
    try {
      const sourcePath = join(tempDir, 'public-exponent-rept.xlsx')
      const outputPath = join(tempDir, 'public-exponent-rept.recalculated.xlsx')
      writeFileSync(sourcePath, buildPublicExponentReptWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        reads: ['Data!A2', 'Data!A3', 'Data!A4', 'Data!A5', 'Data!A6', 'Data!A7', 'Data!A8'],
      })

      expect(readNumber(result.reads['Data!A2'])).toBe(1_000_000)
      expect(readNumber(result.reads['Data!A3'])).toBe(16)
      expect(readString(result.reads['Data!A4'])).toBe('----')
      expect(readString(result.reads['Data!A5'])).toBe('Q1Q1Q1')
      expect(readErrorCode(result.reads['Data!A6'])).toBe(ErrorCode.Value)
      expect(readString(result.reads['Data!A7'])).toBe('Carbon dioxide equivalent emissions ($/tonne)...')
      expect(readErrorCode(result.reads['Data!A8'])).toBe(ErrorCode.Ref)
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.formulaCounts.unsupportedFormulaCellCount).toBe(0)
      expect(result.diagnostics?.formulaCounts.patchedFormulaCacheCount).toBe(7)
      const outputBytes = readFileSync(outputPath)
      const sheetXml = strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain('<c r="A2"><f>10^6</f><v>1000000</v></c>')
      expect(sheetXml).toContain('<c r="A3"><f>B3^C3</f><v>16</v></c>')
      expect(sheetXml).toContain('<c r="A4" t="str"><f>REPT(&quot;-&quot;,4)</f><v>----</v></c>')
      expect(sheetXml).toContain('<c r="A5" t="str"><f>REPT(B5,C5)</f><v>Q1Q1Q1</v></c>')
      expect(sheetXml).toContain('<c r="A6" t="e"><f>REPT(&quot;x&quot;,B6)</f><v>#VALUE!</v></c>')
      expect(sheetXml).toContain(
        '<c r="A7" t="str"><f>+&quot;Carbon dioxide equivalent emissions ($/tonne)&quot;&amp;REPT(&quot;.&quot;,3)</f><v>Carbon dioxide equivalent emissions ($/tonne)...</v></c>',
      )
      expect(sheetXml).toContain('<c r="A8" t="e"><f>B8*1000000/#REF!</f><v>#REF!</v></c>')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('evaluates exact INDEX MATCH header lookups from public forecasts', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-index-match-'))
    try {
      const sourcePath = join(tempDir, 'public-index-match.xlsx')
      const outputPath = join(tempDir, 'public-index-match.recalculated.xlsx')
      writeFileSync(sourcePath, buildPublicIndexMatchWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        reads: ['Data!A2'],
      })

      expect(readNumber(result.reads['Data!A2'])).toBe(30)
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.targetRowCount).toBe(5)
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(1)
      expect(result.diagnostics?.formulaCounts.unsupportedFormulaCellCount).toBe(0)
      expect(result.diagnostics?.formulaCounts.patchedFormulaCacheCount).toBe(1)
      const outputBytes = readFileSync(outputPath)
      const sheetXml = strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain('<c r="A2"><f>INDEX(B5:B8,MATCH(B6,B5:B8,0)+1,1)</f><v>30</v></c>')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('loads same-sheet scalar dependency rows for xlsx-fixture-corpus formulas', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-scalar-dependency-'))
    try {
      const sourcePath = join(tempDir, 'public-scalar-dependency.xlsx')
      const outputPath = join(tempDir, 'public-scalar-dependency.recalculated.xlsx')
      writeFileSync(sourcePath, buildPublicScalarDependencyWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        reads: ['Budget!C7'],
      })

      expect(readNumber(result.reads['Budget!C7'])).toBe(250)
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.targetRowCount).toBe(3)
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(1)
      expect(result.diagnostics?.formulaCounts.unsupportedFormulaCellCount).toBe(0)
      expect(result.diagnostics?.formulaCounts.patchedFormulaCacheCount).toBe(1)
      const outputBytes = readFileSync(outputPath)
      const sheetXml = strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain('<c r="C7"><f>B6+B8*2</f><v>250</v></c>')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('loads cross-sheet scalar dependency rows for xlsx-fixture-corpus formulas', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-cross-sheet-scalar-'))
    try {
      const sourcePath = join(tempDir, 'public-cross-sheet-scalar.xlsx')
      const outputPath = join(tempDir, 'public-cross-sheet-scalar.recalculated.xlsx')
      writeFileSync(sourcePath, buildPublicCrossSheetScalarDependencyWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        reads: ['Budget!C7'],
      })

      expect(readNumber(result.reads['Budget!C7'])).toBe(160)
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.targetRowCount).toBe(3)
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(1)
      expect(result.diagnostics?.formulaCounts.unsupportedFormulaCellCount).toBe(0)
      expect(result.diagnostics?.formulaCounts.patchedFormulaCacheCount).toBe(1)
      const outputBytes = readFileSync(outputPath)
      const sheetXml = strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet2.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain("<c r=\"C7\"><f>'Data Sheet'!B2+'Data Sheet'!C7*2</f><v>160</v></c>")
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('evaluates exact VLOOKUP formulas over scanned xlsx-fixture-corpus table rows', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-vlookup-'))
    try {
      const sourcePath = join(tempDir, 'public-vlookup.xlsx')
      const outputPath = join(tempDir, 'public-vlookup.recalculated.xlsx')
      writeFileSync(sourcePath, buildPublicVlookupWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        reads: ['Budget!D3'],
      })

      expect(readNumber(result.reads['Budget!D3'])).toBe(250)
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.targetRowCount).toBe(4)
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(1)
      expect(result.diagnostics?.formulaCounts.unsupportedFormulaCellCount).toBe(0)
      expect(result.diagnostics?.formulaCounts.patchedFormulaCacheCount).toBe(1)
      expect(result.diagnostics?.formulaCounts.nativeKernelFormulaCellCount).toBe(1)
      expect(result.diagnostics?.formulaCounts.nativeKernelBatchCount).toBe(1)
      const outputBytes = readFileSync(outputPath)
      const sheetXml = strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet2.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain('<c r="D3"><f>VLOOKUP(B2,\'Lookup Table\'!$B$9:$D$10,2,FALSE)</f><v>250</v></c>')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('evaluates omitted approximate VLOOKUP formulas over scanned table rows', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-approximate-vlookup-'))
    try {
      const sourcePath = join(tempDir, 'public-approximate-vlookup.xlsx')
      const outputPath = join(tempDir, 'public-approximate-vlookup.recalculated.xlsx')
      writeFileSync(sourcePath, buildPublicApproximateVlookupWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        reads: ['Budget!D3'],
      })

      expect(readNumber(result.reads['Budget!D3'])).toBe(250)
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.targetRowCount).toBe(5)
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(1)
      expect(result.diagnostics?.formulaCounts.unsupportedFormulaCellCount).toBe(0)
      expect(result.diagnostics?.formulaCounts.patchedFormulaCacheCount).toBe(1)
      expect(result.diagnostics?.formulaCounts.nativeKernelFormulaCellCount).toBe(1)
      expect(result.diagnostics?.formulaCounts.nativeKernelBatchCount).toBe(1)
      const outputBytes = readFileSync(outputPath)
      const sheetXml = strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet2.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain('<c r="D3"><f>VLOOKUP(B2,\'Lookup Table\'!$B$9:$D$11,2)</f><v>250</v></c>')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('evaluates VLOOKUP formulas with INDIRECT scalar cell-range table rows', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-indirect-vlookup-'))
    try {
      const sourcePath = join(tempDir, 'public-indirect-vlookup.xlsx')
      const outputPath = join(tempDir, 'public-indirect-vlookup.recalculated.xlsx')
      writeFileSync(sourcePath, buildPublicIndirectVlookupWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        reads: ['Budget!D3'],
      })

      expect(readNumber(result.reads['Budget!D3'])).toBe(250)
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.targetRowCount).toBe(6)
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(1)
      expect(result.diagnostics?.formulaCounts.unsupportedFormulaCellCount).toBe(0)
      expect(result.diagnostics?.formulaCounts.patchedFormulaCacheCount).toBe(1)
      const outputBytes = readFileSync(outputPath)
      const sheetXml = strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet2.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain('<c r="D3"><f>VLOOKUP(B2,INDIRECT("\'"&amp;C$4&amp;"\'!"&amp;"$B$9:$D$11"),2)</f><v>250</v></c>')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('fails closed when external workbook references have no cached values', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-external-reference-'))
    try {
      const sourcePath = join(tempDir, 'external-reference.xlsx')
      const outputPath = join(tempDir, 'external-reference.recalculated.xlsx')
      writeFileSync(sourcePath, buildExternalReferenceWorkbook())

      await expect(
        recalculateXlsxFileToFile(sourcePath, {
          outputPath,
          engine: 'streaming-native',
          reads: ['Budget!B2'],
        }),
      ).rejects.toMatchObject({
        diagnostics: expect.objectContaining({
          engineMode: 'streaming-native',
          fallbackUsed: false,
          unsupportedReason: expect.stringContaining('external workbook cache sheet is missing'),
        }),
      })
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('evaluates external workbook references from cached external-link values', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-cached-external-reference-'))
    try {
      const sourcePath = join(tempDir, 'cached-external-reference.xlsx')
      const outputPath = join(tempDir, 'cached-external-reference.recalculated.xlsx')
      writeFileSync(sourcePath, buildCachedExternalReferenceWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        reads: ['Budget!B2', 'Budget!C2', 'Budget!D2', 'Budget!E2'],
      })

      expect(readNumber(result.reads['Budget!B2'])).toBeCloseTo(1460.21607666, 9)
      expect(readNumber(result.reads['Budget!C2'])).toBeCloseTo(99.81585952, 9)
      expect(readNumber(result.reads['Budget!D2'])).toBe(2021)
      expect(readNumber(result.reads['Budget!E2'])).toBeCloseTo(1460.59744118, 9)
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(4)
      expect(result.diagnostics?.formulaCounts.unsupportedFormulaCellCount).toBe(0)
      expect(result.diagnostics?.formulaCounts.patchedFormulaCacheCount).toBe(4)
      const outputBytes = readFileSync(outputPath)
      const sheetXml = strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain('<f>[1]QTR_detailed!$B130/10^6</f>')
      expect(sheetXml).toContain("<f>+'[1]BS &amp; Op Stat'!$D$5</f>")
      expect(sheetXml).toContain('<v>2021</v>')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('hydrates external workbook companion values on the streaming-native file path', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-companion-external-reference-'))
    try {
      const sourcePath = join(tempDir, 'cached-external-reference.xlsx')
      const outputPath = join(tempDir, 'cached-external-reference.recalculated.xlsx')
      writeFileSync(sourcePath, buildCachedExternalReferenceWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        externalWorkbooks: [
          {
            fileName: 'source.xlsx',
            target: 'file:///tmp/source.xlsx',
            bytes: buildExternalCompanionReferenceWorkbook(),
          },
        ],
        reads: ['Budget!B2', 'Budget!C2', 'Budget!D2', 'Budget!E2'],
      })

      expect(readNumber(result.reads['Budget!B2'])).toBe(2000)
      expect(readNumber(result.reads['Budget!C2'])).toBe(120)
      expect(readNumber(result.reads['Budget!D2'])).toBe(2024)
      expect(readNumber(result.reads['Budget!E2'])).toBe(2050)
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(4)
      expect(result.diagnostics?.formulaCounts.unsupportedFormulaCellCount).toBe(0)
      expect(result.diagnostics?.formulaCounts.patchedFormulaCacheCount).toBe(4)
      const outputBytes = readFileSync(outputPath)
      const sheetXml = strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain('<c r="B2"><f>[1]QTR_detailed!$B130/10^6</f><v>2000</v></c>')
      expect(sheetXml).toContain('<c r="C2"><f>SUM([1]QTR_detailed!$R130:$T130)/10^6</f><v>120</v></c>')
      expect(sheetXml).toContain('<c r="D2"><f>+\'[1]BS &amp; Op Stat\'!$D$5</f><v>2024</v></c>')
      expect(sheetXml).toContain('<c r="E2"><f>SUM([1]QTR_detailed!$B130,[1]QTR_detailed!$T130,)/10^6</f><v>2050</v></c>')
      const externalLinkXml = strFromU8(unzipSync(outputBytes)['xl/externalLinks/externalLink1.xml'] ?? new Uint8Array())
      expect(externalLinkXml).toContain('<cell r="B130"><v>2000000000</v></cell>')
      expect(externalLinkXml).toContain('<cell r="R130"><v>30000000</v></cell>')
      expect(externalLinkXml).toContain('<cell r="S130"><v>40000000</v></cell>')
      expect(externalLinkXml).toContain('<cell r="T130"><v>50000000</v></cell>')
      expect(externalLinkXml).toContain('<cell r="D5"><v>2024</v></cell>')
      expect(externalLinkXml).not.toContain('<cell r="B130"><v>1460216076.6600001</v></cell>')
      expect(externalLinkXml).not.toContain('<cell r="R130"><v>99434495</v></cell>')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })
})

function readNumber(value: unknown): number {
  if (typeof value === 'object' && value !== null && 'value' in value && typeof value.value === 'number') {
    return value.value
  }
  throw new Error(`Expected numeric cell value, received ${JSON.stringify(value)}`)
}

function readString(value: unknown): string {
  if (typeof value === 'object' && value !== null && 'value' in value && typeof value.value === 'string') {
    return value.value
  }
  throw new Error(`Expected string cell value, received ${JSON.stringify(value)}`)
}

function readErrorCode(value: unknown): ErrorCode {
  if (
    typeof value === 'object' &&
    value !== null &&
    'tag' in value &&
    value.tag === ValueTag.Error &&
    'code' in value &&
    typeof value.code === 'number'
  ) {
    return value.code
  }
  throw new Error(`Expected error cell value, received ${JSON.stringify(value)}`)
}

function readBoolean(value: unknown): boolean {
  if (typeof value === 'object' && value !== null && 'value' in value && typeof value.value === 'boolean') {
    return value.value
  }
  throw new Error(`Expected boolean cell value, received ${JSON.stringify(value)}`)
}

function buildHighIndexSharedStringWorkbook(): Uint8Array {
  const sharedStrings = Array.from({ length: 74 }, (_value, index) => `<si><t>unused-${String(index)}</t></si>`)
  sharedStrings.push('<si><t>target-shared-string</t></si>')
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdWorkbook" Type="${officeRelationshipNamespace}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    'xl/sharedStrings.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="75" uniqueCount="75">${sharedStrings.join('')}</sst>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData>
    <row r="1"><c r="A1" t="s"><v>74</v></c></row>
  </sheetData>
</worksheet>`),
  })
}

function buildPublicSumRangeWorkbook(): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets><sheet name="Table 4" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdWorkbook" Type="${officeRelationshipNamespace}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B8"/>
  <sheetData>
    <row r="1"><c r="A1"><f>SUM(B5:B8)</f><v>0</v></c></row>
    <row r="5"><c r="B5"><v>100</v></c></row>
    <row r="6"><c r="B6"><v>125</v></c></row>
    <row r="7"><c r="B7"><v>75</v></c></row>
    <row r="8"><c r="B8"><v>125</v></c></row>
  </sheetData>
</worksheet>`),
  })
}

function buildPublicSumCountaRangeWorkbook(): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdWorkbook" Type="${officeRelationshipNamespace}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A2:E6"/>
  <sheetData>
    <row r="2"><c r="A2"><f>SUM(B2:B5)</f><v>0</v></c><c r="B2"><v>5</v></c></row>
    <row r="3"><c r="B3" t="inlineStr"><is><t>ignored</t></is></c></row>
    <row r="4"/>
    <row r="5"><c r="B5"><v>7</v></c></row>
    <row r="6"><c r="A6"><f>IF(COUNTA(B6:D6)=0,"",COUNTA(C6:E6))</f><v>0</v></c><c r="C6"><v>1</v></c><c r="D6" t="inlineStr"><is><t>x</t></is></c></row>
  </sheetData>
</worksheet>`),
  })
}

function buildPublicRoundAverageWorkbook(): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdWorkbook" Type="${officeRelationshipNamespace}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A2:B10"/>
  <sheetData>
    <row r="2"><c r="A2"><f>ROUND(AVERAGE(B2:B6),1)</f><v>0</v></c><c r="B2"><v>2</v></c></row>
    <row r="3"><c r="A3"><f>ROUND(B7,0)</f><v>0</v></c><c r="B3"><v>3</v></c></row>
    <row r="4"><c r="A4"><f>ROUND(B8,-2)</f><v>0</v></c><c r="B4" t="inlineStr"><is><t>ignored</t></is></c></row>
    <row r="5"><c r="A5"><f>ISERROR(B9/B10)</f><v>0</v></c></row>
    <row r="6"><c r="A6"><f>IF(ISERROR(B9/B10),0,B9/B10)</f><v>99</v></c><c r="B6"><v>4.4</v></c></row>
    <row r="7"><c r="B7"><v>-1.5</v></c></row>
    <row r="8"><c r="B8"><v>1234.56</v></c></row>
    <row r="9"><c r="B9"><v>10</v></c></row>
    <row r="10"><c r="B10"><v>0</v></c></row>
  </sheetData>
</worksheet>`),
  })
}

function buildPublicErrorPropagationWorkbook(): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdWorkbook" Type="${officeRelationshipNamespace}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A2:C5"/>
  <sheetData>
    <row r="2"><c r="A2"><f>ROUND(B2/C2,1)</f><v>stale</v></c><c r="B2"><v>10</v></c><c r="C2"><v>0</v></c></row>
    <row r="3"><c r="A3"><f>IF(B3=#REF!,"-",ROUND(B3/#REF!,1))</f><v>stale</v></c><c r="B3"><v>10</v></c></row>
    <row r="4"><c r="A4"><f>SUM(B4:C4)</f><v>stale</v></c><c r="B4"><v>10</v></c><c r="C4" t="e"><v>#VALUE!</v></c></row>
    <row r="5"><c r="A5"><f>AVERAGE(B5:C5)</f><v>stale</v></c><c r="B5"><v>10</v></c><c r="C5" t="e"><v>#N/A</v></c></row>
  </sheetData>
</worksheet>`),
  })
}

function buildPublicBlankDirectReferenceWorkbook(): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets>
    <sheet name="Summary" sheetId="1" r:id="rId1"/>
    <sheet name="Source" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`),
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdWorkbook" Type="${officeRelationshipNamespace}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A2:A3"/>
  <sheetData>
    <row r="2"><c r="A2"><f>'Source'!B2</f><v>stale</v></c></row>
    <row r="3"><c r="A3"><f>A2+25</f><v>stale</v></c></row>
  </sheetData>
</worksheet>`),
    'xl/worksheets/sheet2.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="B2:B2"/>
  <sheetData><row r="2"><c r="B2"/></row></sheetData>
</worksheet>`),
  })
}

function buildPublicExponentReptWorkbook(): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdWorkbook" Type="${officeRelationshipNamespace}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    'xl/sharedStrings.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
  <si><t>Q1</t></si>
  <si><t>bad</t></si>
</sst>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A2:C8"/>
  <sheetData>
    <row r="2"><c r="A2"><f>10^6</f><v>0</v></c></row>
    <row r="3"><c r="A3"><f>B3^C3</f><v>0</v></c><c r="B3"><v>2</v></c><c r="C3"><v>4</v></c></row>
    <row r="4"><c r="A4"><f>REPT(&quot;-&quot;,4)</f><v>stale</v></c></row>
    <row r="5"><c r="A5"><f>REPT(B5,C5)</f><v>stale</v></c><c r="B5" t="s"><v>0</v></c><c r="C5"><v>3</v></c></row>
    <row r="6"><c r="A6"><f>REPT(&quot;x&quot;,B6)</f><v>stale</v></c><c r="B6" t="s"><v>1</v></c></row>
    <row r="7"><c r="A7"><f>+&quot;Carbon dioxide equivalent emissions ($/tonne)&quot;&amp;REPT(&quot;.&quot;,3)</f><v>stale</v></c></row>
    <row r="8"><c r="A8"><f>B8*1000000/#REF!</f><v>stale</v></c><c r="B8"><v>2</v></c></row>
  </sheetData>
</worksheet>`),
  })
}

function buildPublicIndexMatchWorkbook(): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdWorkbook" Type="${officeRelationshipNamespace}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A2:B8"/>
  <sheetData>
    <row r="2"><c r="A2"><f>INDEX(B5:B8,MATCH(B6,B5:B8,0)+1,1)</f><v>0</v></c></row>
    <row r="5"><c r="B5"><v>10</v></c></row>
    <row r="6"><c r="B6"><v>20</v></c></row>
    <row r="7"><c r="B7"><v>30</v></c></row>
    <row r="8"><c r="B8"><v>40</v></c></row>
  </sheetData>
</worksheet>`),
  })
}

function buildPublicScalarDependencyWorkbook(): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets><sheet name="Budget" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdWorkbook" Type="${officeRelationshipNamespace}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="B6:C8"/>
  <sheetData>
    <row r="6"><c r="B6"><v>50</v></c></row>
    <row r="7"><c r="C7"><f>B6+B8*2</f><v>0</v></c></row>
    <row r="8"><c r="B8"><v>100</v></c></row>
  </sheetData>
</worksheet>`),
  })
}

function buildPublicCrossSheetScalarDependencyWorkbook(): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets>
    <sheet name="Data Sheet" sheetId="1" r:id="rId1"/>
    <sheet name="Budget" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`),
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdWorkbook" Type="${officeRelationshipNamespace}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="B2:C7"/>
  <sheetData>
    <row r="2"><c r="B2"><v>40</v></c></row>
    <row r="7"><c r="C7"><v>60</v></c></row>
  </sheetData>
</worksheet>`),
    'xl/worksheets/sheet2.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="C7"/>
  <sheetData>
    <row r="7"><c r="C7"><f>'Data Sheet'!B2+'Data Sheet'!C7*2</f><v>0</v></c></row>
  </sheetData>
</worksheet>`),
  })
}

function buildPublicVlookupWorkbook(): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets>
    <sheet name="Lookup Table" sheetId="1" r:id="rId1"/>
    <sheet name="Budget" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`),
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdWorkbook" Type="${officeRelationshipNamespace}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="B9:D10"/>
  <sheetData>
    <row r="9"><c r="B9"><v>100</v></c><c r="C9"><v>150</v></c><c r="D9"><v>175</v></c></row>
    <row r="10"><c r="B10"><v>200</v></c><c r="C10"><v>250</v></c><c r="D10"><v>275</v></c></row>
  </sheetData>
</worksheet>`),
    'xl/worksheets/sheet2.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="B2:D3"/>
  <sheetData>
    <row r="2"><c r="B2"><v>200</v></c></row>
    <row r="3"><c r="D3"><f>VLOOKUP(B2,'Lookup Table'!$B$9:$D$10,2,FALSE)</f><v>0</v></c></row>
  </sheetData>
</worksheet>`),
  })
}

function buildPublicApproximateVlookupWorkbook(): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets>
    <sheet name="Lookup Table" sheetId="1" r:id="rId1"/>
    <sheet name="Budget" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`),
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdWorkbook" Type="${officeRelationshipNamespace}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="B9:D11"/>
  <sheetData>
    <row r="9"><c r="B9"><v>100</v></c><c r="C9"><v>150</v></c><c r="D9"><v>175</v></c></row>
    <row r="10"><c r="B10"><v>200</v></c><c r="C10"><v>250</v></c><c r="D10"><v>275</v></c></row>
    <row r="11"><c r="B11"><v>300</v></c><c r="C11"><v>350</v></c><c r="D11"><v>375</v></c></row>
  </sheetData>
</worksheet>`),
    'xl/worksheets/sheet2.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="B2:D3"/>
  <sheetData>
    <row r="2"><c r="B2"><v>250</v></c></row>
    <row r="3"><c r="D3"><f>VLOOKUP(B2,'Lookup Table'!$B$9:$D$11,2)</f><v>0</v></c></row>
  </sheetData>
</worksheet>`),
  })
}

function buildPublicIndirectVlookupWorkbook(): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets>
    <sheet name="Lookup Table" sheetId="1" r:id="rId1"/>
    <sheet name="Budget" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`),
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdWorkbook" Type="${officeRelationshipNamespace}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    'xl/sharedStrings.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si><t>Lookup Table</t></si>
</sst>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="B9:D11"/>
  <sheetData>
    <row r="9"><c r="B9"><v>100</v></c><c r="C9"><v>150</v></c><c r="D9"><v>175</v></c></row>
    <row r="10"><c r="B10"><v>200</v></c><c r="C10"><v>250</v></c><c r="D10"><v>275</v></c></row>
    <row r="11"><c r="B11"><v>300</v></c><c r="C11"><v>350</v></c><c r="D11"><v>375</v></c></row>
  </sheetData>
</worksheet>`),
    'xl/worksheets/sheet2.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="B2:D4"/>
  <sheetData>
    <row r="2"><c r="B2"><v>250</v></c></row>
    <row r="3"><c r="D3"><f>VLOOKUP(B2,INDIRECT("'"&amp;C$4&amp;"'!"&amp;"$B$9:$D$11"),2)</f><v>0</v></c></row>
    <row r="4"><c r="C4" t="s"><v>0</v></c></row>
  </sheetData>
</worksheet>`),
  })
}

function buildExternalReferenceWorkbook(): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets><sheet name="Budget" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdWorkbook" Type="${officeRelationshipNamespace}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="B2"/>
  <sheetData>
    <row r="2"><c r="B2"><f>[1]QTR_detailed!$B130/10^6</f><v>0</v></c></row>
  </sheetData>
</worksheet>`),
  })
}

function buildCachedExternalReferenceWorkbook(): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets><sheet name="Budget" sheetId="1" r:id="rId1"/></sheets>
  <externalReferences><externalReference r:id="rId2"/></externalReferences>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="${officeRelationshipNamespace}/externalLink" Target="externalLinks/externalLink1.xml"/>
</Relationships>`),
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/externalLinks/externalLink1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>
</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdWorkbook" Type="${officeRelationshipNamespace}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    'xl/externalLinks/externalLink1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <externalBook r:id="rId1">
    <sheetNames>
      <sheetName val="Unused"/>
      <sheetName val="QTR_detailed"/>
      <sheetName val="BS &amp; Op Stat"/>
    </sheetNames>
    <sheetDataSet>
      <sheetData sheetId="0"/>
      <sheetData sheetId="1">
        <row r="130">
          <cell r="B130"><v>1460216076.6600001</v></cell>
          <cell r="R130"><v>99434495</v></cell>
          <cell r="S130"><v>0</v></cell>
          <cell r="T130"><v>381364.52</v></cell>
        </row>
      </sheetData>
      <sheetData sheetId="2">
        <row r="5"><cell r="D5"><v>2021</v></cell></row>
      </sheetData>
    </sheetDataSet>
  </externalBook>
</externalLink>`),
    'xl/externalLinks/_rels/externalLink1.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${officeRelationshipNamespace}/externalLinkPath" Target="file:///tmp/source.xlsx" TargetMode="External"/>
</Relationships>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="B2:E2"/>
  <sheetData>
    <row r="2">
      <c r="B2"><f>[1]QTR_detailed!$B130/10^6</f><v>0</v></c>
      <c r="C2"><f>SUM([1]QTR_detailed!$R130:$T130)/10^6</f><v>0</v></c>
      <c r="D2"><f>+'[1]BS &amp; Op Stat'!$D$5</f><v>0</v></c>
      <c r="E2"><f>SUM([1]QTR_detailed!$B130,[1]QTR_detailed!$T130,)/10^6</f><v>0</v></c>
    </row>
  </sheetData>
</worksheet>`),
  })
}

function buildExternalCompanionReferenceWorkbook(): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets>
    <sheet name="Unused" sheetId="1" r:id="rId1"/>
    <sheet name="QTR_detailed" sheetId="2" r:id="rId2"/>
    <sheet name="BS &amp; Op Stat" sheetId="3" r:id="rId3"/>
  </sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet3.xml"/>
</Relationships>`),
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdWorkbook" Type="${officeRelationshipNamespace}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1"/>
  <sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData>
</worksheet>`),
    'xl/worksheets/sheet2.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="B130:T130"/>
  <sheetData>
    <row r="130">
      <c r="B130"><v>2000000000</v></c>
      <c r="R130"><v>30000000</v></c>
      <c r="S130"><v>40000000</v></c>
      <c r="T130"><v>50000000</v></c>
    </row>
  </sheetData>
</worksheet>`),
    'xl/worksheets/sheet3.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="D5"/>
  <sheetData><row r="5"><c r="D5"><v>2024</v></c></row></sheetData>
</worksheet>`),
  })
}
