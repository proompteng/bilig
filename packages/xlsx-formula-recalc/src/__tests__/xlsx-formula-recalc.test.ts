import { existsSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate'
import { describe, expect, it } from 'vitest'

import { runXlsxFormulaRecalcCli } from '../cli-api.js'
import {
  WorkPaper,
  exportXlsx,
  importXlsx,
  inspectXlsxCache,
  parseQualifiedA1,
  recalculateSheetjsWorkbook,
  recalculateXlsx,
} from '@bilig/workpaper/xlsx'
import {
  inspectXlsxCacheFile,
  recalculateXlsx as recalculateNativeXlsx,
  recalculateXlsxFileToFile,
  xlsxFormulaRecalcBytesApiLimit,
} from '../index.js'

const officeRelationshipNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

describe('@bilig/xlsx-formula-recalc', () => {
  it('edits XLSX inputs, recalculates formulas, and exports a reimportable workbook', () => {
    const sourceWorkbook = WorkPaper.buildFromSheets({
      Inputs: [
        ['Metric', 'Value'],
        ['Units', 40],
        ['Price', 1200],
      ],
      Summary: [
        ['Metric', 'Value'],
        ['Revenue', '=Inputs!B2*Inputs!B3'],
      ],
    })
    const sourceBytes = exportXlsx(sourceWorkbook.exportSnapshot())
    sourceWorkbook.dispose()

    const result = recalculateXlsx(sourceBytes, {
      fileName: 'pricing.xlsx',
      edits: [
        { target: 'Inputs!B2', value: 48 },
        { target: 'Inputs!B3', value: 1500 },
      ],
      reads: ['Summary!B2'],
    })

    expect(readNumber(result.reads['Summary!B2'])).toBe(72_000)
    expect(result.warnings).toEqual([])
    expect(result.changes.length).toBeGreaterThan(0)

    const imported = importXlsx(result.xlsx, 'pricing.recalculated.xlsx')
    const restored = WorkPaper.buildFromSnapshot(imported.snapshot)
    const summary = restored.getSheetId('Summary')
    expect(summary).toBeTypeOf('number')
    expect(readNumber(restored.getCellValue({ sheet: summary!, row: 1, col: 1 }))).toBe(72_000)
    restored.dispose()
  })

  it('recalculates XLSX bytes through the primary native package API', async () => {
    const sourceWorkbook = WorkPaper.buildFromSheets({
      Inputs: [
        ['Metric', 'Value'],
        ['Units', 40],
        ['Price', 1200],
      ],
      Summary: [
        ['Metric', 'Value'],
        ['Revenue', '=Inputs!B2*Inputs!B3'],
      ],
    })
    const sourceBytes = exportXlsx(sourceWorkbook.exportSnapshot())
    sourceWorkbook.dispose()

    const result = await recalculateNativeXlsx(sourceBytes, {
      fileName: 'pricing.xlsx',
      edits: [
        { target: 'Inputs!B2', value: 48 },
        { target: 'Inputs!B3', value: 1500 },
      ],
      reads: ['Summary!B2'],
    })

    expect(readNumber(result.reads['Summary!B2'])).toBe(72_000)
    expect(result.diagnostics?.engineMode).toBe('streaming-native')
    expect(result.diagnostics?.fallbackUsed).toBe(false)
    expect(readCachedFormulaValue(result.xlsx, 'xl/worksheets/sheet2.xml', 'B2')).toBe('72000')
  })

  it('rejects large byte-buffer recalculation before temp-file materialization', async () => {
    await expect(
      recalculateNativeXlsx(new Uint8Array(xlsxFormulaRecalcBytesApiLimit + 1), {
        fileName: 'large-byte-api.xlsx',
      }),
    ).rejects.toThrow(/recalculateXlsx byte input is small-workbook only/u)
  })

  it('recalculates formula cells written without cached formula values', () => {
    const sourceWorkbook = WorkPaper.buildFromSheets({
      Inputs: [
        ['Metric', 'Value'],
        ['Units', 40],
        ['Price', 1200],
      ],
      Summary: [
        ['Metric', 'Value'],
        ['Revenue', '=Inputs!B2*Inputs!B3'],
      ],
    })
    const sourceBytes = replaceCellXml(
      exportXlsx(sourceWorkbook.exportSnapshot()),
      'xl/worksheets/sheet2.xml',
      'B2',
      '<c r="B2"><f>Inputs!B2*Inputs!B3</f></c>',
    )
    sourceWorkbook.dispose()

    const result = recalculateXlsx(sourceBytes, {
      fileName: 'pricing-without-formula-cache.xlsx',
      edits: [
        { target: 'Inputs!B2', value: 48 },
        { target: 'Inputs!B3', value: 1500 },
      ],
      reads: ['Summary!B2'],
    })

    expect(readNumber(result.reads['Summary!B2'])).toBe(72_000)
  })

  it('refreshes stale manual-calc caches for supported formulas before reads and export', () => {
    const sourceWorkbook = WorkPaper.buildFromSheets({
      Model: [
        ['Input', 'Output'],
        [2, '=A2*10'],
      ],
    })
    const sourceBytes = setWorkbookCalcPr(
      replaceCellXml(exportXlsx(sourceWorkbook.exportSnapshot()), 'xl/worksheets/sheet1.xml', 'B2', '<c r="B2"><f>A2*10</f><v>999</v></c>'),
      '<calcPr calcMode="manual" fullCalcOnLoad="0"/>',
    )
    sourceWorkbook.dispose()

    const result = recalculateXlsx(sourceBytes, {
      fileName: 'manual-stale-cache.xlsx',
      reads: ['Model!B2'],
    })

    expect(readNumber(result.reads['Model!B2'])).toBe(20)
    expect(readCachedFormulaValue(result.xlsx, 'xl/worksheets/sheet1.xml', 'B2')).toBe('20')
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'calcMode')).toBe('manual')
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'fullCalcOnLoad')).toBeNull()

    const imported = importXlsx(result.xlsx, 'manual-stale-cache.recalculated.xlsx')
    const restored = WorkPaper.buildFromSnapshot(imported.snapshot)
    const model = restored.getSheetId('Model')
    expect(model).toBeTypeOf('number')
    expect(readNumber(restored.getCellValue({ sheet: model!, row: 1, col: 1 }))).toBe(20)
    restored.dispose()
  })

  it('recalculates manual-calc lazy imported worksheets without cloning lazy cell proxies', () => {
    const result = recalculateXlsx(buildLargeManualCalcWorkbookBytes(), {
      fileName: 'large-manual-lazy-import.xlsx',
      reads: ['Data!B1'],
    })

    expect(readNumber(result.reads['Data!B1'])).toBe(4)
    expect(readCachedFormulaValue(result.xlsx, 'xl/worksheets/sheet1.xml', 'B1')).toBe('4')
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'calcMode')).toBe('manual')
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'fullCalcOnLoad')).toBeNull()
  })

  it('recalculates supported row-local table formulas through the streaming-native file path', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-table-recalc-'))
    try {
      const sourcePath = join(tempDir, 'ocha-like.xlsx')
      const outputPath = join(tempDir, 'ocha-like.recalculated.xlsx')
      writeFileSync(sourcePath, buildNativeTableFormulaWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        edits: [{ target: 'Data!R57152', value: 16 }],
        reads: ['Data!U57152', 'Data!V57152', 'Data!AD57152', 'Data!AF57152', 'Data!AG57152', 'Data!AI57152'],
      })

      expect(readNumber(result.reads['Data!U57152'])).toBe(168.75)
      expect(readNumber(result.reads['Data!V57152'])).toBe(28.125)
      expect(readNumber(result.reads['Data!AD57152'])).toBe(3)
      expect(readString(result.reads['Data!AF57152'])).toBe('Cash')
      expect(readString(result.reads['Data!AG57152'])).toBe('Child Protection')
      expect(readString(result.reads['Data!AI57152'])).toBe('Mar')
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.fallbackUsed).toBe(false)
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(6)
      expect(result.diagnostics?.formulaCounts.nativeKernelFormulaCellCount).toBe(6)
      expect(result.diagnostics?.formulaCounts.nativeKernelBatchCount).toBe(2)
      const outputBytes = readFileSync(outputPath)
      const sheetXml = strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain('<f>DataTable[[#This Row],[V Households]]*6</f><v>168.75</v>')
      expect(sheetXml).toContain('<f>DataTable[[#This Row],[T USD]]/DataTable[[#This Row],[R Input]]</f><v>28.125</v>')
      expect(sheetXml).toContain(
        '<f>_xlfn.IFS(A57152=&quot;January&quot;,1,A57152=&quot;February&quot;,2,A57152=&quot;March&quot;,3)</f><v>3</v>',
      )
      expect(sheetXml).toContain(
        '<f>IF(DataTable[[#This Row],[Assistance modality]]=&quot;In kind&quot;,&quot;In kind&quot;,&quot;Cash&quot;)</f><v>Cash</v>',
      )
      expect(sheetXml).toContain('<f>DataTable[[#This Row],[Cluster/AoR]]</f><v>Child Protection</v>')
      expect(sheetXml).toContain(
        '<f>_xlfn.IFS(A57152=&quot;January&quot;,&quot;Jan&quot;,A57152=&quot;February&quot;,&quot;Feb&quot;,A57152=&quot;March&quot;,&quot;Mar&quot;)</f><v>Mar</v>',
      )
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('refuses WorkPaper fallback from the primary file-to-file API', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-primary-no-workpaper-'))
    try {
      const sourcePath = join(tempDir, 'source.xlsx')
      const engineOutputPath = join(tempDir, 'engine.recalculated.xlsx')
      const fallbackOutputPath = join(tempDir, 'fallback.recalculated.xlsx')
      writeFileSync(sourcePath, buildNativeTableFormulaWorkbook())

      await expect(
        // @ts-expect-error Exercising the runtime guard for legacy JS callers.
        recalculateXlsxFileToFile(sourcePath, {
          outputPath: engineOutputPath,
          engine: 'workpaper',
          reads: ['Data!U57152'],
        }),
      ).rejects.toThrow(/@bilig\/workpaper/u)
      expect(existsSync(engineOutputPath)).toBe(false)

      await expect(
        // @ts-expect-error Exercising the runtime guard for legacy JS callers.
        recalculateXlsxFileToFile(sourcePath, {
          outputPath: fallbackOutputPath,
          fallbackPolicy: 'workpaper',
          reads: ['Data!U57152'],
        }),
      ).rejects.toThrow(/@bilig\/workpaper/u)
      expect(existsSync(fallbackOutputPath)).toBe(false)
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('recalculates xlsx-fixture-corpus style ratio row chains through the native kernel', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-ratio-chain-'))
    try {
      const sourcePath = join(tempDir, 'public-ratio-chain.xlsx')
      const outputPath = join(tempDir, 'public-ratio-chain.recalculated.xlsx')
      writeFileSync(sourcePath, buildPublicRatioChainWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        reads: ["'Tableau 12'!G5", "'Tableau 12'!H5"],
      })

      expect(readNumber(result.reads["'Tableau 12'!G5"])).toBe(1_200)
      expect(readNumber(result.reads["'Tableau 12'!H5"])).toBe(300)
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(2)
      expect(result.diagnostics?.formulaCounts.nativeKernelFormulaCellCount).toBe(2)
      expect(result.diagnostics?.formulaCounts.nativeKernelBatchCount).toBe(1)
      const outputBytes = readFileSync(outputPath)
      const sheetXml = strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain('<c r="G5"><f>F5+E5</f><v>1200</v></c>')
      expect(sheetXml).toContain('<c r="H5"><f>G5/C5</f><v>300</v></c>')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('recalculates xlsx-fixture-corpus direct scalar formulas through the native kernel', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-direct-scalar-'))
    try {
      const sourcePath = join(tempDir, 'public-direct-scalar.xlsx')
      const outputPath = join(tempDir, 'public-direct-scalar.recalculated.xlsx')
      writeFileSync(sourcePath, buildPublicDirectScalarWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        reads: ["'Tableau 12'!D5"],
      })

      expect(readNumber(result.reads["'Tableau 12'!D5"])).toBe(0.8603058823529411)
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(1)
      expect(result.diagnostics?.formulaCounts.nativeKernelFormulaCellCount).toBe(1)
      expect(result.diagnostics?.formulaCounts.nativeKernelBatchCount).toBe(1)
      const outputBytes = readFileSync(outputPath)
      const sheetXml = strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain('<c r="D5"><f>C5/B5</f><v>0.8603058823529411</v></c>')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('recalculates xlsx-fixture-corpus inline ratio row chains through the native kernel', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-inline-ratio-chain-'))
    try {
      const sourcePath = join(tempDir, 'public-inline-ratio-chain.xlsx')
      const outputPath = join(tempDir, 'public-inline-ratio-chain.recalculated.xlsx')
      writeFileSync(sourcePath, buildPublicInlineRatioChainWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        reads: ["'Table 8'!G5", "'Table 8'!H5"],
      })

      expect(readNumber(result.reads["'Table 8'!G5"])).toBe(1_274_633_000)
      expect(readNumber(result.reads["'Table 8'!H5"])).toBe(1.3314110230792131)
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      expect(result.diagnostics?.formulaCounts.evaluatedFormulaCellCount).toBe(2)
      expect(result.diagnostics?.formulaCounts.nativeKernelFormulaCellCount).toBe(2)
      expect(result.diagnostics?.formulaCounts.nativeKernelBatchCount).toBe(1)
      const outputBytes = readFileSync(outputPath)
      const sheetXml = strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
      expect(sheetXml).toContain('<c r="G5"><f>F5+E5</f><v>1274633000</v></c>')
      expect(sheetXml).toContain('<c r="H5"><f>(F5+E5)/C5</f><v>1.331411023079213</v></c>')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('translates shared formulas before streaming-native row-local evaluation', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-native-shared-recalc-'))
    try {
      const sourcePath = join(tempDir, 'shared.xlsx')
      const outputPath = join(tempDir, 'shared.recalculated.xlsx')
      writeFileSync(sourcePath, buildNativeSharedFormulaWorkbook())

      const result = await recalculateXlsxFileToFile(sourcePath, {
        outputPath,
        engine: 'streaming-native',
        edits: [{ target: 'Data!A3', value: 7 }],
        reads: ['Data!B3'],
      })

      expect(readNumber(result.reads['Data!B3'])).toBe(70)
      expect(result.diagnostics?.engineMode).toBe('streaming-native')
      const outputBytes = readFileSync(outputPath)
      expect(readCachedFormulaValue(outputBytes, 'xl/worksheets/sheet1.xml', 'B3')).toBe('70')
      expect(strFromU8(unzipSync(outputBytes)['xl/worksheets/sheet1.xml'] ?? new Uint8Array())).toContain('<f t="shared" si="0"/><v>70</v>')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('inspects stale XLSX formula caches through the public API', () => {
    const sourceWorkbook = WorkPaper.buildFromSheets({
      Sheet1: [
        ['Input', 'Output'],
        [2, '=A2*10'],
      ],
    })
    const sourceBytes = replaceCellXml(
      exportXlsx(sourceWorkbook.exportSnapshot()),
      'xl/worksheets/sheet1.xml',
      'B2',
      '<c r="B2"><f>A2*10</f><v>999</v></c>',
    )
    sourceWorkbook.dispose()

    const report = inspectXlsxCache(sourceBytes, { fileName: 'stale-cache.xlsx' })

    expect(report.schemaVersion).toBe('xlsx-cache-doctor.v1')
    expect(report.formulaCellCount).toBe(1)
    expect(report.inspectedFormulaCellCount).toBe(1)
    expect(report.uninspectedFormulaCellCount).toBe(0)
    expect(report.inspectionLimit).toBe('all')
    expect(report.staleCachedFormulaCount).toBe(1)
    expect(report.cacheStatusSummary).toEqual({
      inspected: 1,
      stale: 1,
      fresh: 0,
      missingCache: 0,
      unsupportedRecalculation: 0,
    })
    expect(report.suggestedReads).toEqual(['Sheet1!B2'])
    expect(report.formulas[0]).toMatchObject({
      target: 'Sheet1!B2',
      formula: '=A2*10',
      cachedValue: 999,
      literalRecalculatedValue: 20,
      cacheStatus: 'stale',
      staleCachedValue: true,
    })
    expect(report.warnings).toEqual([])
    expect(report.inspectionCompleted).toBe(true)
    expect(report.recalculationCompleted).toBe(true)
    expect(report.excelParity).toBe('not_proven')
  })

  it('inspects stale XLSX formula caches through the public native file API', async () => {
    const tempDir = mkdtempSync(join(tmpdir(), 'xlsx-formula-recalc-file-cache-api-'))
    try {
      const inputPath = join(tempDir, 'stale-cache.xlsx')
      const sourceWorkbook = WorkPaper.buildFromSheets({
        Sheet1: [
          ['Input', 'Output'],
          [2, '=A2*10'],
        ],
      })
      const sourceBytes = replaceCellXml(
        exportXlsx(sourceWorkbook.exportSnapshot()),
        'xl/worksheets/sheet1.xml',
        'B2',
        '<c r="B2"><f>A2*10</f><v>999</v></c>',
      )
      sourceWorkbook.dispose()
      writeFileSync(inputPath, sourceBytes)

      const report = await inspectXlsxCacheFile(inputPath, {
        inspectLimit: 'all',
      })

      expect(report.schemaVersion).toBe('xlsx-cache-doctor.v1')
      expect(report.formulaCellCount).toBe(1)
      expect(report.inspectedFormulaCellCount).toBe(1)
      expect(report.uninspectedFormulaCellCount).toBe(0)
      expect(report.inspectionLimit).toBe('all')
      expect(report.staleCachedFormulaCount).toBe(1)
      expect(report.cacheStatusSummary).toEqual({
        inspected: 1,
        stale: 1,
        fresh: 0,
        missingCache: 0,
        unsupportedRecalculation: 0,
      })
      expect(report.suggestedReads).toEqual(['Sheet1!B2'])
      expect(report.formulas[0]).toMatchObject({
        target: 'Sheet1!B2',
        formula: '=A2*10',
        cachedValue: 999,
        literalRecalculatedValue: 20,
        cacheStatus: 'stale',
        staleCachedValue: true,
      })
      expect(report.diagnostics.engineMode).toBe('streaming-native')
      expect(report.diagnostics.fallbackUsed).toBe(false)
      expect(report.recalculationCompleted).toBe(true)
      expect(report.excelParity).toBe('not_proven')
    } finally {
      rmSync(tempDir, { recursive: true, force: true })
    }
  })

  it('clears stale calculation metadata after explicit formula recalculation', () => {
    const sourceWorkbook = WorkPaper.buildFromSheets({
      Model: [
        ['Input', 'Output'],
        [2, '=A2*10'],
      ],
    })
    const sourceBytes = setWorkbookCalcPr(
      replaceCellXml(exportXlsx(sourceWorkbook.exportSnapshot()), 'xl/worksheets/sheet1.xml', 'B2', '<c r="B2"><f>A2*10</f><v>999</v></c>'),
      '<calcPr calcMode="manual" calcOnSave="1" calcCompleted="0" fullCalcOnLoad="0" forceFullCalc="1"/>',
    )
    sourceWorkbook.dispose()

    const result = recalculateXlsx(sourceBytes, {
      fileName: 'manual-stale-cache-flags.xlsx',
      reads: ['Model!B2'],
    })

    expect(readNumber(result.reads['Model!B2'])).toBe(20)
    expect(readCachedFormulaValue(result.xlsx, 'xl/worksheets/sheet1.xml', 'B2')).toBe('20')
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'calcMode')).toBe('manual')
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'calcOnSave')).toBeNull()
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'calcCompleted')).toBeNull()
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'fullCalcOnLoad')).toBeNull()
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'forceFullCalc')).toBeNull()
  })

  it('drops incomplete automatic calc metadata after explicit formula recalculation', () => {
    const sourceWorkbook = WorkPaper.buildFromSheets({
      Model: [
        ['Input', 'Output'],
        [2, '=A2*10'],
      ],
    })
    const sourceBytes = setWorkbookCalcPr(
      replaceCellXml(exportXlsx(sourceWorkbook.exportSnapshot()), 'xl/worksheets/sheet1.xml', 'B2', '<c r="B2"><f>A2*10</f><v>999</v></c>'),
      '<calcPr calcCompleted="0"/>',
    )
    sourceWorkbook.dispose()

    const result = recalculateXlsx(sourceBytes, {
      fileName: 'automatic-incomplete-cache.xlsx',
      reads: ['Model!B2'],
    })

    expect(readNumber(result.reads['Model!B2'])).toBe(20)
    expect(readCachedFormulaValue(result.xlsx, 'xl/worksheets/sheet1.xml', 'B2')).toBe('20')
    expect(readWorkbookXml(result.xlsx)).not.toContain('<calcPr')
  })

  it('carries import warnings when unsupported cached formulas recalculate to Excel errors', () => {
    const sourceWorkbook = WorkPaper.buildFromSheets({
      Model: [['AAPL', 0]],
    })
    const sourceBytes = replaceCellXml(
      exportXlsx(sourceWorkbook.exportSnapshot()),
      'xl/worksheets/sheet1.xml',
      'B1',
      '<c r="B1"><f>_xldudf_WISEPRICE(A1,&quot;Shares Outstanding&quot;)</f><v>14935800000</v></c>',
    )
    sourceWorkbook.dispose()

    const result = recalculateXlsx(sourceBytes, {
      fileName: 'unsupported-cache-warning.xlsx',
      reads: ['Model!B1'],
    })

    expect(result.warnings).toContain(
      'Unsupported formulas were preserved from cached XLSX values; recalculation may return Excel error values.',
    )
    expect(readCachedFormulaValue(result.xlsx, 'xl/worksheets/sheet1.xml', 'B1')).toBe('#NAME?')
  })

  it('hydrates external-link caches from companion workbook bytes before recalculation', () => {
    const sourceBytes = buildExternalLinkRangeCacheWorkbook('file:///tmp/rates.xlsx')
    const result = recalculateXlsx(sourceBytes, {
      fileName: 'external-link-cache.xlsx',
      externalWorkbooks: [{ fileName: 'rates.xlsx', bytes: buildExternalSourceWorkbook([20, 30, 40]) }],
      reads: ['Model!C1', 'Model!C2', 'Model!C3'],
    })

    expect(readNumber(result.reads['Model!C1'])).toBe(180)
    expect(readNumber(result.reads['Model!C2'])).toBe(60)
    expect(readNumber(result.reads['Model!C3'])).toBe(80)
    expect(readCachedFormulaValue(result.xlsx, 'xl/worksheets/sheet1.xml', 'C1')).toBe('180')
    expect(readCachedFormulaValue(result.xlsx, 'xl/worksheets/sheet1.xml', 'C2')).toBe('60')
    expect(readCachedFormulaValue(result.xlsx, 'xl/worksheets/sheet1.xml', 'C3')).toBe('80')
    expect(result.diagnostics?.externalWorkbookHydration).toMatchObject({
      externalWorkbookCount: 1,
      externalReferenceCount: 1,
      refreshedBookIndices: [1],
      refreshedSheetCount: 1,
      refreshedCellCount: 6,
      skippedNoMatchCount: 0,
      skippedAmbiguousMatchCount: 0,
      references: [
        expect.objectContaining({
          bookIndex: 1,
          status: 'refreshed',
          candidateCount: 1,
          referenceCandidateCount: 1,
          matchKind: 'unique-workbook-identity',
          matchedFileName: 'rates.xlsx',
          refreshedCellCount: 6,
        }),
      ],
    })
    expect(readExternalLinkCacheCellValue(result.xlsx, 'B2')).toBe('20')
    expect(readExternalLinkCacheCellValue(result.xlsx, 'B3')).toBe('30')
    expect(readExternalLinkCacheCellValue(result.xlsx, 'B4')).toBe('40')
    const externalLinkCacheXml = readExternalLinkCacheXml(result.xlsx)
    expect(externalLinkCacheXml).not.toContain('<row r="1">')
    expect(externalLinkCacheXml).toContain('<row r="2">')
    expect(externalLinkCacheXml).toContain('<row r="3">')
    expect(externalLinkCacheXml).toContain('<row r="4">')
    expect(externalLinkCacheXml).not.toContain('<row r="0">')
  })

  it('hydrates sparse blank and error external-link ranges before recalculation', () => {
    const sourceBytes = buildSparseExternalLinkRangeCacheWorkbook('file:///tmp/rates.xlsx')
    const result = recalculateXlsx(sourceBytes, {
      fileName: 'external-link-sparse-cache.xlsx',
      externalWorkbooks: [{ fileName: 'rates.xlsx', bytes: buildSparseExternalSourceWorkbook() }],
      reads: ['Model!C1', 'Model!C2'],
    })

    expect(readNumber(result.reads['Model!C1'])).toBe(70)
    expect(readNumber(result.reads['Model!C2'])).toBe(99)
    expect(readCachedFormulaValue(result.xlsx, 'xl/worksheets/sheet1.xml', 'C1')).toBe('70')
    expect(readCachedFormulaValue(result.xlsx, 'xl/worksheets/sheet1.xml', 'C2')).toBe('99')
    expect(result.diagnostics?.externalWorkbookHydration).toMatchObject({
      externalWorkbookCount: 1,
      externalReferenceCount: 1,
      refreshedBookIndices: [1],
      refreshedSheetCount: 1,
      refreshedCellCount: 4,
      skippedNoMatchCount: 0,
      skippedAmbiguousMatchCount: 0,
      skippedEmptyRefreshCount: 0,
    })
    expect(readExternalLinkCacheCellValue(result.xlsx, 'B2')).toBe('20')
    expect(readExternalLinkCacheCellValue(result.xlsx, 'B3')).toBeNull()
    expect(readExternalLinkCacheCellValue(result.xlsx, 'B4')).toBe('50')
    expect(readExternalLinkCacheCellValue(result.xlsx, 'B5')).toBe('#N/A')
    expect(readExternalLinkCacheXml(result.xlsx)).toContain('<cell r="B5" t="e"><v>#N/A</v></cell>')
  })

  it('preserves cached external-link values when companion workbook matching is ambiguous', () => {
    const sourceBytes = buildExternalLinkRangeCacheWorkbook('file:///tmp/rates.xlsx')
    const result = recalculateXlsx(sourceBytes, {
      fileName: 'external-link-cache.xlsx',
      externalWorkbooks: [
        { fileName: 'rates.xlsx', bytes: buildExternalSourceWorkbook([20, 30, 40]) },
        { fileName: 'rates.xlsx', bytes: buildExternalSourceWorkbook([200, 300, 400]) },
      ],
      reads: ['Model!C1', 'Model!C2', 'Model!C3'],
    })

    expect(readNumber(result.reads['Model!C1'])).toBe(120)
    expect(readNumber(result.reads['Model!C2'])).toBe(40)
    expect(readNumber(result.reads['Model!C3'])).toBe(60)
    expect(result.warnings).toContain(
      'Some supplied external workbook companions matched ambiguously; existing external-link cache values were preserved.',
    )
    expect(result.diagnostics?.externalWorkbookHydration).toMatchObject({
      externalWorkbookCount: 2,
      externalReferenceCount: 1,
      refreshedBookIndices: [],
      skippedNoMatchCount: 0,
      skippedAmbiguousMatchCount: 1,
      references: [
        expect.objectContaining({
          bookIndex: 1,
          status: 'skipped-ambiguous-match',
          candidateCount: 2,
          referenceCandidateCount: 1,
          matchKind: 'unique-workbook-identity',
        }),
      ],
    })
    expect(readExternalLinkCacheCellValue(result.xlsx, 'B2')).toBe('10')
    expect(readExternalLinkCacheCellValue(result.xlsx, 'B3')).toBe('20')
    expect(readExternalLinkCacheCellValue(result.xlsx, 'B4')).toBe('30')
  })

  it('preserves non-stale calculation preferences after explicit formula recalculation', () => {
    const sourceWorkbook = WorkPaper.buildFromSheets({
      Model: [
        ['Input', 'Output'],
        [2, '=A2*10'],
      ],
    })
    const sourceBytes = setWorkbookCalcPr(
      replaceCellXml(exportXlsx(sourceWorkbook.exportSnapshot()), 'xl/worksheets/sheet1.xml', 'B2', '<c r="B2"><f>A2*10</f><v>999</v></c>'),
      '<calcPr calcMode="manual" fullPrecision="0" iterate="1" iterateCount="32" iterateDelta="0.001" concurrentCalc="0" calcCompleted="0" fullCalcOnLoad="0"/>',
    )
    sourceWorkbook.dispose()

    const result = recalculateXlsx(sourceBytes, {
      fileName: 'manual-stale-cache-with-preferences.xlsx',
      reads: ['Model!B2'],
    })

    expect(readNumber(result.reads['Model!B2'])).toBe(20)
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'calcMode')).toBe('manual')
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'fullPrecision')).toBe('0')
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'iterate')).toBe('1')
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'iterateCount')).toBe('32')
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'iterateDelta')).toBe('0.001')
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'concurrentCalc')).toBe('0')
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'calcCompleted')).toBeNull()
    expect(readWorkbookCalcPrAttribute(result.xlsx, 'fullCalcOnLoad')).toBeNull()
  })

  it('parses quoted sheet names and absolute A1 addresses', () => {
    expect(parseQualifiedA1("'Pricing Model'!$AB$12")).toEqual({
      sheetName: 'Pricing Model',
      row: 11,
      col: 27,
    })
  })

  it('exposes a SheetJS-named API and CLI alias from the live xlsx package', () => {
    const sourceWorkbook = WorkPaper.buildFromSheets({
      Inputs: [
        ['Metric', 'Value'],
        ['Units', 40],
        ['Price', 1200],
      ],
      Summary: [
        ['Metric', 'Value'],
        ['Revenue', '=Inputs!B2*Inputs!B3'],
      ],
    })
    const sourceBytes = exportXlsx(sourceWorkbook.exportSnapshot())
    sourceWorkbook.dispose()

    const result = recalculateSheetjsWorkbook(sourceBytes, {
      edits: [{ target: 'Inputs!B2', value: 48 }],
      reads: ['Summary!B2'],
    })
    expect(readNumber(result.reads['Summary!B2'])).toBe(57_600)

    let help = ''
    const exitCode = runXlsxFormulaRecalcCli(['--help'], {
      commandName: 'sheetjs-recalc',
      stdout: (text) => {
        help += text
      },
    })
    expect(exitCode).toBe(0)
    expect(help).toContain('Usage: sheetjs-recalc')
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

function buildLargeManualCalcWorkbookBytes(): Uint8Array {
  const rowCount = 65_537
  const rows = ['<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1*2</f><v>999</v></c></row>']
  for (let row = 2; row <= rowCount; row += 1) {
    rows.push(`<row r="${String(row)}"><c r="A${String(row)}"><v>${String(row)}</v></c></row>`)
  }
  return buildIndependentWorkbook(
    [
      {
        name: 'Data',
        path: 'xl/worksheets/sheet1.xml',
        xml: [
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
          `<dimension ref="A1:B${String(rowCount)}"/>`,
          `<sheetData>${rows.join('')}</sheetData>`,
          '</worksheet>',
        ].join(''),
      },
    ],
    '<calcPr calcMode="manual" fullCalcOnLoad="0"/>',
  )
}

function buildNativeTableFormulaWorkbook(): Uint8Array {
  const headers = Array.from({ length: 35 }, (_value, index) => `Column ${String(index + 1)}`)
  headers[0] = 'Month'
  headers[17] = 'R Input'
  headers[19] = 'T USD'
  headers[20] = 'U Individuals'
  headers[21] = 'V Households'
  headers[28] = 'Assistance modality'
  headers[29] = 'AD Month index'
  headers[30] = 'Cluster/AoR'
  headers[31] = 'AF Modality group'
  headers[32] = 'AG Cluster copy'
  headers[34] = 'AI Month label'
  const headerCells = headers.map((header, index) => `<c r="${columnName(index)}1" t="inlineStr"><is><t>${header}</t></is></c>`).join('')
  const dataCells = [
    '<c r="A57152" t="inlineStr"><is><t>March</t></is></c>',
    '<c r="R57152"><v>30</v></c>',
    '<c r="T57152"><v>450</v></c>',
    '<c r="U57152"><f>DataTable[[#This Row],[V Households]]*6</f><v>90</v></c>',
    '<c r="V57152"><f>DataTable[[#This Row],[T USD]]/DataTable[[#This Row],[R Input]]</f><v>15</v></c>',
    '<c r="AC57152" t="inlineStr"><is><t>Cash and voucher assistance</t></is></c>',
    '<c r="AD57152"><f>_xlfn.IFS(A57152=&quot;January&quot;,1,A57152=&quot;February&quot;,2,A57152=&quot;March&quot;,3)</f><v>1</v></c>',
    '<c r="AE57152" t="inlineStr"><is><t>Child Protection</t></is></c>',
    '<c r="AF57152" t="str"><f>IF(DataTable[[#This Row],[Assistance modality]]=&quot;In kind&quot;,&quot;In kind&quot;,&quot;Cash&quot;)</f><v>In kind</v></c>',
    '<c r="AG57152" t="str"><f>DataTable[[#This Row],[Cluster/AoR]]</f><v>Old Cluster</v></c>',
    '<c r="AI57152" t="str"><f>_xlfn.IFS(A57152=&quot;January&quot;,&quot;Jan&quot;,A57152=&quot;February&quot;,&quot;Feb&quot;,A57152=&quot;March&quot;,&quot;Mar&quot;)</f><v>Jan</v></c>',
  ].join('')
  return zipSync({
    '[Content_Types].xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>
</Types>`),
    '_rels/.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`),
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${officeRelationshipNamespace}/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`),
    'xl/worksheets/sheet1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <dimension ref="A1:AI57152"/>
  <sheetData><row r="1">${headerCells}</row><row r="57152">${dataCells}</row></sheetData>
  <tableParts count="1"><tablePart r:id="rIdTable1"/></tableParts>
</worksheet>`),
    'xl/worksheets/_rels/sheet1.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdTable1" Type="${officeRelationshipNamespace}/table" Target="../tables/table1.xml"/>
</Relationships>`),
    'xl/tables/table1.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="DataTable" displayName="DataTable" ref="A1:AI57152" headerRowCount="1">
  <tableColumns count="35">${headers
    .map((header, index) => `<tableColumn id="${String(index + 1)}" name="${header}"/>`)
    .join('')}</tableColumns>
</table>`),
  })
}

function buildPublicRatioChainWorkbook(): Uint8Array {
  return buildIndependentWorkbook([
    {
      name: 'Tableau 12',
      path: 'xl/worksheets/sheet1.xml',
      xml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:H5"/>',
        '<sheetData>',
        '<row r="1"><c r="C1" t="inlineStr"><is><t>Investors</t></is></c><c r="E1" t="inlineStr"><is><t>Other</t></is></c><c r="F1" t="inlineStr"><is><t>Primary</t></is></c><c r="G1" t="inlineStr"><is><t>Total</t></is></c><c r="H1" t="inlineStr"><is><t>Ratio</t></is></c></row>',
        '<row r="5"><c r="C5"><v>4</v></c><c r="E5"><v>300</v></c><c r="F5"><v>900</v></c><c r="G5"><f>F5+E5</f><v>0</v></c><c r="H5"><f>G5/C5</f><v>0</v></c></row>',
        '</sheetData>',
        '</worksheet>',
      ].join(''),
    },
  ])
}

function buildPublicDirectScalarWorkbook(): Uint8Array {
  return buildIndependentWorkbook([
    {
      name: 'Tableau 12',
      path: 'xl/worksheets/sheet1.xml',
      xml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:D5"/>',
        '<sheetData>',
        '<row r="1"><c r="B1" t="inlineStr"><is><t>Capital</t></is></c><c r="C1" t="inlineStr"><is><t>Value</t></is></c><c r="D1" t="inlineStr"><is><t>Ratio</t></is></c></row>',
        '<row r="5"><c r="B5"><v>42500000</v></c><c r="C5"><v>36563000</v></c><c r="D5"><f>C5/B5</f><v>0</v></c></row>',
        '</sheetData>',
        '</worksheet>',
      ].join(''),
    },
  ])
}

function buildPublicInlineRatioChainWorkbook(): Uint8Array {
  return buildIndependentWorkbook([
    {
      name: 'Table 8',
      path: 'xl/worksheets/sheet1.xml',
      xml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:H5"/>',
        '<sheetData>',
        '<row r="1"><c r="C1" t="inlineStr"><is><t>Capital Invested</t></is></c><c r="E1" t="inlineStr"><is><t>Realized</t></is></c><c r="F1" t="inlineStr"><is><t>Unrealized</t></is></c><c r="G1" t="inlineStr"><is><t>Total Value</t></is></c><c r="H1" t="inlineStr"><is><t>TVPI</t></is></c></row>',
        '<row r="5"><c r="C5"><v>957355000</v></c><c r="E5"><v>176091000</v></c><c r="F5"><v>1098542000</v></c><c r="G5"><f>F5+E5</f><v>0</v></c><c r="H5"><f>(F5+E5)/C5</f><v>0</v></c></row>',
        '</sheetData>',
        '</worksheet>',
      ].join(''),
    },
  ])
}

function buildNativeSharedFormulaWorkbook(): Uint8Array {
  return buildIndependentWorkbook([
    {
      name: 'Data',
      path: 'xl/worksheets/sheet1.xml',
      xml: [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        '<dimension ref="A1:B3"/>',
        '<sheetData>',
        '<row r="1"><c r="A1" t="inlineStr"><is><t>Input</t></is></c><c r="B1" t="inlineStr"><is><t>Output</t></is></c></row>',
        '<row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0" ref="B2:B3">A2*10</f><v>20</v></c></row>',
        '<row r="3"><c r="A3"><v>3</v></c><c r="B3"><f t="shared" si="0"/><v>999</v></c></row>',
        '</sheetData>',
        '</worksheet>',
      ].join(''),
    },
  ])
}

function columnName(index: number): string {
  let value = index + 1
  let output = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    output = String.fromCharCode(65 + remainder) + output
    value = Math.floor((value - 1) / 26)
  }
  return output
}

function buildIndependentWorkbook(
  sheets: readonly { readonly name: string; readonly path: string; readonly xml: string }[],
  calcPrXml = '',
): Uint8Array {
  return zipSync({
    'xl/workbook.xml': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">
  <sheets>${sheets
    .map((sheet, index) => `<sheet name="${sheet.name}" sheetId="${String(index + 1)}" r:id="rId${String(index + 1)}"/>`)
    .join('')}</sheets>${calcPrXml}
</workbook>`),
    'xl/_rels/workbook.xml.rels': strToU8(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${sheets
  .map(
    (sheet, index) =>
      `<Relationship Id="rId${String(index + 1)}" Type="${officeRelationshipNamespace}/worksheet" Target="${sheet.path.slice('xl/'.length)}"/>`,
  )
  .join('')}
</Relationships>`),
    ...Object.fromEntries(sheets.map((sheet) => [sheet.path, strToU8(sheet.xml)])),
  })
}

function replaceCellXml(bytes: Uint8Array, sheetPath: string, address: string, replacement: string): Uint8Array {
  const zip = unzipSync(bytes)
  const sheetXml = strFromU8(zip[sheetPath] ?? new Uint8Array())
  const pattern = new RegExp(`<c\\b(?=[^>]*\\br="${address}")[\\s\\S]*?<\\/c>`, 'u')
  if (!pattern.test(sheetXml)) {
    throw new Error(`Missing cell XML for ${address}`)
  }
  zip[sheetPath] = strToU8(sheetXml.replace(pattern, replacement))
  return zipSync(zip)
}

function setWorkbookCalcPr(bytes: Uint8Array, calcPrXml: string): Uint8Array {
  const zip = unzipSync(bytes)
  const workbookXml = strFromU8(zip['xl/workbook.xml'] ?? new Uint8Array())
  if (/<calcPr\b[\s\S]*?\/>/u.test(workbookXml)) {
    zip['xl/workbook.xml'] = strToU8(workbookXml.replace(/<calcPr\b[\s\S]*?\/>/u, calcPrXml))
    return zipSync(zip)
  }
  zip['xl/workbook.xml'] = strToU8(workbookXml.replace('</workbook>', `${calcPrXml}</workbook>`))
  return zipSync(zip)
}

function readWorkbookXml(bytes: Uint8Array): string {
  return strFromU8(unzipSync(bytes)['xl/workbook.xml'] ?? new Uint8Array())
}

function readWorkbookCalcPrAttribute(bytes: Uint8Array, name: string): string | null {
  const calcPr = /<calcPr\b([^>]*)\/?>/u.exec(readWorkbookXml(bytes))?.[1] ?? ''
  const match = new RegExp(`\\b${name}="([^"]*)"`, 'u').exec(calcPr)
  return match?.[1] ?? null
}

function readCachedFormulaValue(bytes: Uint8Array, sheetPath: string, address: string): string | null {
  const zip = unzipSync(bytes)
  const sheetXml = strFromU8(zip[sheetPath] ?? new Uint8Array())
  const match = new RegExp(`<c\\b(?=[^>]*\\br="${address}")[\\s\\S]*?<v>([\\s\\S]*?)<\\/v>[\\s\\S]*?<\\/c>`, 'u').exec(sheetXml)
  return match?.[1] ?? null
}

function buildExternalSourceWorkbook(rates: readonly [number, number, number]): Uint8Array {
  const workbook = WorkPaper.buildFromSheets({
    Rates: [
      ['SKU', 'Rate'],
      ['A', rates[0]],
      ['B', rates[1]],
      ['C', rates[2]],
    ],
  })
  try {
    return exportXlsx(workbook.exportSnapshot())
  } finally {
    workbook.dispose()
  }
}

function buildSparseExternalSourceWorkbook(): Uint8Array {
  const workbook = WorkPaper.buildFromSheets({
    Rates: [
      ['SKU', 'Rate'],
      ['A', 20],
      ['B', null],
      ['C', 50],
      ['D', 0],
    ],
  })
  try {
    return replaceCellXml(exportXlsx(workbook.exportSnapshot()), 'xl/worksheets/sheet1.xml', 'B5', '<c r="B5" t="e"><v>#N/A</v></c>')
  } finally {
    workbook.dispose()
  }
}

function buildExternalLinkRangeCacheWorkbook(target: string): Uint8Array {
  const workbook = WorkPaper.buildFromSheets({
    Model: [
      [null, 2, 120],
      [null, null, 40],
      [null, null, 60],
    ],
  })
  try {
    const zip = unzipSync(exportXlsx(workbook.exportSnapshot()))
    zip['xl/worksheets/sheet1.xml'] = strToU8(
      strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
        .replace(/<c\b[^>]*\br=(["'])C1\1[^>]*>[\s\S]*?<\/c>/u, '<c r="C1"><f>SUM(\'[1]Rates\'!$B$2:$B$4)*B1</f><v>120</v></c>')
        .replace(
          /<c\b[^>]*\br=(["'])C2\1[^>]*>[\s\S]*?<\/c>/u,
          "<c r=\"C2\"><f>_xlfn.XLOOKUP(&quot;B&quot;,'[1]Rates'!$A$2:$A$4,'[1]Rates'!$B$2:$B$4)*B1</f><v>40</v></c>",
        )
        .replace(
          /<c\b[^>]*\br=(["'])C3\1[^>]*>[\s\S]*?<\/c>/u,
          "<c r=\"C3\"><f>SUMIFS('[1]Rates'!$B$2:$B$4,'[1]Rates'!$A$2:$A$4,&quot;C&quot;)*B1</f><v>60</v></c>",
        ),
    )
    zip['xl/workbook.xml'] = strToU8(
      ensureRelationshipNamespace(strFromU8(zip['xl/workbook.xml'] ?? new Uint8Array())).replace(
        '</sheets>',
        '</sheets><externalReferences><externalReference r:id="rId99"/></externalReferences>',
      ),
    )
    zip['xl/_rels/workbook.xml.rels'] = strToU8(
      strFromU8(zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array()).replace(
        '</Relationships>',
        '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink5.xml"/></Relationships>',
      ),
    )
    zip['xl/externalLinks/externalLink5.xml'] = strToU8(
      [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        `<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">`,
        '<externalBook r:id="rId1">',
        '<sheetNames><sheetName val="Rates"/></sheetNames>',
        '<sheetDataSet><sheetData sheetId="0">',
        '<row r="1"><cell r="A1" t="str"><v>SKU</v></cell><cell r="B1" t="str"><v>Rate</v></cell></row>',
        '<row r="2"><cell r="A2" t="str"><v>A</v></cell><cell r="B2"><v>10</v></cell></row>',
        '<row r="3"><cell r="A3" t="str"><v>B</v></cell><cell r="B3"><v>20</v></cell></row>',
        '<row r="4"><cell r="A4" t="str"><v>C</v></cell><cell r="B4"><v>30</v></cell></row>',
        '</sheetData></sheetDataSet>',
        '</externalBook>',
        '</externalLink>',
      ].join(''),
    )
    zip['xl/externalLinks/_rels/externalLink5.xml.rels'] = strToU8(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="${target}" TargetMode="External"/>` +
        '</Relationships>',
    )
    return zipSync(zip)
  } finally {
    workbook.dispose()
  }
}

function buildSparseExternalLinkRangeCacheWorkbook(target: string): Uint8Array {
  const workbook = WorkPaper.buildFromSheets({
    Model: [
      [null, 1, 60],
      [null, null, 60],
    ],
  })
  try {
    const zip = unzipSync(exportXlsx(workbook.exportSnapshot()))
    zip['xl/worksheets/sheet1.xml'] = strToU8(
      strFromU8(zip['xl/worksheets/sheet1.xml'] ?? new Uint8Array())
        .replace(/<c\b[^>]*\br=(["'])C1\1[^>]*>[\s\S]*?<\/c>/u, '<c r="C1"><f>SUM(\'[1]Rates\'!$B$2:$B$4)*B1</f><v>60</v></c>')
        .replace(/<c\b[^>]*\br=(["'])C2\1[^>]*>[\s\S]*?<\/c>/u, '<c r="C2"><f>IFERROR(SUM(\'[1]Rates\'!$B$2:$B$5),99)</f><v>60</v></c>'),
    )
    zip['xl/workbook.xml'] = strToU8(
      ensureRelationshipNamespace(strFromU8(zip['xl/workbook.xml'] ?? new Uint8Array())).replace(
        '</sheets>',
        '</sheets><externalReferences><externalReference r:id="rId99"/></externalReferences>',
      ),
    )
    zip['xl/_rels/workbook.xml.rels'] = strToU8(
      strFromU8(zip['xl/_rels/workbook.xml.rels'] ?? new Uint8Array()).replace(
        '</Relationships>',
        '<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink5.xml"/></Relationships>',
      ),
    )
    zip['xl/externalLinks/externalLink5.xml'] = strToU8(
      [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        `<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${officeRelationshipNamespace}">`,
        '<externalBook r:id="rId1">',
        '<sheetNames><sheetName val="Rates"/></sheetNames>',
        '<sheetDataSet><sheetData sheetId="0">',
        '<row r="2"><cell r="B2"><v>10</v></cell></row>',
        '<row r="3"><cell r="B3"><v>20</v></cell></row>',
        '<row r="4"><cell r="B4"><v>30</v></cell></row>',
        '<row r="5"><cell r="B5"><v>40</v></cell></row>',
        '</sheetData></sheetDataSet>',
        '</externalBook>',
        '</externalLink>',
      ].join(''),
    )
    zip['xl/externalLinks/_rels/externalLink5.xml.rels'] = strToU8(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="${target}" TargetMode="External"/>` +
        '</Relationships>',
    )
    return zipSync(zip)
  } finally {
    workbook.dispose()
  }
}

function ensureRelationshipNamespace(xml: string): string {
  return xml.replace(/<workbook\b([^>]*)>/u, (match) =>
    match.includes('xmlns:r=') ? match : match.replace('>', ` xmlns:r="${officeRelationshipNamespace}">`),
  )
}

function readExternalLinkCacheCellValue(bytes: Uint8Array, address: string): string | null {
  const xml = readExternalLinkCacheXml(bytes)
  const match = new RegExp(`<cell\\b(?=[^>]*\\br="${address}")[\\s\\S]*?<v>([\\s\\S]*?)<\\/v>[\\s\\S]*?<\\/cell>`, 'u').exec(xml)
  return match?.[1] ?? null
}

function readExternalLinkCacheXml(bytes: Uint8Array): string {
  const zip = unzipSync(bytes)
  return strFromU8(zip['xl/externalLinks/externalLink5.xml'] ?? new Uint8Array())
}
