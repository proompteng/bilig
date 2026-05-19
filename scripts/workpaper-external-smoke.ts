#!/usr/bin/env bun

import { copyFileSync, mkdirSync, mkdtempSync, readdirSync, rmSync, statSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'

import { assertAlignedVersions, loadRuntimePackages, type RuntimePackageManifest } from './runtime-package-set.ts'
import {
  parseJsonRecord,
  parseNodeFormulaDiagnosticsOutput,
  parseNodeHttpJsonSummaryOutput,
  parseNodeAgentToolCallOutput,
  parseNodeAgentVerificationOutput,
  parseNodeJsonFileOutput,
  parseNodeMarkdownReportOutput,
  parseNodeMcpChallengeOutput,
  parseNodeMcpStdioOutput,
  parseNodeMcpTranscriptOutput,
  parseNodePersistenceOutput,
  parseNodeRangeReadbackOutput,
  parseNodeRevenueScenarioOutput,
  parseNodeSheetInspectionOutput,
  parseNodeSmokeOutput,
  parseNodeSnapshotDiffOutput,
  parseNodeSnapshotImportOutput,
  parseNodeXlsxImportOutput,
  type AgentToolCallSummary,
  type AgentVerificationSummary,
  type McpTranscriptSummary,
  type RevenueScenarioSummary,
} from './workpaper-external-smoke-parsers.ts'
import { writeXlsxImportScript } from './workpaper-external-smoke-fixtures.ts'
import { parseNodeMcpStdioErrorOutput } from './workpaper-external-smoke-mcp-parsers.ts'
import { resolveKeepWorkpaperSmokeStage } from './workpaper-external-smoke-config.ts'

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const packDir = join(rootDir, 'build', 'npm-packages-runtime-smoke')
const headlessExampleDir = join(rootDir, 'examples', 'headless-workpaper')
const keepStage = resolveKeepStageOrExit()
const stageDir = mkdtempSync(join(tmpdir(), 'bilig-workpaper-external-smoke-'))
const textDecoder = new TextDecoder()

const runtimePackages = loadRuntimePackages(rootDir)
const alignedVersion = assertAlignedVersions(runtimePackages)

rmSync(packDir, { recursive: true, force: true })
mkdirSync(packDir, { recursive: true })

try {
  const tarballPaths = packRuntimePackages(runtimePackages)
  const nodeProjectDir = join(stageDir, 'node-consumer')
  const viteProjectDir = join(stageDir, 'vite-consumer')

  const nodeSummary = runNodeSmoke(nodeProjectDir, tarballPaths)
  const viteSummary = runViteSmoke(viteProjectDir, tarballPaths)

  console.log(
    JSON.stringify(
      {
        version: alignedVersion,
        packDir,
        stageDir: keepStage ? stageDir : undefined,
        node: nodeSummary,
        vite: viteSummary,
      },
      null,
      2,
    ),
  )
} finally {
  if (!keepStage) {
    rmSync(stageDir, { recursive: true, force: true })
  }
}

function resolveKeepStageOrExit(): boolean {
  try {
    return resolveKeepWorkpaperSmokeStage(process.env)
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error))
    process.exit(1)
  }
}

function packRuntimePackages(runtimePackageSet: RuntimePackageManifest[]): string[] {
  for (const runtimePackage of runtimePackageSet) {
    const packageDir = join(rootDir, runtimePackage.dir)
    runCommand('pnpm', ['pack', '--pack-destination', packDir], { cwd: packageDir })
  }

  const tarballIndex = new Map<string, string>()
  for (const tarballName of readdirSync(packDir).filter((entry) => entry.endsWith('.tgz'))) {
    const tarballPath = join(packDir, tarballName)
    const packedManifest = readPackedManifest(tarballPath)
    if (!packedManifest.name || !packedManifest.version) {
      throw new Error(`Packed tarball is missing name/version: ${tarballPath}`)
    }
    tarballIndex.set(packedManifest.name, tarballPath)
  }

  return runtimePackageSet.map((runtimePackage) => {
    const tarballPath = tarballIndex.get(runtimePackage.name)
    if (!tarballPath) {
      throw new Error(`Missing packed tarball for ${runtimePackage.name}`)
    }
    return tarballPath
  })
}

function runNodeSmoke(
  projectDir: string,
  tarballPaths: string[],
): {
  persistence: {
    afterRestoreAndEdit: {
      annualizedRunRate: number
      expansionAdjustedArr: number
      quarterNetMrr: number
    }
    beforeSave: {
      annualizedRunRate: number
      expansionAdjustedArr: number
      quarterNetMrr: number
    }
    persistedNamedExpressions: string[]
    persistedSheets: string[]
    saveFileBytes: number
  }
  scenarios: {
    afterEdit: RevenueScenarioSummary
    beforeEdit: RevenueScenarioSummary
    persistedSheets: string[]
    serializedBytes: number
  }
  agentToolCall: AgentToolCallSummary
  agentVerification: AgentVerificationSummary
  output: {
    afterAgentEdit: {
      enterpriseArpa: number
      qualifiedCustomerCounts: number[]
      targetRevenue: number
      totalRevenue: number
      westCustomers: number
    }
    initial: {
      targetRevenue: number
      totalRevenue: number
      westCustomers: number
    }
    persistedNamedExpressions: string[]
    persistedSheets: string[]
    restoredGrowthRatePercent: number
  }
  npmEval: {
    before: number
    after: number
    afterRestore: number
    sheets: string[]
    bytes: number
    verified: boolean
  }
  exceljsFormulaRecalc: {
    cachedResult: number
    readback: number
    workbookMutated: boolean
  }
  projectDir: string
  snapshotImport: {
    currencyLabel: string
    firstPeriod: number
    secondPeriod: number
    totalValue: number
    updatedFirstPeriod: number
  }
  xlsxImport: {
    currencyLabel: string
    editedTotalValue: number
    exportedBytes: number
    headerPeriod: number
    heightFeet: number
    firstPeriod: number
    roundTripSheetNames: string[]
    roundTripTotalValue: number
    secondPeriod: number
    totalValue: number
    translatedStructuredRefs: boolean
  }
  httpJsonSummary: {
    computed: {
      committedMrr: number
      largestOpportunityMrr: number
      weightedPipelineMrr: number
      westSeats: number
    }
    sourceRecords: number
    verified: boolean
  }
  formulaDiagnostics: {
    invalidDiagnostics: {
      code: string
      errorText: string
      functionName: string
      references: string[]
    }[]
    invalidDisplay: string
    validDisplay: string
    validValue: number
    verified: boolean
  }
  jsonFile: {
    computed: {
      committedMrr: number
      largestOpportunityMrr: number
      weightedPipelineMrr: number
      westSeats: number
    }
    source: string
    sourceRecords: number
    verified: boolean
  }
  markdownReport: {
    report: string
    verified: boolean
  }
  mcpStdio: {
    editedCell: string
    initialized: boolean
    toolNames: string[]
    verified: {
      expectedArrChanged: boolean
      formulasPersisted: boolean
      newValue: number
      previousValue: number
      restoredMatchesAfter: boolean
      serializedBytes: number
    }
  }
  packageMcpStdio: {
    editedCell: string
    initialized: boolean
    toolNames: string[]
    verified: {
      expectedArrChanged: boolean
      formulasPersisted: boolean
      newValue: number
      previousValue: number
      restoredMatchesAfter: boolean
      serializedBytes: number
    }
  }
  mcpChallenge: {
    editedCell: string
    dependentCell: string
    before: number
    after: number
    afterRestart: number
    displayValue: string
    toolNames: string[]
    resourceUris: string[]
    promptNames: string[]
    persistence: {
      persisted: boolean
      serializedBytes: number
    }
    verified: boolean
  }
  mcpStdioErrors: {
    invalidJson: {
      code: number
      id: null
    }
    invalidRequest: {
      code: number
      id: null
    }
  }
  mcpTranscript: McpTranscriptSummary
  rangeReadback: {
    range: string
    serializedReadback: unknown[][]
    valueReadback: unknown[][]
    verified: boolean
  }
  sheetInspection: {
    lookup: {
      dimensions: {
        height: number
        width: number
      }
      query: string
      sheetId: number
      sheetName: string
    }
    restoredSheets: string[]
    verified: boolean
  }
  snapshotDiff: {
    afterSerializedInput: number
    beforeSerializedInput: number
    changedCell: string
    changedSummaryValues: {
      after: {
        annualizedArr: number
        netMrr: number
      }
      before: {
        annualizedArr: number
        netMrr: number
      }
    }
    documentBytes: {
      after: number
      before: number
    }
    verified: boolean
  }
} {
  mkdirSync(projectDir, { recursive: true })
  mkdirSync(join(projectDir, 'fixtures'), { recursive: true })
  copyFileSync(join(headlessExampleDir, 'package.json'), join(projectDir, 'package.json'))
  copyFileSync(join(headlessExampleDir, 'tsconfig.json'), join(projectDir, 'tsconfig.json'))
  for (const entry of readdirSync(headlessExampleDir).filter((fileName) => fileName.endsWith('.ts'))) {
    copyFileSync(join(headlessExampleDir, entry), join(projectDir, entry))
  }
  copyFileSync(join(headlessExampleDir, 'fixtures', 'opportunities.json'), join(projectDir, 'fixtures', 'opportunities.json'))
  writeFileSync(
    join(projectDir, 'snapshot-import.ts'),
    [
      'import { ValueTag, type WorkbookSnapshot } from "@bilig/protocol";',
      'import { WorkPaper } from "@bilig/headless";',
      '',
      'const snapshot: WorkbookSnapshot = {',
      '  version: 1,',
      '  workbook: {',
      '    name: "Structured Financial Model",',
      '    metadata: {',
      '      definedNames: [',
      '        { name: "Currency", value: { kind: "cell-ref", sheetName: "Constants", address: "F7" } },',
      '        { name: "Start_Year", value: { kind: "cell-ref", sheetName: "Constants", address: "B10" } },',
      '      ],',
      '      tables: [',
      '        {',
      '          name: "tblActuals",',
      '          sheetName: "Imports",',
      '          startAddress: "A6",',
      '          endAddress: "D8",',
      '          columnNames: ["Account", "Value", "Year", "Period"],',
      '          headerRow: true,',
      '          totalsRow: false,',
      '        },',
      '      ],',
      '    },',
      '  },',
      '  sheets: [',
      '    {',
      '      id: 1,',
      '      name: "Constants",',
      '      order: 0,',
      '      cells: [',
      '        { address: "B10", value: 2012 },',
      '        { address: "F7", value: "USD" },',
      '        { address: "F9", formula: "Currency & \\"  000s\\"" },',
      '      ],',
      '    },',
      '    {',
      '      id: 2,',
      '      name: "Imports",',
      '      order: 1,',
      '      cells: [',
      '        { address: "A6", value: "Account" },',
      '        { address: "B6", value: "Value" },',
      '        { address: "C6", value: "Year" },',
      '        { address: "D6", value: "Period" },',
      '        { address: "A7", value: "Revenue" },',
      '        { address: "B7", value: 100 },',
      '        { address: "C7", value: 2011 },',
      '        { address: "D7", formula: "\'Imports\'!C7-Start_Year+1" },',
      '        { address: "A8", value: "Revenue" },',
      '        { address: "B8", value: 125 },',
      '        { address: "C8", value: 2012 },',
      '        { address: "D8", formula: "\'Imports\'!C8-Start_Year+1" },',
      '        { address: "F10", formula: "SUM(\'Imports\'!B7:B8)" },',
      '      ],',
      '    },',
      '  ],',
      '};',
      '',
      'const workbook = WorkPaper.buildFromSnapshot(snapshot, { maxRows: 20, maxColumns: 8, useColumnIndex: true });',
      'const constantsId = workbook.getSheetId("Constants");',
      'const importsId = workbook.getSheetId("Imports");',
      'if (constantsId === undefined || importsId === undefined) throw new Error("Imported snapshot sheets are missing");',
      'const read = (sheet: number, row: number, col: number) => workbook.getCellValue({ sheet, row, col });',
      'const currencyLabel = read(constantsId, 8, 5);',
      'const firstPeriod = read(importsId, 6, 3);',
      'const secondPeriod = read(importsId, 7, 3);',
      'const totalValue = read(importsId, 9, 5);',
      'workbook.setCellContents({ sheet: importsId, row: 6, col: 2 }, 2013);',
      'const updatedFirstPeriod = read(importsId, 6, 3);',
      'if (currencyLabel.tag !== ValueTag.String || firstPeriod.tag !== ValueTag.Number || secondPeriod.tag !== ValueTag.Number || totalValue.tag !== ValueTag.Number || updatedFirstPeriod.tag !== ValueTag.Number) {',
      '  throw new Error(`Unexpected snapshot values: ${JSON.stringify({ currencyLabel, firstPeriod, secondPeriod, totalValue, updatedFirstPeriod })}`);',
      '}',
      'console.log(JSON.stringify({',
      '  currencyLabel: currencyLabel.value,',
      '  firstPeriod: firstPeriod.value,',
      '  secondPeriod: secondPeriod.value,',
      '  totalValue: totalValue.value,',
      '  updatedFirstPeriod: updatedFirstPeriod.value,',
      '}));',
      '',
    ].join('\n'),
  )
  writeXlsxImportScript(projectDir)
  writeExceljsFormulaRecalcScript(projectDir)

  installTarballs(projectDir, tarballPaths, ['exceljs@4.4.0'])
  const npmEval = parseNodeNpmEvalOutput(runTextCommand('npm', ['run', '--silent', 'npm-eval'], { cwd: projectDir }))
  const exceljsFormulaRecalc = parseNodeExceljsFormulaRecalcOutput(
    runTextCommand('npx', ['--no-install', 'tsx', 'exceljs-formula-recalc.ts'], { cwd: projectDir }),
  )
  const output = parseNodeSmokeOutput(runTextCommand('npm', ['run', '--silent', 'start'], { cwd: projectDir }))
  const persistence = parseNodePersistenceOutput(runTextCommand('npm', ['run', '--silent', 'persistence'], { cwd: projectDir }))
  const scenarios = parseNodeRevenueScenarioOutput(runTextCommand('npm', ['run', '--silent', 'scenarios'], { cwd: projectDir }))
  const agentToolCall = parseNodeAgentToolCallOutput(runTextCommand('npm', ['run', '--silent', 'agent:tool-call'], { cwd: projectDir }))
  runTextCommand('npm', ['run', '--silent', 'agent:framework-adapters'], { cwd: projectDir })
  runTextCommand('npm', ['run', '--silent', 'typecheck'], { cwd: projectDir })
  runTextCommand('npm', ['run', '--silent', 'agent:mcp-tools'], { cwd: projectDir })
  const mcpStdio = parseNodeMcpStdioOutput(
    runTextCommand(
      'sh',
      [
        '-c',
        [
          "printf '%s\\n'",
          '\'{"jsonrpc":"2.0","id":1,"method":"initialize"}\'',
          '\'{"jsonrpc":"2.0","method":"notifications/initialized"}\'',
          '\'{"jsonrpc":"2.0","id":2,"method":"tools/list"}\'',
          '\'{"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"set_workpaper_input_cell","arguments":{"sheetName":"Inputs","address":"B3","value":0.4}}}\'',
          '|',
          'npm run --silent agent:mcp-stdio',
        ].join(' '),
      ],
      { cwd: projectDir },
    ),
  )
  const packageMcpStdio = parseNodeMcpStdioOutput(
    runTextCommand(
      'sh',
      [
        '-c',
        [
          "printf '%s\\n'",
          '\'{"jsonrpc":"2.0","id":1,"method":"initialize"}\'',
          '\'{"jsonrpc":"2.0","method":"notifications/initialized"}\'',
          '\'{"jsonrpc":"2.0","id":2,"method":"tools/list"}\'',
          '\'{"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"set_workpaper_input_cell","arguments":{"sheetName":"Inputs","address":"B3","value":0.4}}}\'',
          '|',
          './node_modules/.bin/bilig-workpaper-mcp',
        ].join(' '),
      ],
      { cwd: projectDir },
    ),
    { expectedServerName: 'bilig-headless-workpaper' },
  )
  const mcpChallenge = parseNodeMcpChallengeOutput(runTextCommand('./node_modules/.bin/bilig-mcp-challenge', [], { cwd: projectDir }))
  const mcpTranscript = parseNodeMcpTranscriptOutput(
    runTextCommand('npm', ['run', '--silent', 'agent:mcp-transcript'], { cwd: projectDir }),
  )
  const mcpStdioErrors = parseNodeMcpStdioErrorOutput(
    runTextCommand(
      'sh',
      [
        '-c',
        ["printf '%s\\n'", "'{not-json'", '\'{"jsonrpc":"2.0","id":4,"params":{}}\'', '|', 'npm run --silent agent:mcp-stdio'].join(' '),
      ],
      { cwd: projectDir },
    ),
  )
  const agentVerification = parseNodeAgentVerificationOutput(
    runTextCommand('npm', ['run', '--silent', 'agent:verify'], { cwd: projectDir }),
  )
  const formulaDiagnostics = parseNodeFormulaDiagnosticsOutput(
    runTextCommand('npm', ['run', '--silent', 'formula-diagnostics'], { cwd: projectDir }),
  )
  const httpJsonSummary = parseNodeHttpJsonSummaryOutput(
    runTextCommand('npm', ['run', '--silent', 'http-json-summary'], { cwd: projectDir }),
  )
  const jsonFile = parseNodeJsonFileOutput(runTextCommand('npm', ['run', '--silent', 'json-file'], { cwd: projectDir }))
  const markdownReport = parseNodeMarkdownReportOutput(runTextCommand('npm', ['run', '--silent', 'markdown-report'], { cwd: projectDir }))
  const rangeReadback = parseNodeRangeReadbackOutput(runTextCommand('npm', ['run', '--silent', 'range-readback'], { cwd: projectDir }))
  const sheetInspection = parseNodeSheetInspectionOutput(
    runTextCommand('npm', ['run', '--silent', 'sheet-inspection'], { cwd: projectDir }),
  )
  const snapshotDiff = parseNodeSnapshotDiffOutput(runTextCommand('npm', ['run', '--silent', 'snapshot-diff'], { cwd: projectDir }))
  const snapshotImport = parseNodeSnapshotImportOutput(
    runTextCommand('npx', ['--no-install', 'tsx', 'snapshot-import.ts'], { cwd: projectDir }),
  )
  const xlsxImport = parseNodeXlsxImportOutput(runTextCommand('npx', ['--no-install', 'tsx', 'xlsx-import.ts'], { cwd: projectDir }))

  return {
    agentToolCall,
    agentVerification,
    formulaDiagnostics,
    httpJsonSummary,
    jsonFile,
    markdownReport,
    mcpStdio,
    mcpStdioErrors,
    mcpTranscript,
    mcpChallenge,
    packageMcpStdio,
    persistence,
    projectDir,
    exceljsFormulaRecalc,
    npmEval,
    rangeReadback,
    scenarios,
    sheetInspection,
    snapshotDiff,
    snapshotImport,
    xlsxImport,
    output,
  }
}

function writeExceljsFormulaRecalcScript(projectDir: string): void {
  writeFileSync(
    join(projectDir, 'exceljs-formula-recalc.ts'),
    [
      'import ExcelJS from "exceljs";',
      'import { recalculateExceljsWorkbook } from "exceljs-formula-recalc";',
      '',
      'const workbook = new ExcelJS.Workbook();',
      'const inputs = workbook.addWorksheet("Inputs");',
      'inputs.getCell("A1").value = "Metric";',
      'inputs.getCell("B1").value = "Value";',
      'inputs.getCell("A2").value = "Units";',
      'inputs.getCell("B2").value = 40;',
      'inputs.getCell("A3").value = "Price";',
      'inputs.getCell("B3").value = 1200;',
      'const summary = workbook.addWorksheet("Summary");',
      'summary.getCell("A1").value = "Metric";',
      'summary.getCell("B1").value = "Value";',
      'summary.getCell("A2").value = "Revenue";',
      'summary.getCell("B2").value = { formula: "Inputs!B2*Inputs!B3", result: 48000 };',
      '',
      'const result = await recalculateExceljsWorkbook(workbook, {',
      '  edits: [',
      '    { target: "Inputs!B2", value: 48 },',
      '    { target: "Inputs!B3", value: 1500 },',
      '  ],',
      '  reads: ["Summary!B2"],',
      '});',
      '',
      'const readbackCell = result.reads["Summary!B2"];',
      'const readback = typeof readbackCell === "object" && readbackCell !== null && "value" in readbackCell ? readbackCell.value : null;',
      'const cachedCell = workbook.getWorksheet("Summary")?.getCell("B2").value;',
      'const cachedResult = typeof cachedCell === "object" && cachedCell !== null && "result" in cachedCell ? cachedCell.result : null;',
      'console.log(JSON.stringify({ cachedResult, readback, workbookMutated: result.workbookMutated }));',
      '',
    ].join('\n'),
  )
}

function parseNodeNpmEvalOutput(output: string): {
  before: number
  after: number
  afterRestore: number
  sheets: string[]
  bytes: number
  verified: boolean
} {
  const record = parseJsonRecord(output, 'npm eval output')
  const before = record.before
  const after = record.after
  const afterRestore = record.afterRestore
  const sheets = record.sheets
  const bytes = record.bytes
  const verified = record.verified

  if (
    before !== 24000 ||
    after !== 38400 ||
    afterRestore !== 38400 ||
    !Array.isArray(sheets) ||
    !sheets.every((sheet): sheet is string => typeof sheet === 'string') ||
    sheets.join(',') !== 'Inputs,Summary' ||
    typeof bytes !== 'number' ||
    bytes <= 0 ||
    verified !== true
  ) {
    throw new Error(`Unexpected npm eval output: ${output}`)
  }

  return {
    before,
    after,
    afterRestore,
    sheets,
    bytes,
    verified,
  }
}

function parseNodeExceljsFormulaRecalcOutput(output: string): {
  cachedResult: number
  readback: number
  workbookMutated: boolean
} {
  const record = parseJsonRecord(output, 'ExcelJS formula recalculation output')
  const cachedResult = record.cachedResult
  const readback = record.readback
  const workbookMutated = record.workbookMutated
  if (cachedResult !== 72000 || readback !== 72000 || workbookMutated !== true) {
    throw new Error(`Unexpected ExcelJS formula recalculation output: ${output}`)
  }

  return {
    cachedResult,
    readback,
    workbookMutated,
  }
}

function runViteSmoke(
  projectDir: string,
  tarballPaths: string[],
): {
  distFiles: string[]
  projectDir: string
  wasmAssets: string[]
} {
  mkdirSync(join(projectDir, 'src'), { recursive: true })
  writeFileSync(
    join(projectDir, 'package.json'),
    `${JSON.stringify(
      {
        name: 'workpaper-vite-consumer',
        private: true,
        type: 'module',
        scripts: {
          build: 'vite build',
          typecheck: 'tsc --noEmit',
        },
      },
      null,
      2,
    )}\n`,
  )
  writeFileSync(
    join(projectDir, 'tsconfig.json'),
    `${JSON.stringify(
      {
        compilerOptions: {
          target: 'ESNext',
          module: 'ESNext',
          moduleResolution: 'Bundler',
          strict: true,
          lib: ['DOM', 'ESNext'],
          types: ['vite/client', 'node'],
          noEmit: true,
        },
        include: ['src', 'vite.config.ts'],
      },
      null,
      2,
    )}\n`,
  )
  writeFileSync(
    join(projectDir, 'vite.config.ts'),
    [
      'import { defineConfig } from "vite";',
      '',
      'export default defineConfig({',
      '  build: {',
      '    target: "es2022",',
      '  },',
      '});',
      '',
    ].join('\n'),
  )
  writeFileSync(
    join(projectDir, 'index.html'),
    [
      '<!doctype html>',
      '<html lang="en">',
      '  <head>',
      '    <meta charset="UTF-8" />',
      '    <meta name="viewport" content="width=device-width, initial-scale=1.0" />',
      '    <title>WorkPaper smoke</title>',
      '  </head>',
      '  <body>',
      '    <div id="app"></div>',
      '    <script type="module" src="/src/main.ts"></script>',
      '  </body>',
      '</html>',
      '',
    ].join('\n'),
  )
  writeFileSync(
    join(projectDir, 'src', 'main.ts'),
    [
      'import { ValueTag } from "@bilig/protocol";',
      'import { WorkPaper } from "@bilig/headless";',
      '',
      'const workbook = WorkPaper.buildFromSheets({',
      '  Dashboard: [["Label", "Value"], ["Total", "=SUM(10,20,30)"], ["Top", "=MAX(10,20,30)"]],',
      '});',
      '',
      'const sheetId = workbook.getSheetId("Dashboard");',
      'if (sheetId === undefined) {',
      '  throw new Error("Dashboard sheet is missing");',
      '}',
      '',
      'const total = workbook.getCellValue({ sheet: sheetId, row: 1, col: 1 });',
      'const top = workbook.getCellValue({ sheet: sheetId, row: 2, col: 1 });',
      '',
      'if (total.tag !== ValueTag.Number || top.tag !== ValueTag.Number) {',
      '  throw new Error(`Unexpected dashboard values: ${JSON.stringify({ total, top })}`);',
      '}',
      '',
      'document.querySelector<HTMLDivElement>("#app")!.textContent = `${total.value}:${top.value}`;',
      '',
    ].join('\n'),
  )

  installTarballs(projectDir, tarballPaths, ['vite@8.0.9', 'typescript@6.0.2', '@types/node@25.5.0'])
  runCommand('npm', ['run', 'typecheck'], { cwd: projectDir })
  runCommand('npm', ['run', 'build'], { cwd: projectDir })

  const distDir = join(projectDir, 'dist')
  const distFiles = listFilesRecursive(distDir).map((filePath) => filePath.slice(distDir.length + 1))
  const wasmAssets = distFiles.filter((entry) => entry.endsWith('.wasm'))
  if (wasmAssets.length === 0) {
    throw new Error(`Vite build did not emit a wasm asset: ${distFiles.join(', ')}`)
  }

  return {
    projectDir,
    distFiles,
    wasmAssets,
  }
}

function installTarballs(projectDir: string, tarballPaths: string[], extraPackages: string[] = []): void {
  runCommand('npm', ['install', '--no-package-lock', ...tarballPaths, ...extraPackages], {
    cwd: projectDir,
  })
}

function listFilesRecursive(targetDir: string): string[] {
  const files: string[] = []
  for (const entry of readdirSync(targetDir)) {
    const entryPath = join(targetDir, entry)
    if (statSync(entryPath).isDirectory()) {
      files.push(...listFilesRecursive(entryPath))
      continue
    }
    files.push(entryPath)
  }
  return files
}

function readPackedManifest(tarballPath: string): { name: string; version: string } {
  const parsed = parseJsonRecord(runTextCommand('tar', ['-xOf', tarballPath, 'package/package.json']), `packed manifest in ${tarballPath}`)
  const name = parsed.name
  const version = parsed.version
  if (typeof name !== 'string' || typeof version !== 'string') {
    throw new Error(`Packed tarball is missing name/version: ${tarballPath}`)
  }
  return { name, version }
}

function runTextCommand(command: string, args: string[], options: { cwd?: string } = {}): string {
  const result = Bun.spawnSync([command, ...args], {
    cwd: options.cwd ?? rootDir,
    env: process.env,
    stdin: 'ignore',
    stdout: 'pipe',
    stderr: 'pipe',
  })
  if (result.exitCode !== 0) {
    const stderr = textDecoder.decode(result.stderr).trim()
    throw new Error(`Command failed: ${command} ${args.join(' ')}${stderr ? `\n${stderr}` : ''}`)
  }
  return textDecoder.decode(result.stdout).trim()
}

function runCommand(command: string, args: string[], options: { cwd?: string } = {}): void {
  const result = Bun.spawnSync([command, ...args], {
    cwd: options.cwd ?? rootDir,
    env: process.env,
    stdin: 'inherit',
    stdout: 'inherit',
    stderr: 'inherit',
  })
  if (result.exitCode !== 0) {
    throw new Error(`Command failed: ${command} ${args.join(' ')}`)
  }
}
