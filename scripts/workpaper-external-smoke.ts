#!/usr/bin/env bun

import { copyFileSync, mkdirSync, mkdtempSync, readdirSync, rmSync, statSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'

import { assertAlignedVersions, loadRuntimePackages, type RuntimePackageManifest } from './runtime-package-set.ts'

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const packDir = join(rootDir, 'build', 'npm-packages-runtime-smoke')
const headlessExampleDir = join(rootDir, 'examples', 'headless-workpaper')
const stageDir = mkdtempSync(join(tmpdir(), 'bilig-workpaper-external-smoke-'))
const textDecoder = new TextDecoder()
const keepStage = process.env.KEEP_WORKPAPER_SMOKE_STAGE === 'true'

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
  projectDir: string
  snapshotImport: {
    currencyLabel: string
    firstPeriod: number
    secondPeriod: number
    totalValue: number
    updatedFirstPeriod: number
  }
} {
  mkdirSync(projectDir, { recursive: true })
  copyFileSync(join(headlessExampleDir, 'package.json'), join(projectDir, 'package.json'))
  copyFileSync(join(headlessExampleDir, 'agent-writeback-verification.mjs'), join(projectDir, 'agent-writeback-verification.mjs'))
  copyFileSync(join(headlessExampleDir, 'revenue-plan.mjs'), join(projectDir, 'revenue-plan.mjs'))
  copyFileSync(join(headlessExampleDir, 'persistence-roundtrip.mjs'), join(projectDir, 'persistence-roundtrip.mjs'))
  copyFileSync(join(headlessExampleDir, 'revenue-scenarios.mjs'), join(projectDir, 'revenue-scenarios.mjs'))
  writeFileSync(
    join(projectDir, 'snapshot-import.mjs'),
    [
      'import { ValueTag } from "@bilig/protocol";',
      'import { WorkPaper } from "@bilig/headless";',
      '',
      'const snapshot = {',
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
      'const read = (sheet, row, col) => workbook.getCellValue({ sheet, row, col });',
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

  installTarballs(projectDir, tarballPaths)
  const output = parseNodeSmokeOutput(runTextCommand('node', ['revenue-plan.mjs'], { cwd: projectDir }))
  const persistence = parseNodePersistenceOutput(runTextCommand('node', ['persistence-roundtrip.mjs'], { cwd: projectDir }))
  const scenarios = parseNodeRevenueScenarioOutput(runTextCommand('node', ['revenue-scenarios.mjs'], { cwd: projectDir }))
  const agentVerification = parseNodeAgentVerificationOutput(
    runTextCommand('node', ['agent-writeback-verification.mjs'], { cwd: projectDir }),
  )
  const snapshotImport = parseNodeSnapshotImportOutput(runTextCommand('node', ['snapshot-import.mjs'], { cwd: projectDir }))

  return {
    agentVerification,
    persistence,
    projectDir,
    scenarios,
    snapshotImport,
    output,
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

function parseNodeSmokeOutput(output: string): {
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
} {
  const parsed = parseJsonRecord(output, 'node smoke output')
  const initial = parseRecordValue(parsed.initial, 'node smoke initial output')
  const afterAgentEdit = parseRecordValue(parsed.afterAgentEdit, 'node smoke edited output')
  const persistedSheets = parsed.persistedSheets
  const persistedNamedExpressions = parsed.persistedNamedExpressions
  const qualifiedCustomerCounts = afterAgentEdit.qualifiedCustomerCounts

  if (
    typeof initial.totalRevenue !== 'number' ||
    typeof initial.westCustomers !== 'number' ||
    typeof initial.targetRevenue !== 'number' ||
    typeof afterAgentEdit.totalRevenue !== 'number' ||
    typeof afterAgentEdit.westCustomers !== 'number' ||
    typeof afterAgentEdit.enterpriseArpa !== 'number' ||
    typeof afterAgentEdit.targetRevenue !== 'number' ||
    !isNumberArray(qualifiedCustomerCounts) ||
    !isStringArray(persistedSheets) ||
    !isStringArray(persistedNamedExpressions) ||
    typeof parsed.restoredGrowthRatePercent !== 'number'
  ) {
    throw new Error(`Unexpected node smoke output: ${output}`)
  }

  return {
    initial: {
      totalRevenue: initial.totalRevenue,
      westCustomers: initial.westCustomers,
      targetRevenue: initial.targetRevenue,
    },
    afterAgentEdit: {
      totalRevenue: afterAgentEdit.totalRevenue,
      westCustomers: afterAgentEdit.westCustomers,
      enterpriseArpa: afterAgentEdit.enterpriseArpa,
      targetRevenue: afterAgentEdit.targetRevenue,
      qualifiedCustomerCounts,
    },
    persistedSheets,
    persistedNamedExpressions,
    restoredGrowthRatePercent: parsed.restoredGrowthRatePercent,
  }
}

function parseNodeSnapshotImportOutput(output: string): {
  currencyLabel: string
  firstPeriod: number
  secondPeriod: number
  totalValue: number
  updatedFirstPeriod: number
} {
  const parsed = parseJsonRecord(output, 'node snapshot import output')
  const currencyLabel = parsed.currencyLabel
  const firstPeriod = parsed.firstPeriod
  const secondPeriod = parsed.secondPeriod
  const totalValue = parsed.totalValue
  const updatedFirstPeriod = parsed.updatedFirstPeriod
  if (currencyLabel !== 'USD  000s' || firstPeriod !== 0 || secondPeriod !== 1 || totalValue !== 225 || updatedFirstPeriod !== 2) {
    throw new Error(`Unexpected node snapshot import output: ${output}`)
  }
  return {
    currencyLabel,
    firstPeriod,
    secondPeriod,
    totalValue,
    updatedFirstPeriod,
  }
}

function parseNodePersistenceOutput(output: string): {
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
} {
  const parsed = parseJsonRecord(output, 'node persistence output')
  const beforeSave = parseRecordValue(parsed.beforeSave, 'node persistence before-save output')
  const afterRestoreAndEdit = parseRecordValue(parsed.afterRestoreAndEdit, 'node persistence restored output')
  const persistedSheets = parsed.persistedSheets
  const persistedNamedExpressions = parsed.persistedNamedExpressions

  if (
    typeof beforeSave.quarterNetMrr !== 'number' ||
    typeof beforeSave.annualizedRunRate !== 'number' ||
    typeof beforeSave.expansionAdjustedArr !== 'number' ||
    typeof afterRestoreAndEdit.quarterNetMrr !== 'number' ||
    typeof afterRestoreAndEdit.annualizedRunRate !== 'number' ||
    typeof afterRestoreAndEdit.expansionAdjustedArr !== 'number' ||
    !isStringArray(persistedSheets) ||
    !isStringArray(persistedNamedExpressions) ||
    typeof parsed.saveFileBytes !== 'number' ||
    parsed.saveFileBytes <= 0
  ) {
    throw new Error(`Unexpected node persistence output: ${output}`)
  }

  return {
    beforeSave: {
      quarterNetMrr: beforeSave.quarterNetMrr,
      annualizedRunRate: beforeSave.annualizedRunRate,
      expansionAdjustedArr: beforeSave.expansionAdjustedArr,
    },
    afterRestoreAndEdit: {
      quarterNetMrr: afterRestoreAndEdit.quarterNetMrr,
      annualizedRunRate: afterRestoreAndEdit.annualizedRunRate,
      expansionAdjustedArr: afterRestoreAndEdit.expansionAdjustedArr,
    },
    persistedSheets,
    persistedNamedExpressions,
    saveFileBytes: parsed.saveFileBytes,
  }
}

type RevenueScenarioSummary = {
  annualRunRate: number
  enterpriseNetMrr: number
  expansionTarget: number
  scenarios: {
    conservativeNetMrr: number
    expansionNetMrr: number
    stretchNetMrr: number
  }
  totalNetMrr: number
}

type AgentVerificationProjection = {
  annualizedArr: number
  arrTargetDelta: number
  customers: number
  expansionMrr: number
  grossMrr: number
}

type AgentVerificationSummary = {
  after: AgentVerificationProjection
  before: AgentVerificationProjection
  edits: {
    after: number
    before: number
    cell: string
  }[]
  formulaContracts: {
    annualizedArr: string
    arrTargetDelta: string
    customers: string
    expansionMrr: string
    grossMrr: string
  }
  restored: AgentVerificationProjection
  verified: {
    formulasPersisted: boolean
    formulasUnchanged: boolean
    restoredMatchesAfter: boolean
    serializedBytes: number
  }
}

function parseNodeRevenueScenarioOutput(output: string): {
  afterEdit: RevenueScenarioSummary
  beforeEdit: RevenueScenarioSummary
  persistedSheets: string[]
  serializedBytes: number
} {
  const parsed = parseJsonRecord(output, 'node revenue scenario output')
  const beforeEdit = parseRevenueScenarioSummary(parsed.beforeEdit, 'before-edit revenue scenario output')
  const afterEdit = parseRevenueScenarioSummary(parsed.afterEdit, 'after-edit revenue scenario output')
  const persistedSheets = parsed.persistedSheets

  if (!isStringArray(persistedSheets) || typeof parsed.serializedBytes !== 'number' || parsed.serializedBytes <= 0) {
    throw new Error(`Unexpected node revenue scenario output: ${output}`)
  }

  return {
    beforeEdit,
    afterEdit,
    persistedSheets,
    serializedBytes: parsed.serializedBytes,
  }
}

function parseRevenueScenarioSummary(value: unknown, context: string): RevenueScenarioSummary {
  const parsed = parseRecordValue(value, context)
  const scenarios = parseRecordValue(parsed.scenarios, `${context} scenarios`)

  if (
    typeof parsed.totalNetMrr !== 'number' ||
    typeof parsed.annualRunRate !== 'number' ||
    typeof parsed.enterpriseNetMrr !== 'number' ||
    typeof parsed.expansionTarget !== 'number' ||
    typeof scenarios.conservativeNetMrr !== 'number' ||
    typeof scenarios.expansionNetMrr !== 'number' ||
    typeof scenarios.stretchNetMrr !== 'number'
  ) {
    throw new Error(`Unexpected ${context}: ${JSON.stringify(value)}`)
  }

  return {
    totalNetMrr: parsed.totalNetMrr,
    annualRunRate: parsed.annualRunRate,
    enterpriseNetMrr: parsed.enterpriseNetMrr,
    expansionTarget: parsed.expansionTarget,
    scenarios: {
      conservativeNetMrr: scenarios.conservativeNetMrr,
      expansionNetMrr: scenarios.expansionNetMrr,
      stretchNetMrr: scenarios.stretchNetMrr,
    },
  }
}

function parseNodeAgentVerificationOutput(output: string): AgentVerificationSummary {
  const parsed = parseJsonRecord(output, 'node agent verification output')
  const before = parseAgentVerificationProjection(parsed.before, 'node agent verification before output')
  const after = parseAgentVerificationProjection(parsed.after, 'node agent verification after output')
  const restored = parseAgentVerificationProjection(parsed.restored, 'node agent verification restored output')
  const edits = parseAgentVerificationEdits(parsed.edits)
  const formulaContracts = parseAgentVerificationFormulaContracts(parsed.formulaContracts)
  const verified = parseRecordValue(parsed.verified, 'node agent verification flags')

  if (
    before.customers !== 40 ||
    before.grossMrr !== 9600 ||
    before.expansionMrr !== 10560 ||
    before.annualizedArr !== 126720 ||
    before.arrTargetDelta !== -23280 ||
    after.customers !== 65 ||
    after.grossMrr !== 15600 ||
    after.expansionMrr !== 18720 ||
    after.annualizedArr !== 224640 ||
    after.arrTargetDelta !== 74640 ||
    restored.customers !== after.customers ||
    restored.grossMrr !== after.grossMrr ||
    restored.expansionMrr !== after.expansionMrr ||
    restored.annualizedArr !== after.annualizedArr ||
    restored.arrTargetDelta !== after.arrTargetDelta ||
    formulaContracts.customers !== '=Assumptions!B2*Assumptions!B3' ||
    formulaContracts.grossMrr !== '=B2*Assumptions!B4' ||
    formulaContracts.expansionMrr !== '=B3*Assumptions!B5' ||
    formulaContracts.annualizedArr !== '=B4*12' ||
    formulaContracts.arrTargetDelta !== '=Plan!B5-150000' ||
    verified.formulasUnchanged !== true ||
    verified.formulasPersisted !== true ||
    verified.restoredMatchesAfter !== true ||
    typeof verified.serializedBytes !== 'number' ||
    verified.serializedBytes <= 0
  ) {
    throw new Error(`Unexpected node agent verification output: ${output}`)
  }

  return {
    after,
    before,
    edits,
    formulaContracts,
    restored,
    verified: {
      formulasPersisted: verified.formulasPersisted,
      formulasUnchanged: verified.formulasUnchanged,
      restoredMatchesAfter: verified.restoredMatchesAfter,
      serializedBytes: verified.serializedBytes,
    },
  }
}

function parseAgentVerificationProjection(value: unknown, context: string): AgentVerificationProjection {
  const parsed = parseRecordValue(value, context)
  if (
    typeof parsed.customers !== 'number' ||
    typeof parsed.grossMrr !== 'number' ||
    typeof parsed.expansionMrr !== 'number' ||
    typeof parsed.annualizedArr !== 'number' ||
    typeof parsed.arrTargetDelta !== 'number'
  ) {
    throw new Error(`Unexpected ${context}: ${JSON.stringify(value)}`)
  }
  return {
    annualizedArr: parsed.annualizedArr,
    arrTargetDelta: parsed.arrTargetDelta,
    customers: parsed.customers,
    expansionMrr: parsed.expansionMrr,
    grossMrr: parsed.grossMrr,
  }
}

function parseAgentVerificationFormulaContracts(value: unknown): AgentVerificationSummary['formulaContracts'] {
  const parsed = parseRecordValue(value, 'node agent verification formula contracts')
  if (
    typeof parsed.customers !== 'string' ||
    typeof parsed.grossMrr !== 'string' ||
    typeof parsed.expansionMrr !== 'string' ||
    typeof parsed.annualizedArr !== 'string' ||
    typeof parsed.arrTargetDelta !== 'string'
  ) {
    throw new Error(`Unexpected node agent verification formula contracts: ${JSON.stringify(value)}`)
  }
  return {
    annualizedArr: parsed.annualizedArr,
    arrTargetDelta: parsed.arrTargetDelta,
    customers: parsed.customers,
    expansionMrr: parsed.expansionMrr,
    grossMrr: parsed.grossMrr,
  }
}

function parseAgentVerificationEdits(value: unknown): AgentVerificationSummary['edits'] {
  if (!Array.isArray(value)) {
    throw new Error(`Expected node agent verification edits to be an array: ${JSON.stringify(value)}`)
  }
  const edits = value.map((entry, index) => {
    const edit = parseRecordValue(entry, `node agent verification edit ${index + 1}`)
    if (typeof edit.cell !== 'string' || typeof edit.before !== 'number' || typeof edit.after !== 'number') {
      throw new Error(`Unexpected node agent verification edit: ${JSON.stringify(entry)}`)
    }
    return {
      after: edit.after,
      before: edit.before,
      cell: edit.cell,
    }
  })
  const expected = [
    { after: 650, before: 500, cell: 'Assumptions!B2' },
    { after: 0.1, before: 0.08, cell: 'Assumptions!B3' },
    { after: 1.2, before: 1.1, cell: 'Assumptions!B5' },
  ]
  if (JSON.stringify(edits) !== JSON.stringify(expected)) {
    throw new Error(`Unexpected node agent verification edits: ${JSON.stringify(value)}`)
  }
  return edits
}

function parseJsonRecord(serialized: string, context: string): Record<string, unknown> {
  const parsed: unknown = JSON.parse(serialized)
  return parseRecordValue(parsed, context)
}

function parseRecordValue(candidate: unknown, context: string): Record<string, unknown> {
  const parsed = candidate
  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
    throw new Error(`Expected ${context} to be a JSON object`)
  }
  const record: Record<string, unknown> = {}
  for (const [key, value] of Object.entries(parsed)) {
    record[key] = value
  }
  return record
}

function isNumberArray(value: unknown): value is number[] {
  return Array.isArray(value) && value.every((entry) => typeof entry === 'number')
}

function isStringArray(value: unknown): value is string[] {
  return Array.isArray(value) && value.every((entry) => typeof entry === 'string')
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
