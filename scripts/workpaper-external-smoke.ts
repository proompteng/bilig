#!/usr/bin/env bun

import { copyFileSync, mkdirSync, mkdtempSync, readdirSync, rmSync, statSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'

import { assertAlignedVersions, loadRuntimePackages, type RuntimePackageManifest } from './runtime-package-set.ts'
import {
  parseJsonRecord,
  parseNodeHttpJsonSummaryOutput,
  parseNodeAgentToolCallOutput,
  parseNodeAgentVerificationOutput,
  parseNodePersistenceOutput,
  parseNodeRevenueScenarioOutput,
  parseNodeSmokeOutput,
  parseNodeSnapshotImportOutput,
  parseNodeXlsxImportOutput,
  type AgentToolCallSummary,
  type AgentVerificationSummary,
  type RevenueScenarioSummary,
} from './workpaper-external-smoke-parsers.ts'
import { writeXlsxImportScript } from './workpaper-external-smoke-fixtures.ts'

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
    headerPeriod: number
    heightFeet: number
    firstPeriod: number
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
} {
  mkdirSync(projectDir, { recursive: true })
  copyFileSync(join(headlessExampleDir, 'package.json'), join(projectDir, 'package.json'))
  copyFileSync(join(headlessExampleDir, 'agent-tool-call-loop.mjs'), join(projectDir, 'agent-tool-call-loop.mjs'))
  copyFileSync(join(headlessExampleDir, 'agent-writeback-verification.mjs'), join(projectDir, 'agent-writeback-verification.mjs'))
  copyFileSync(join(headlessExampleDir, 'http-json-summary.mjs'), join(projectDir, 'http-json-summary.mjs'))
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
  writeXlsxImportScript(projectDir)

  installTarballs(projectDir, tarballPaths)
  const output = parseNodeSmokeOutput(runTextCommand('node', ['revenue-plan.mjs'], { cwd: projectDir }))
  const persistence = parseNodePersistenceOutput(runTextCommand('node', ['persistence-roundtrip.mjs'], { cwd: projectDir }))
  const scenarios = parseNodeRevenueScenarioOutput(runTextCommand('node', ['revenue-scenarios.mjs'], { cwd: projectDir }))
  const agentToolCall = parseNodeAgentToolCallOutput(runTextCommand('node', ['agent-tool-call-loop.mjs'], { cwd: projectDir }))
  const agentVerification = parseNodeAgentVerificationOutput(
    runTextCommand('node', ['agent-writeback-verification.mjs'], { cwd: projectDir }),
  )
  const httpJsonSummary = parseNodeHttpJsonSummaryOutput(runTextCommand('node', ['http-json-summary.mjs'], { cwd: projectDir }))
  const snapshotImport = parseNodeSnapshotImportOutput(runTextCommand('node', ['snapshot-import.mjs'], { cwd: projectDir }))
  const xlsxImport = parseNodeXlsxImportOutput(runTextCommand('node', ['xlsx-import.mjs'], { cwd: projectDir }))

  return {
    agentToolCall,
    agentVerification,
    httpJsonSummary,
    persistence,
    projectDir,
    scenarios,
    snapshotImport,
    xlsxImport,
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
