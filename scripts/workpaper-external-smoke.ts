#!/usr/bin/env bun

import { mkdirSync, mkdtempSync, readdirSync, rmSync, statSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'

import { assertAlignedVersions, loadRuntimePackages, type RuntimePackageManifest } from './runtime-package-set.ts'

const rootDir = resolve(new URL('..', import.meta.url).pathname)
const packDir = join(rootDir, 'build', 'npm-packages-runtime-smoke')
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
  output: {
    eastPrice: number
    thresholdUnits: number[]
    westUnits: number
  }
  projectDir: string
} {
  mkdirSync(projectDir, { recursive: true })
  writeFileSync(
    join(projectDir, 'package.json'),
    `${JSON.stringify(
      {
        name: 'workpaper-node-consumer',
        private: true,
        type: 'module',
      },
      null,
      2,
    )}\n`,
  )
  writeFileSync(
    join(projectDir, 'index.mjs'),
    [
      'import { ValueTag } from "@bilig/protocol";',
      'import { WorkPaper } from "@bilig/headless";',
      '',
      'const workbook = WorkPaper.buildFromSheets({',
      '  Inputs: [["Region", "Units", "Price"], ["West", 12, 4], ["East", 8, 5], ["West", 3, 6]],',
      '  Summary: [',
      '    ["West Units", \'=SUMIF(Inputs!A2:A4,"West",Inputs!B2:B4)\'],',
      '    ["East Price", \'=XLOOKUP("East",Inputs!A2:A4,Inputs!C2:C4)\'],',
      '    ["Threshold Units", "=FILTER(Inputs!B2:B4,Inputs!B2:B4>5)"],',
      '  ],',
      '});',
      '',
      'const summaryId = workbook.getSheetId("Summary");',
      'if (summaryId === undefined) {',
      '  throw new Error("Summary sheet is missing");',
      '}',
      '',
      'const westUnits = workbook.getCellValue({ sheet: summaryId, row: 0, col: 1 });',
      'const eastPrice = workbook.getCellValue({ sheet: summaryId, row: 1, col: 1 });',
      'const thresholdUnits = workbook.getRangeValues({',
      '  start: { sheet: summaryId, row: 2, col: 1 },',
      '  end: { sheet: summaryId, row: 3, col: 1 },',
      '});',
      '',
      'if (westUnits.tag !== ValueTag.Number || westUnits.value !== 15) {',
      '  throw new Error(`Expected west units to equal 15, received ${JSON.stringify(westUnits)}`);',
      '}',
      'if (eastPrice.tag !== ValueTag.Number || eastPrice.value !== 5) {',
      '  throw new Error(`Expected east price to equal 5, received ${JSON.stringify(eastPrice)}`);',
      '}',
      '',
      'const thresholdUnitValues = thresholdUnits.flat().map((value) => {',
      '  if (value.tag !== ValueTag.Number) {',
      '    throw new Error(`Expected filtered spill numbers, received ${JSON.stringify(value)}`);',
      '  }',
      '  return value.value;',
      '});',
      '',
      'const summary = {',
      '  westUnits: westUnits.value,',
      '  eastPrice: eastPrice.value,',
      '  thresholdUnits: thresholdUnitValues,',
      '};',
      '',
      'if (JSON.stringify(summary.thresholdUnits) !== JSON.stringify([12, 8])) {',
      '  throw new Error(`Unexpected filtered spill: ${JSON.stringify(summary)}`);',
      '}',
      '',
      'console.log(JSON.stringify(summary));',
      '',
    ].join('\n'),
  )

  installTarballs(projectDir, tarballPaths)
  const output = parseNodeSmokeOutput(runTextCommand('node', ['index.mjs'], { cwd: projectDir }))

  return {
    projectDir,
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

  installTarballs(projectDir, tarballPaths, ['vite@8.0.3', 'typescript@6.0.2', '@types/node@25.5.0'])
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
  eastPrice: number
  thresholdUnits: number[]
  westUnits: number
} {
  const parsed = parseJsonRecord(output, 'node smoke output')
  const westUnits = parsed.westUnits
  const eastPrice = parsed.eastPrice
  const thresholdUnits = parsed.thresholdUnits

  if (typeof westUnits !== 'number' || typeof eastPrice !== 'number' || !isNumberArray(thresholdUnits)) {
    throw new Error(`Unexpected node smoke output: ${output}`)
  }

  return {
    westUnits,
    eastPrice,
    thresholdUnits,
  }
}

function parseJsonRecord(serialized: string, context: string): Record<string, unknown> {
  const parsed: unknown = JSON.parse(serialized)
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
