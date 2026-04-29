#!/usr/bin/env bun

import { gzipSync } from 'node:zlib'
import { readdir, readFile } from 'node:fs/promises'
import { basename, resolve } from 'node:path'

const budgets = {
  mainJsGzipBytes: 350 * 1024,
  workerJsGzipBytes: 420 * 1024,
  runtimeWasmGzipBytes: 250 * 1024,
  sqliteWasmGzipBytes: 400 * 1024,
  cssGzipBytes: 32 * 1024,
  startupFontGzipBytes: 500 * 1024,
  startupShellGzipBytes: 520 * 1024,
  startupFontFileCount: 40,
}

async function findAssets() {
  const distDir = resolve('apps/web/dist')
  const assetDir = resolve('apps/web/dist/assets')
  const entries = await readdir(assetDir, { withFileTypes: true })
  const files = entries.filter((entry) => entry.isFile()).map((entry) => resolve(assetDir, entry.name))

  const jsAssets = files.filter((file) => file.endsWith('.js'))
  const cssAssets = files.filter((file) => file.endsWith('.css'))
  const wasmAssets = files.filter((file) => file.endsWith('.wasm'))

  if (jsAssets.length === 0) {
    throw new Error('No built JavaScript assets were found in apps/web/dist/assets')
  }

  if (cssAssets.length === 0) {
    throw new Error('No built CSS assets were found in apps/web/dist/assets')
  }

  if (wasmAssets.length === 0) {
    throw new Error('No built WASM assets were found in apps/web/dist/assets')
  }

  return {
    cssAssets,
    distDir,
    indexHtml: resolve(distDir, 'index.html'),
    jsAssets,
    wasmAssets,
  }
}

async function measureAsset(file) {
  const bytes = new Uint8Array(await Bun.file(file).arrayBuffer())
  return {
    file,
    rawBytes: bytes.byteLength,
    gzipBytes: gzipSync(bytes).byteLength,
  }
}

function assertBudget(label, actual, budget) {
  if (actual > budget) {
    throw new Error(`${label} exceeded budget: ${actual} bytes > ${budget} bytes`)
  }
}

function normalizeAssetReference(reference) {
  return reference.trim().replace(/^['"]|['"]$/g, '')
}

function parseStartupAssetReferences(indexHtml) {
  const stylesheetRefs = [...indexHtml.matchAll(/<link[^>]+rel=["']stylesheet["'][^>]+href=["']([^"']+)["']/g)].map((match) =>
    normalizeAssetReference(match[1] ?? ''),
  )
  const modulePreloadRefs = [...indexHtml.matchAll(/<link[^>]+rel=["']modulepreload["'][^>]+href=["']([^"']+)["']/g)].map((match) =>
    normalizeAssetReference(match[1] ?? ''),
  )
  const moduleScriptRefs = [...indexHtml.matchAll(/<script[^>]+type=["']module["'][^>]+src=["']([^"']+)["']/g)].map((match) =>
    normalizeAssetReference(match[1] ?? ''),
  )
  return {
    stylesheetRefs,
    startupScriptRefs: [...new Set([...modulePreloadRefs, ...moduleScriptRefs])],
  }
}

function resolveBuiltAsset(distDir, reference) {
  if (!reference || reference.startsWith('http://') || reference.startsWith('https://')) {
    return null
  }
  if (reference.startsWith('data:')) {
    return null
  }
  const normalized = reference.startsWith('/') ? reference.slice(1) : reference
  return resolve(distDir, normalized)
}

function parseCssFontReferences(css) {
  return [...css.matchAll(/url\(([^)]+)\)/g)]
    .map((match) => normalizeAssetReference(match[1] ?? ''))
    .filter((reference) => /\.(woff2?|ttf|otf)$/i.test(reference))
}

function isSqliteWasmAsset(file) {
  return basename(file).startsWith('sqlite3-')
}

function isWorkerJsAsset(file) {
  const name = basename(file)
  return name.includes('.worker-') || name.startsWith('sqlite3-worker')
}

const { cssAssets, distDir, indexHtml, jsAssets, wasmAssets } = await findAssets()
const jsMeasurements = await Promise.all(jsAssets.map((file) => measureAsset(file)))
const cssMeasurements = await Promise.all(cssAssets.map((file) => measureAsset(file)))
const wasmMeasurements = await Promise.all(wasmAssets.map((file) => measureAsset(file)))
const mainJsMeasurements = jsMeasurements.filter((entry) => !isWorkerJsAsset(entry.file))
const workerJsMeasurements = jsMeasurements.filter((entry) => isWorkerJsAsset(entry.file))

if (mainJsMeasurements.length === 0) {
  throw new Error('No main JavaScript assets were found in apps/web/dist/assets')
}

if (workerJsMeasurements.length === 0) {
  throw new Error('No worker JavaScript assets were found in apps/web/dist/assets')
}

const largestJs = mainJsMeasurements.reduce((largest, entry) => (entry.gzipBytes > largest.gzipBytes ? entry : largest))
const largestWorkerJs = workerJsMeasurements.reduce((largest, entry) => (entry.gzipBytes > largest.gzipBytes ? entry : largest))
const largestCss = cssMeasurements.reduce((largest, entry) => (entry.gzipBytes > largest.gzipBytes ? entry : largest))
const runtimeWasmMeasurements = wasmMeasurements.filter((entry) => !isSqliteWasmAsset(entry.file))
const sqliteWasmMeasurements = wasmMeasurements.filter((entry) => isSqliteWasmAsset(entry.file))

if (runtimeWasmMeasurements.length === 0) {
  throw new Error('No first-party runtime WASM assets were found in apps/web/dist/assets')
}

if (sqliteWasmMeasurements.length === 0) {
  throw new Error('No SQLite WASM assets were found in apps/web/dist/assets')
}

const largestRuntimeWasm = runtimeWasmMeasurements.reduce((largest, entry) => (entry.gzipBytes > largest.gzipBytes ? entry : largest))
const largestSqliteWasm = sqliteWasmMeasurements.reduce((largest, entry) => (entry.gzipBytes > largest.gzipBytes ? entry : largest))
const indexHtmlBytes = new Uint8Array(await readFile(indexHtml))
const indexHtmlMeasurement = {
  file: indexHtml,
  rawBytes: indexHtmlBytes.byteLength,
  gzipBytes: gzipSync(indexHtmlBytes).byteLength,
}
const startupRefs = parseStartupAssetReferences(await Bun.file(indexHtml).text())
const startupCssFiles = [...new Set(startupRefs.stylesheetRefs)]
  .map((reference) => resolveBuiltAsset(distDir, reference))
  .flatMap((file) => (file ? [file] : []))
const startupScriptFiles = [...new Set(startupRefs.startupScriptRefs)]
  .map((reference) => resolveBuiltAsset(distDir, reference))
  .flatMap((file) => (file ? [file] : []))
const startupFontFiles = new Set(
  (
    await Promise.all(
      startupCssFiles.map(async (cssFile) => {
        const css = await Bun.file(cssFile).text()
        return parseCssFontReferences(css)
          .map((reference) => resolveBuiltAsset(distDir, reference))
          .flatMap((file) => (file ? [file] : []))
      }),
    )
  ).flat(),
)

const [missingScriptFile, missingCssFile, missingFontFile] = await Promise.all([
  findMissingStartupAsset(startupScriptFiles),
  findMissingStartupAsset(startupCssFiles),
  findMissingStartupAsset([...startupFontFiles]),
])

if (missingScriptFile) {
  throw new Error(`Startup script asset was not found: ${missingScriptFile}`)
}

if (missingCssFile) {
  throw new Error(`Startup stylesheet asset was not found: ${missingCssFile}`)
}

if (missingFontFile) {
  throw new Error(`Startup font asset was not found: ${missingFontFile}`)
}

const startupCssMeasurements = await Promise.all(startupCssFiles.map((file) => measureAsset(file)))
const startupScriptMeasurements = await Promise.all(startupScriptFiles.map((file) => measureAsset(file)))
const startupFontMeasurements = await Promise.all([...startupFontFiles].map((file) => measureAsset(file)))
const startupFontGzipBytes = startupFontMeasurements.reduce((sum, entry) => sum + entry.gzipBytes, 0)
const startupShellGzipBytes =
  indexHtmlMeasurement.gzipBytes +
  startupCssMeasurements.reduce((sum, entry) => sum + entry.gzipBytes, 0) +
  startupScriptMeasurements.reduce((sum, entry) => sum + entry.gzipBytes, 0) +
  startupFontGzipBytes

assertBudget('Main JavaScript gzip size', largestJs.gzipBytes, budgets.mainJsGzipBytes)
assertBudget('Worker JavaScript gzip size', largestWorkerJs.gzipBytes, budgets.workerJsGzipBytes)
assertBudget('Runtime WASM gzip size', largestRuntimeWasm.gzipBytes, budgets.runtimeWasmGzipBytes)
assertBudget('SQLite WASM gzip size', largestSqliteWasm.gzipBytes, budgets.sqliteWasmGzipBytes)
assertBudget('Largest CSS gzip size', largestCss.gzipBytes, budgets.cssGzipBytes)
assertBudget('Startup font gzip size', startupFontGzipBytes, budgets.startupFontGzipBytes)
assertBudget('Startup shell gzip size', startupShellGzipBytes, budgets.startupShellGzipBytes)
assertBudget('Startup font file count', startupFontMeasurements.length, budgets.startupFontFileCount)

console.log(
  JSON.stringify(
    {
      budgets,
      indexHtmlMeasurement,
      largestJs,
      largestWorkerJs,
      largestCss,
      largestRuntimeWasm,
      largestSqliteWasm,
      startupCssMeasurements,
      startupFontMeasurements,
      startupFontFileCount: startupFontMeasurements.length,
      startupScriptMeasurements,
      startupShellGzipBytes,
    },
    null,
    2,
  ),
)

async function findMissingStartupAsset(files: readonly string[]): Promise<string | null> {
  const existResults = await Promise.all(
    files.map(async (file) => ({
      file,
      exists: await Bun.file(file).exists(),
    })),
  )
  return existResults.find((result) => !result.exists)?.file ?? null
}
