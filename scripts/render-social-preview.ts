import { execFile } from 'node:child_process'
import type { Buffer } from 'node:buffer'
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const previewWidth = 1280
const previewHeight = 640
const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..')
const assetsRoot = join(repoRoot, 'docs', 'assets')
const backgroundPath = join(assetsRoot, 'bilig-social-background.png')
const outputPath = join(assetsRoot, 'github-social-preview.png')
const checkMode = process.argv.includes('--check')

function escapeXml(value: string): string {
  return value.replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;')
}

function renderTextLines(lines: readonly string[], x: number, y: number, lineHeight: number): string {
  return lines.map((line, index) => `<tspan x="${x}" dy="${index === 0 ? 0 : lineHeight}">${escapeXml(line)}</tspan>`).join('')
}

function execFileBuffer(file: string, args: readonly string[]): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    execFile(
      file,
      [...args],
      {
        encoding: 'buffer',
        maxBuffer: 16 * 1024 * 1024,
      },
      (error, stdout, stderr) => {
        if (error !== null) {
          const message = Buffer.isBuffer(stderr) ? stderr.toString('utf8') : String(stderr)
          reject(new Error(`${file} failed: ${message.trim() || error.message}`))
          return
        }
        resolve(Buffer.from(stdout))
      },
    )
  })
}

async function buildSvg(): Promise<string> {
  const background = await readFile(backgroundPath)
  const backgroundUri = `data:image/png;base64,${background.toString('base64')}`
  const titleLines = ['spreadsheet formulas', 'for Node.js programs']
  const subtitleLines = ['build, recalc, and save workbooks', 'without opening a browser grid']

  return String.raw`<svg xmlns="http://www.w3.org/2000/svg" width="${previewWidth}" height="${previewHeight}" viewBox="0 0 ${previewWidth} ${previewHeight}">
  <defs>
    <linearGradient id="leftFade" x1="0" x2="1" y1="0" y2="0">
      <stop offset="0" stop-color="#f8fafc" stop-opacity="0.99"/>
      <stop offset="0.72" stop-color="#f8fafc" stop-opacity="0.95"/>
      <stop offset="1" stop-color="#f8fafc" stop-opacity="0"/>
    </linearGradient>
    <linearGradient id="cardShade" x1="0" x2="1" y1="0" y2="1">
      <stop offset="0" stop-color="#ffffff" stop-opacity="0.98"/>
      <stop offset="1" stop-color="#edf4f7" stop-opacity="0.94"/>
    </linearGradient>
    <filter id="softShadow" x="-20%" y="-20%" width="140%" height="140%">
      <feDropShadow dx="0" dy="18" stdDeviation="24" flood-color="#0f172a" flood-opacity="0.18"/>
    </filter>
  </defs>

  <rect width="${previewWidth}" height="${previewHeight}" fill="#f5f8fa"/>
  <image href="${backgroundUri}" x="0" y="0" width="${previewWidth}" height="${previewHeight}" preserveAspectRatio="xMidYMid slice"/>
  <rect width="760" height="${previewHeight}" fill="url(#leftFade)"/>
  <rect x="64" y="48" width="1152" height="544" rx="24" fill="none" stroke="#cbd7e0" stroke-width="1.4"/>

  <g filter="url(#softShadow)">
    <rect x="74" y="72" width="520" height="496" rx="22" fill="url(#cardShade)" stroke="#cbd7e0" stroke-width="1"/>
  </g>

  <g transform="translate(102 102)">
    <rect x="0" y="0" width="58" height="58" rx="14" fill="#111820"/>
    <rect x="13" y="13" width="13" height="13" rx="3" fill="#78d38b"/>
    <rect x="32" y="13" width="13" height="13" rx="3" fill="#78d38b"/>
    <rect x="13" y="32" width="13" height="13" rx="3" fill="#78d38b"/>
    <rect x="32" y="32" width="13" height="13" rx="3" fill="#78d38b"/>
  </g>

  <text x="176" y="126" fill="#176d3f" font-family="Inter, Arial, Helvetica, sans-serif" font-size="25" font-weight="760">npm i @bilig/headless</text>
  <text x="102" y="238" fill="#0f1720" font-family="Inter, Arial, Helvetica, sans-serif" font-size="86" font-weight="820" letter-spacing="0">bilig</text>
  <text x="102" y="296" fill="#263548" font-family="Inter, Arial, Helvetica, sans-serif" font-size="36" font-weight="750" letter-spacing="0">
    ${renderTextLines(titleLines, 102, 296, 43)}
  </text>
  <text x="102" y="396" fill="#4b5d70" font-family="Inter, Arial, Helvetica, sans-serif" font-size="23" font-weight="590" letter-spacing="0">
    ${renderTextLines(subtitleLines, 102, 396, 31)}
  </text>

  <g fill="#f8fbfd" stroke="#cbd7e0" stroke-width="1">
    <rect x="102" y="474" width="164" height="42" rx="10"/>
    <rect x="278" y="474" width="152" height="42" rx="10"/>
    <rect x="442" y="474" width="116" height="42" rx="10"/>
  </g>
  <g fill="#223042" font-family="Inter, Arial, Helvetica, sans-serif" font-size="20" font-weight="740">
    <text x="120" y="501">formula api</text>
    <text x="296" y="501">node 24+</text>
    <text x="460" y="501">mit</text>
  </g>

  <g transform="translate(760 438)">
    <rect x="0" y="0" width="420" height="92" rx="18" fill="#102033" fill-opacity="0.94"/>
    <text x="24" y="35" fill="#93a8bc" font-family="Inter, Arial, Helvetica, sans-serif" font-size="18" font-weight="740">workbook round trip</text>
    <text x="24" y="66" fill="#ffffff" font-family="Inter, Arial, Helvetica, sans-serif" font-size="28" font-weight="800">write &#8594; recalc &#8594; save</text>
    <rect x="334" y="24" width="56" height="44" rx="14" fill="#e6f8eb"/>
    <path d="M350 46 l11 12 l23 -28" fill="none" stroke="#177344" stroke-width="7" stroke-linecap="round" stroke-linejoin="round"/>
  </g>

  <text x="102" y="552" fill="#5b6a7a" font-family="SFMono-Regular, Menlo, Consolas, monospace" font-size="21">github.com/proompteng/bilig</text>
</svg>`
}

async function renderPreview(): Promise<Buffer> {
  const svg = await buildSvg()
  const tempRoot = await mkdtemp(join(tmpdir(), 'bilig-social-preview-'))
  const svgPath = join(tempRoot, 'preview.svg')
  try {
    await writeFile(svgPath, svg)
    return await execFileBuffer('rsvg-convert', [
      '--format=png',
      '--width',
      String(previewWidth),
      '--height',
      String(previewHeight),
      svgPath,
    ])
  } finally {
    await rm(tempRoot, { recursive: true, force: true })
  }
}

const image = await renderPreview()

if (checkMode) {
  const existing = await readFile(outputPath)
  if (!existing.equals(image)) {
    throw new Error(`${outputPath} is stale. Run pnpm docs:social-preview:generate.`)
  }
  console.log(`social preview is current: ${outputPath}`)
} else {
  await mkdir(dirname(outputPath), { recursive: true })
  await writeFile(outputPath, image)
  console.log(`wrote ${outputPath}`)
}
