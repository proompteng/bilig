import type { Buffer } from 'node:buffer'
import { execFile } from 'node:child_process'
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const heroWidth = 1600
const heroHeight = 900
const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..')
const assetsRoot = join(repoRoot, 'docs', 'assets')
const outputPath = join(assetsRoot, 'bilig-hero-workbook-api.png')
const svgOutputPath = join(assetsRoot, 'bilig-hero-workbook-api.svg')
const checkMode = process.argv.includes('--check')

function fontFace(family: string, weight: number, data: Buffer): string {
  return String.raw`
    @font-face {
      font-family: '${family}';
      font-weight: ${weight};
      src: url(data:font/woff2;base64,${data.toString('base64')}) format('woff2');
    }`
}

function execFileBuffer(file: string, args: readonly string[]): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    execFile(
      file,
      [...args],
      {
        encoding: 'buffer',
        maxBuffer: 24 * 1024 * 1024,
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

function requirePngDimensions(image: Buffer, context: string): void {
  const pngSignature = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
  if (image.length < 24 || !image.subarray(0, pngSignature.length).equals(pngSignature)) {
    throw new Error(`${context} is not a PNG image`)
  }

  const width = image.readUInt32BE(16)
  const height = image.readUInt32BE(20)
  if (width !== heroWidth || height !== heroHeight) {
    throw new Error(`${context} must be ${heroWidth.toString()}x${heroHeight.toString()}; got ${width.toString()}x${height.toString()}`)
  }
}

async function buildSvg(): Promise<string> {
  const sansRegular = await readFile(join(assetsRoot, 'fonts', 'ibm-plex-sans-400.woff2'))
  const sansMedium = await readFile(join(assetsRoot, 'fonts', 'ibm-plex-sans-500.woff2'))
  const sansSemiBold = await readFile(join(assetsRoot, 'fonts', 'ibm-plex-sans-600.woff2'))
  const sansBold = await readFile(join(assetsRoot, 'fonts', 'ibm-plex-sans-700.woff2'))
  const monoMedium = await readFile(join(assetsRoot, 'fonts', 'ibm-plex-mono-500.woff2'))
  const monoSemiBold = await readFile(join(assetsRoot, 'fonts', 'ibm-plex-mono-600.woff2'))

  return String.raw`<svg xmlns="http://www.w3.org/2000/svg" width="${heroWidth}" height="${heroHeight}" viewBox="0 0 ${heroWidth} ${heroHeight}">
  <defs>
    <style>
      ${fontFace('Bilig Sans', 400, sansRegular)}
      ${fontFace('Bilig Sans', 500, sansMedium)}
      ${fontFace('Bilig Sans', 600, sansSemiBold)}
      ${fontFace('Bilig Sans', 700, sansBold)}
      ${fontFace('Bilig Mono', 500, monoMedium)}
      ${fontFace('Bilig Mono', 600, monoSemiBold)}
      text {
        font-family: 'Bilig Sans', Arial, sans-serif;
        letter-spacing: 0;
      }
      .mono {
        font-family: 'Bilig Mono', Menlo, monospace;
      }
    </style>
    <linearGradient id="page" x1="0" x2="1" y1="0" y2="1">
      <stop offset="0" stop-color="#f7f6ef"/>
      <stop offset="0.55" stop-color="#edf3eb"/>
      <stop offset="1" stop-color="#dbeade"/>
    </linearGradient>
    <linearGradient id="surface" x1="0" x2="1" y1="0" y2="1">
      <stop offset="0" stop-color="#fbfaf4"/>
      <stop offset="1" stop-color="#edf3e8"/>
    </linearGradient>
    <linearGradient id="code" x1="0" x2="1" y1="0" y2="1">
      <stop offset="0" stop-color="#111917"/>
      <stop offset="1" stop-color="#14251d"/>
    </linearGradient>
    <linearGradient id="greenGlow" x1="0" x2="1" y1="0" y2="1">
      <stop offset="0" stop-color="#2ee177" stop-opacity="0.22"/>
      <stop offset="1" stop-color="#2ee177" stop-opacity="0"/>
    </linearGradient>
    <filter id="softShadow" x="-20%" y="-20%" width="140%" height="140%">
      <feDropShadow dx="0" dy="30" stdDeviation="36" flood-color="#0f140f" flood-opacity="0.22"/>
    </filter>
  </defs>

  <rect width="${heroWidth}" height="${heroHeight}" fill="url(#page)"/>
  <rect x="732" y="-80" width="700" height="700" rx="350" fill="url(#greenGlow)"/>
  <g opacity="0.5" stroke="#dce4d8" stroke-width="1">
    <path d="M0 90 H1600"/>
    <path d="M0 180 H1600"/>
    <path d="M0 270 H1600"/>
    <path d="M0 360 H1600"/>
    <path d="M0 450 H1600"/>
    <path d="M0 540 H1600"/>
    <path d="M0 630 H1600"/>
    <path d="M0 720 H1600"/>
    <path d="M0 810 H1600"/>
    <path d="M100 0 V900"/>
    <path d="M200 0 V900"/>
    <path d="M300 0 V900"/>
    <path d="M400 0 V900"/>
    <path d="M500 0 V900"/>
    <path d="M600 0 V900"/>
    <path d="M700 0 V900"/>
    <path d="M800 0 V900"/>
    <path d="M900 0 V900"/>
    <path d="M1000 0 V900"/>
    <path d="M1100 0 V900"/>
    <path d="M1200 0 V900"/>
    <path d="M1300 0 V900"/>
    <path d="M1400 0 V900"/>
    <path d="M1500 0 V900"/>
  </g>

  <g filter="url(#softShadow)">
    <rect x="128" y="86" width="1344" height="708" rx="24" fill="#ffffff" stroke="#b7c1b0"/>
  </g>
  <rect x="128" y="86" width="1344" height="64" rx="24" fill="#f2f5ef"/>
  <path d="M128 150 H1472" stroke="#cfd8c8"/>

  <g transform="translate(170 104)">
    <rect x="0" y="0" width="32" height="32" rx="8" fill="#10140f"/>
    <rect x="8" y="8" width="6" height="6" rx="2" fill="#34d57b"/>
    <rect x="18" y="8" width="6" height="6" rx="2" fill="#34d57b"/>
    <rect x="8" y="18" width="6" height="6" rx="2" fill="#34d57b"/>
    <rect x="18" y="18" width="6" height="6" rx="2" fill="#34d57b"/>
    <text x="46" y="24" fill="#171a15" font-size="24" font-weight="700">WorkPaper</text>
  </g>
  <rect x="1136" y="106" width="236" height="24" rx="12" fill="#dff3e6"/>
  <circle cx="1156" cy="118" r="4" fill="#16814f"/>
  <text x="1170" y="124" fill="#16814f" font-size="15" font-weight="700">D5 updates after B2</text>

  <g transform="translate(170 196)">
    <rect x="0" y="0" width="780" height="456" rx="18" fill="url(#surface)" stroke="#cfd9c9"/>
    <rect x="30" y="30" width="720" height="52" rx="10" fill="#f8f8f2" stroke="#cfd9c9"/>
    <rect x="46" y="42" width="78" height="28" rx="6" fill="#edf3e9" stroke="#cfd9c9"/>
    <text x="85" y="63" text-anchor="middle" fill="#596357" font-size="17" font-weight="700">D5</text>
    <text x="154" y="63" class="mono" fill="#13794b" font-size="18" font-weight="600">=SUM(D2:D4)</text>

    <g transform="translate(30 118)">
      <rect x="0" y="0" width="720" height="300" fill="#fafaf5" stroke="#d5ded0"/>
      <rect x="0" y="0" width="720" height="54" fill="#e9f0e5"/>
      <path d="M178 0 V300 M356 0 V300 M534 0 V300 M0 54 H720 M0 108 H720 M0 162 H720 M0 216 H720 M0 270 H720" stroke="#d5ded0"/>
      <text x="88" y="35" text-anchor="middle" fill="#71806b" font-size="18" font-weight="700">A</text>
      <text x="267" y="35" text-anchor="middle" fill="#71806b" font-size="18" font-weight="700">B</text>
      <text x="445" y="35" text-anchor="middle" fill="#71806b" font-size="18" font-weight="700">C</text>
      <text x="623" y="35" text-anchor="middle" fill="#71806b" font-size="18" font-weight="700">D</text>
      <text x="28" y="89" fill="#242822" font-size="21" font-weight="600">Region</text>
      <text x="230" y="89" fill="#242822" font-size="21" font-weight="600">Units</text>
      <text x="408" y="89" fill="#242822" font-size="21" font-weight="600">ARPA</text>
      <text x="586" y="89" fill="#242822" font-size="21" font-weight="600">Total</text>
      <text x="28" y="143" fill="#242822" font-size="22">West</text>
      <rect x="178" y="108" width="178" height="54" fill="#dff3e6"/>
      <text x="324" y="143" text-anchor="end" fill="#12824d" font-size="24" font-weight="700">32</text>
      <text x="502" y="143" text-anchor="end" fill="#242822" font-size="22">1200</text>
      <text x="686" y="143" text-anchor="end" fill="#12824d" font-size="24" font-weight="700">38,400</text>
      <text x="28" y="197" fill="#242822" font-size="22">East</text>
      <text x="324" y="197" text-anchor="end" fill="#242822" font-size="22">30</text>
      <text x="502" y="197" text-anchor="end" fill="#242822" font-size="22">250</text>
      <text x="686" y="197" text-anchor="end" fill="#242822" font-size="22">7,500</text>
      <text x="28" y="251" fill="#242822" font-size="22">Central</text>
      <text x="324" y="251" text-anchor="end" fill="#242822" font-size="22">18</text>
      <text x="502" y="251" text-anchor="end" fill="#242822" font-size="22">300</text>
      <text x="686" y="251" text-anchor="end" fill="#242822" font-size="22">5,400</text>
      <rect x="534" y="270" width="186" height="30" fill="#dff3e6"/>
      <text x="28" y="291" fill="#242822" font-size="22" font-weight="700">Total</text>
      <text x="686" y="291" text-anchor="end" fill="#12824d" font-size="25" font-weight="700">51,300</text>
    </g>
  </g>

  <g transform="translate(1010 196)">
    <rect x="0" y="0" width="334" height="456" rx="18" fill="url(#code)" stroke="#33473b"/>
    <rect x="0" y="0" width="334" height="58" rx="18" fill="#202a22"/>
    <text x="24" y="36" class="mono" fill="#9bd2ff" font-size="18" font-weight="600">workbook.ts</text>
    <g transform="translate(26 94)" class="mono" font-size="15" font-weight="600">
      <text x="0" y="0" fill="#7dc7ff">set</text>
      <text x="44" y="0" fill="#f4f0e7">B2</text>
      <text x="84" y="0" fill="#6ee29a">32</text>
      <rect x="0" y="28" width="250" height="12" rx="6" fill="#d6e6d3" opacity="0.82"/>
      <rect x="0" y="56" width="190" height="12" rx="6" fill="#7dc7ff" opacity="0.72"/>
      <rect x="0" y="84" width="272" height="12" rx="6" fill="#d6e6d3" opacity="0.82"/>
      <rect x="0" y="112" width="138" height="12" rx="6" fill="#6ee29a" opacity="0.86"/>
      <rect x="0" y="160" width="204" height="12" rx="6" fill="#d6e6d3" opacity="0.72"/>
      <rect x="0" y="188" width="248" height="12" rx="6" fill="#7dc7ff" opacity="0.68"/>
      <rect x="0" y="216" width="156" height="12" rx="6" fill="#d6e6d3" opacity="0.72"/>
      <text x="0" y="296" fill="#7dc7ff">read</text>
      <text x="54" y="296" fill="#f4f0e7">D5</text>
      <text x="98" y="296" fill="#6ee29a">51,300</text>
    </g>
  </g>

  <g transform="translate(170 702)">
    <rect x="0" y="0" width="244" height="58" rx="14" fill="#f6f7f0" stroke="#cfd9c9"/>
    <text x="28" y="38" fill="#5b6459" font-size="17" font-weight="600">36,900</text>
    <path d="M124 29 H164" stroke="#9aa895" stroke-width="2"/>
    <path d="M158 22 L166 29 L158 36" fill="none" stroke="#9aa895" stroke-width="2"/>
    <text x="184" y="38" fill="#12784a" font-size="18" font-weight="700">51,300</text>
    <rect x="276" y="0" width="136" height="58" rx="14" fill="#112018"/>
    <circle cx="304" cy="29" r="7" fill="#34d57b"/>
    <text x="324" y="36" fill="#dff3e6" font-size="18" font-weight="700">JSON ok</text>
  </g>
</svg>`
}

async function renderPng(svg: string): Promise<Buffer> {
  const tempRoot = await mkdtemp(join(tmpdir(), 'bilig-hero-workbook-api-'))
  const svgPath = join(tempRoot, 'hero.svg')

  try {
    await writeFile(svgPath, svg)
    return await execFileBuffer('rsvg-convert', ['--format=png', '--width', String(heroWidth), '--height', String(heroHeight), svgPath])
  } finally {
    await rm(tempRoot, { recursive: true, force: true })
  }
}

const svg = await buildSvg()
const image = await renderPng(svg)
requirePngDimensions(image, 'rendered hero asset')

if (checkMode) {
  const existingSvg = await readFile(svgOutputPath, 'utf8')
  if (existingSvg !== svg) {
    throw new Error(`${svgOutputPath} is stale. Run pnpm docs:hero-asset:generate.`)
  }
  const existingImage = await readFile(outputPath)
  requirePngDimensions(existingImage, outputPath)
  console.log(`hero asset is current: ${outputPath}`)
} else {
  await mkdir(assetsRoot, { recursive: true })
  await writeFile(svgOutputPath, svg)
  await writeFile(outputPath, image)
  console.log(`wrote ${svgOutputPath}`)
  console.log(`wrote ${outputPath}`)
}
