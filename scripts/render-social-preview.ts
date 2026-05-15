import type { Buffer } from 'node:buffer'
import { execFile } from 'node:child_process'
import { mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const previewWidth = 1280
const previewHeight = 640
const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..')
const assetsRoot = join(repoRoot, 'docs', 'assets')
const outputPath = join(assetsRoot, 'github-social-preview.png')
const svgOutputPath = join(assetsRoot, 'github-social-preview.svg')
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

function requirePngDimensions(image: Buffer, context: string): void {
  const expectedSignature = '89504e470d0a1a0a'
  const signature = image.subarray(0, 8).toString('hex')
  if (signature !== expectedSignature) {
    throw new Error(`${context} is not a PNG image`)
  }

  const width = image.readUInt32BE(16)
  const height = image.readUInt32BE(20)
  if (width !== previewWidth || height !== previewHeight) {
    throw new Error(`${context} must be ${previewWidth}x${previewHeight}; got ${width}x${height}`)
  }
}

async function buildSvg(): Promise<string> {
  const sansRegular = await readFile(join(assetsRoot, 'fonts', 'ibm-plex-sans-400.woff2'))
  const sansMedium = await readFile(join(assetsRoot, 'fonts', 'ibm-plex-sans-500.woff2'))
  const sansSemiBold = await readFile(join(assetsRoot, 'fonts', 'ibm-plex-sans-600.woff2'))
  const sansBold = await readFile(join(assetsRoot, 'fonts', 'ibm-plex-sans-700.woff2'))
  const monoMedium = await readFile(join(assetsRoot, 'fonts', 'ibm-plex-mono-500.woff2'))

  return String.raw`<svg xmlns="http://www.w3.org/2000/svg" width="${previewWidth}" height="${previewHeight}" viewBox="0 0 ${previewWidth} ${previewHeight}">
  <defs>
    <style>
      ${fontFace('Bilig Sans', 400, sansRegular)}
      ${fontFace('Bilig Sans', 500, sansMedium)}
      ${fontFace('Bilig Sans', 600, sansSemiBold)}
      ${fontFace('Bilig Sans', 700, sansBold)}
      ${fontFace('Bilig Mono', 500, monoMedium)}
      text {
        font-family: 'Bilig Sans', Arial, sans-serif;
        letter-spacing: 0;
      }
      .mono {
        font-family: 'Bilig Mono', Menlo, monospace;
      }
    </style>
    <linearGradient id="page" x1="0" x2="1" y1="0" y2="1">
      <stop offset="0" stop-color="#070906"/>
      <stop offset="0.55" stop-color="#10150f"/>
      <stop offset="1" stop-color="#162216"/>
    </linearGradient>
    <radialGradient id="glow" cx="72%" cy="35%" r="55%">
      <stop offset="0" stop-color="#37ef91" stop-opacity="0.24"/>
      <stop offset="0.52" stop-color="#17804e" stop-opacity="0.12"/>
      <stop offset="1" stop-color="#050705" stop-opacity="0"/>
    </radialGradient>
    <linearGradient id="sheet" x1="0" x2="1" y1="0" y2="1">
      <stop offset="0" stop-color="#fcfbf7"/>
      <stop offset="1" stop-color="#eaf4eb"/>
    </linearGradient>
    <linearGradient id="edge" x1="0" x2="1" y1="0" y2="1">
      <stop offset="0" stop-color="#ffffff"/>
      <stop offset="1" stop-color="#9cb4a1"/>
    </linearGradient>
    <linearGradient id="green" x1="0" x2="1">
      <stop offset="0" stop-color="#4ef4a0"/>
      <stop offset="1" stop-color="#1da765"/>
    </linearGradient>
    <filter id="shadow" x="-25%" y="-25%" width="150%" height="150%">
      <feDropShadow dx="0" dy="34" stdDeviation="32" flood-color="#000000" flood-opacity="0.45"/>
    </filter>
    <filter id="chipShadow" x="-40%" y="-50%" width="180%" height="200%">
      <feDropShadow dx="0" dy="18" stdDeviation="18" flood-color="#000000" flood-opacity="0.35"/>
    </filter>
  </defs>

  <rect width="${previewWidth}" height="${previewHeight}" fill="url(#page)"/>
  <rect width="${previewWidth}" height="${previewHeight}" fill="url(#glow)"/>
  <g opacity="0.18" stroke="#6a7164" stroke-width="1">
    <path d="M0 80 H1280 M0 160 H1280 M0 240 H1280 M0 320 H1280 M0 400 H1280 M0 480 H1280 M0 560 H1280"/>
    <path d="M96 0 V640 M192 0 V640 M288 0 V640 M384 0 V640 M480 0 V640 M576 0 V640 M672 0 V640 M768 0 V640 M864 0 V640 M960 0 V640 M1056 0 V640 M1152 0 V640"/>
  </g>

  <g transform="translate(74 70)">
    <rect x="0" y="0" width="44" height="44" rx="11" fill="#eff4eb"/>
    <rect x="11" y="11" width="9" height="9" rx="2" fill="#178553"/>
    <rect x="24" y="11" width="9" height="9" rx="2" fill="#178553"/>
    <rect x="11" y="24" width="9" height="9" rx="2" fill="#178553"/>
    <rect x="24" y="24" width="9" height="9" rx="2" fill="#178553"/>
    <text x="60" y="30" fill="#f4f1e8" font-size="26" font-weight="700">bilig</text>
  </g>

  <text x="74" y="174" fill="#37df88" font-size="20" font-weight="700">@bilig/headless</text>
  <text x="74" y="260" fill="#f4f1e8" font-size="72" font-weight="700">Formulas</text>
  <text x="74" y="334" fill="#f4f1e8" font-size="72" font-weight="700">for TypeScript.</text>
  <text x="78" y="397" fill="#c9d1c2" font-size="30" font-weight="400">Edit cells. Recalculate. Save JSON.</text>

  <g transform="translate(74 452)" filter="url(#chipShadow)">
    <rect x="0" y="0" width="438" height="62" rx="0" fill="#151914" stroke="#5d675a"/>
    <rect x="0" y="0" width="64" height="62" fill="#1b211b" stroke="#5d675a"/>
    <text x="27" y="40" fill="#39e98f" class="mono" font-size="23" font-weight="500">$</text>
    <text x="86" y="40" fill="#f4f1e8" class="mono" font-size="23" font-weight="500">npm i @bilig/headless</text>
  </g>

  <g transform="translate(74 558)" fill="#c9d1c2">
    <text x="0" y="0" fill="#f4f1e8" font-size="20" font-weight="700">Node.js</text>
    <path d="M92 -23 V28" stroke="#5d675a"/>
    <text x="118" y="0" fill="#f4f1e8" font-size="20" font-weight="700">MCP</text>
    <path d="M188 -23 V28" stroke="#5d675a"/>
    <text x="214" y="0" fill="#f4f1e8" font-size="20" font-weight="700">XLSX import/export</text>
  </g>

  <g transform="translate(650 86) rotate(-7 282 232)" filter="url(#shadow)">
    <path d="M44 42 H540 L590 430 H88 Z" fill="#0c1710" opacity="0.78"/>
    <path d="M0 0 H502 L552 362 H48 Z" fill="url(#edge)" stroke="#cee0cf" stroke-width="2"/>
    <path d="M36 40 H486 L520 314 H70 Z" fill="url(#sheet)" stroke="#d8e4d7" stroke-width="2"/>
    <path d="M36 84 H492" stroke="#d6e2d6" stroke-width="2"/>
    <path d="M92 40 L118 314 M204 40 L220 314 M316 40 L322 314 M428 40 L424 314" stroke="#d6e2d6" stroke-width="1.4"/>
    <path d="M54 140 H502 M62 194 H510 M70 248 H518" stroke="#d6e2d6" stroke-width="1.4"/>
    <path d="M207 141 H316 L322 193 H211 Z" fill="#b7f2cf"/>
    <path d="M398 249 H518 L526 314 H406 Z" fill="#b7f2cf"/>
    <text x="105" y="71" fill="#7b897a" font-size="19" font-weight="700">A</text>
    <text x="218" y="71" fill="#7b897a" font-size="19" font-weight="700">B</text>
    <text x="331" y="71" fill="#7b897a" font-size="19" font-weight="700">C</text>
    <text x="445" y="71" fill="#7b897a" font-size="19" font-weight="700">D</text>
    <text x="92" y="122" fill="#1d241e" font-size="23">Region</text>
    <text x="208" y="122" fill="#1d241e" font-size="21">Customers</text>
    <text x="348" y="122" fill="#1d241e" font-size="21">ARPA</text>
    <text x="434" y="122" fill="#1d241e" font-size="21">Revenue</text>
    <text x="96" y="175" fill="#1d241e" font-size="23">West</text>
    <text x="254" y="175" fill="#15804f" font-size="24" font-weight="700">32</text>
    <text x="349" y="175" fill="#1d241e" font-size="23">1200</text>
    <text x="441" y="175" fill="#15804f" font-size="24" font-weight="700">38,400</text>
    <text x="96" y="229" fill="#1d241e" font-size="23">East</text>
    <text x="255" y="229" fill="#1d241e" font-size="23">30</text>
    <text x="359" y="229" fill="#1d241e" font-size="23">250</text>
    <text x="455" y="229" fill="#1d241e" font-size="23">7,500</text>
    <text x="96" y="283" fill="#1d241e" font-size="23" font-weight="700">Total</text>
    <text x="450" y="288" fill="#15804f" font-size="26" font-weight="700">51,300</text>
    <path d="M260 154 C340 152 372 199 420 277" fill="none" stroke="url(#green)" stroke-width="10" stroke-linecap="round"/>
    <circle cx="260" cy="154" r="17" fill="url(#green)"/>
    <circle cx="420" cy="277" r="18" fill="url(#green)"/>
    <rect x="54" y="-54" width="392" height="46" rx="8" fill="#f8fbf6" stroke="#d8e4d7"/>
    <rect x="74" y="-42" width="70" height="24" rx="5" fill="#eef5ed" stroke="#d8e4d7"/>
    <text x="91" y="-24" class="mono" fill="#697667" font-size="14" font-weight="500">D5</text>
    <text x="168" y="-24" class="mono" fill="#12824f" font-size="20" font-weight="500">=SUM(D2:D4)</text>
  </g>

  <g transform="translate(1016 444)" filter="url(#chipShadow)">
    <rect x="0" y="0" width="190" height="92" rx="12" fill="#0b1711" stroke="#284936"/>
    <text x="22" y="35" fill="#c9d1c2" font-size="19">after restore</text>
    <text x="22" y="72" fill="#39e98f" font-size="37" font-weight="700">51,300</text>
  </g>

  <text x="1206" y="594" text-anchor="end" fill="#aeb9aa" font-size="18">github.com/proompteng/bilig</text>
</svg>`
}

async function renderPreview(svg: string): Promise<Buffer> {
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

const svg = await buildSvg()
const image = await renderPreview(svg)
requirePngDimensions(image, 'rendered social preview')

if (checkMode) {
  const existingSvg = await readFile(svgOutputPath, 'utf8')
  if (existingSvg !== svg) {
    throw new Error(`${svgOutputPath} is stale. Run pnpm docs:social-preview:generate.`)
  }

  const existingImage = await readFile(outputPath)
  requirePngDimensions(existingImage, 'committed social preview')
  console.log(`social preview source and dimensions are current: ${outputPath}`)
} else {
  await mkdir(dirname(outputPath), { recursive: true })
  await writeFile(svgOutputPath, svg)
  await writeFile(outputPath, image)
  console.log(`wrote ${outputPath}`)
}
