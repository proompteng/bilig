import { chromium } from '@playwright/test'
import type { Buffer } from 'node:buffer'
import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..')
const outputDir = join(repoRoot, 'docs', 'assets')
const checkMode = process.argv.includes('--check')

const galleryWidth = 1270
const galleryHeight = 760
const thumbnailSize = 240

interface AssetSpec {
  readonly fileName: string
  readonly width: number
  readonly height: number
  readonly html: string
}

interface GallerySpec {
  readonly fileName: string
  readonly eyebrow: string
  readonly title: string
  readonly subtitle: string
  readonly accent: string
  readonly visual: string
}

const commonCss = String.raw`
  * {
    box-sizing: border-box;
  }

  body {
    margin: 0;
    background: #eef4f7;
    color: #101820;
    font-family: Inter, Arial, Helvetica, sans-serif;
  }

  .gallery-frame {
    width: ${galleryWidth}px;
    height: ${galleryHeight}px;
    overflow: hidden;
    padding: 52px;
    background:
      linear-gradient(rgba(15, 23, 42, 0.055) 1px, transparent 1px),
      linear-gradient(90deg, rgba(15, 23, 42, 0.055) 1px, transparent 1px),
      #eef4f7;
    background-size: 34px 34px;
  }

  .gallery-shell {
    display: grid;
    grid-template-columns: 0.82fr 1.18fr;
    gap: 34px;
    width: 100%;
    height: 100%;
    border: 1px solid #c8d5de;
    border-radius: 18px;
    background: #ffffff;
    box-shadow: 0 24px 70px rgba(15, 23, 42, 0.13);
    padding: 38px;
  }

  .copy {
    display: flex;
    min-width: 0;
    flex-direction: column;
    justify-content: space-between;
  }

  .brand {
    display: flex;
    align-items: center;
    gap: 16px;
    color: #2f6f47;
    font-size: 21px;
    font-weight: 780;
  }

  .mark {
    display: grid;
    width: 50px;
    height: 50px;
    grid-template-columns: repeat(2, 1fr);
    grid-template-rows: repeat(2, 1fr);
    gap: 7px;
    border: 1px solid #a9bdc9;
    border-radius: 12px;
    background: #101820;
    padding: 10px;
  }

  .mark span {
    border-radius: 3px;
    background: #77c98b;
  }

  .eyebrow {
    margin: 44px 0 0;
    color: var(--accent);
    font-size: 22px;
    font-weight: 820;
    letter-spacing: 0;
    text-transform: uppercase;
  }

  h1 {
    margin: 20px 0 0;
    color: #111820;
    font-size: 58px;
    font-weight: 820;
    letter-spacing: 0;
    line-height: 1.02;
  }

  .subtitle {
    margin: 24px 0 0;
    color: #526273;
    font-size: 26px;
    font-weight: 600;
    letter-spacing: 0;
    line-height: 1.28;
  }

  .footer {
    color: #526273;
    font-family: "SFMono-Regular", Menlo, Consolas, monospace;
    font-size: 21px;
    letter-spacing: 0;
  }

  .visual {
    display: grid;
    min-width: 0;
    align-content: center;
    gap: 18px;
  }

  .panel {
    overflow: hidden;
    border: 1px solid #cbd8e2;
    border-radius: 14px;
    background: #f8fbfd;
  }

  .panel-title {
    display: flex;
    min-height: 46px;
    align-items: center;
    justify-content: space-between;
    border-bottom: 1px solid #d7e1ea;
    background: #ffffff;
    padding: 0 18px;
    color: #506173;
    font-size: 15px;
    font-weight: 800;
    text-transform: uppercase;
  }

  .grid {
    width: 100%;
    border-collapse: collapse;
    font-size: 17px;
  }

  .grid th,
  .grid td {
    height: 39px;
    border-right: 1px solid #d7e1ea;
    border-bottom: 1px solid #d7e1ea;
    padding: 0 14px;
    text-align: left;
  }

  .grid th {
    background: #edf4f8;
    color: #617486;
    font-size: 13px;
    font-weight: 820;
    text-transform: uppercase;
  }

  .grid td {
    color: #253241;
    font-weight: 660;
  }

  .grid .number {
    color: #207146;
    font-family: "SFMono-Regular", Menlo, Consolas, monospace;
    font-weight: 760;
  }

  .terminal {
    background: #111820;
    color: #dce9f2;
    font-family: "SFMono-Regular", Menlo, Consolas, monospace;
  }

  .terminal-body {
    padding: 18px 22px 20px;
    font-size: 18px;
    line-height: 1.38;
  }

  .prompt,
  .ok {
    color: #7ed894;
  }

  .muted {
    color: #95a7b7;
  }

  .cards {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 14px;
  }

  .card {
    min-height: 170px;
    border: 1px solid #cbd8e2;
    border-radius: 14px;
    background: #ffffff;
    padding: 18px;
  }

  .card strong {
    display: block;
    color: #111820;
    font-size: 24px;
    line-height: 1.08;
  }

  .card p {
    margin: 16px 0 0;
    color: #526273;
    font-size: 18px;
    font-weight: 610;
    line-height: 1.25;
  }

  .pill-row {
    display: flex;
    flex-wrap: wrap;
    gap: 10px;
  }

  .pill {
    display: inline-flex;
    min-height: 35px;
    align-items: center;
    border: 1px solid #bfd0dc;
    border-radius: 999px;
    background: #f7fafc;
    padding: 7px 12px;
    color: #263647;
    font-size: 16px;
    font-weight: 760;
  }

  .metric-row {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 14px;
  }

  .metric {
    border: 1px solid #cbd8e2;
    border-radius: 14px;
    background: #ffffff;
    padding: 18px;
  }

  .metric span {
    display: block;
    color: #607184;
    font-size: 15px;
    font-weight: 800;
    text-transform: uppercase;
  }

  .metric strong {
    display: block;
    margin-top: 10px;
    color: #111820;
    font-size: 35px;
    line-height: 1;
  }
`

function galleryHtml(spec: GallerySpec): string {
  return String.raw`<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <style>
      ${commonCss}
    </style>
  </head>
  <body>
    <main class="gallery-frame" style="--accent:${spec.accent}">
      <section class="gallery-shell">
        <div class="copy">
          <div>
            <div class="brand">
              <div class="mark" aria-hidden="true">
                <span></span>
                <span></span>
                <span></span>
                <span></span>
              </div>
              <span>@bilig/headless</span>
            </div>
            <div class="eyebrow">${spec.eyebrow}</div>
            <h1>${spec.title}</h1>
            <p class="subtitle">${spec.subtitle}</p>
          </div>
          <div class="footer">github.com/proompteng/bilig</div>
        </div>
        <div class="visual">
          ${spec.visual}
        </div>
      </section>
    </main>
  </body>
</html>`
}

const gallerySpecs: GallerySpec[] = [
  {
    fileName: 'product-hunt-gallery-01-workbook-api.png',
    accent: '#2f7a4e',
    eyebrow: 'workbook api',
    title: 'spreadsheet logic without screen scraping',
    subtitle: 'create sheets, formulas, structural edits, persistence, and exact readback from node.',
    visual: String.raw`
      <section class="panel">
        <div class="panel-title">
          <span>workpaper after agent edit</span>
          <span>computed</span>
        </div>
        <table class="grid">
          <thead>
            <tr>
              <th>metric</th>
              <th>value</th>
              <th>formula</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td>visitors</td>
              <td class="number">650</td>
              <td>input</td>
            </tr>
            <tr>
              <td>conversion</td>
              <td class="number">10%</td>
              <td>input</td>
            </tr>
            <tr>
              <td>customers</td>
              <td class="number">65</td>
              <td>=B2*B3</td>
            </tr>
            <tr>
              <td>annual arr</td>
              <td class="number">224640</td>
              <td>=B4*240*1.2*12</td>
            </tr>
          </tbody>
        </table>
      </section>
      <section class="panel terminal">
        <div class="panel-title">
          <span>node proof</span>
          <span>readback</span>
        </div>
        <div class="terminal-body">
          <div><span class="prompt">&gt;</span> npm run agent:verify</div>
          <div class="ok">ok writeback readback verified</div>
          <div class="ok">ok formulas persisted after restore</div>
        </div>
      </section>
    `,
  },
  {
    fileName: 'product-hunt-gallery-02-agent-readback.png',
    accent: '#286f9b',
    eyebrow: 'agent workflow',
    title: 'agents need evidence after every workbook edit',
    subtitle: 'bilig turns workbook mutations into logged values, formulas, restored documents, and fixtures.',
    visual: String.raw`
      <div class="cards">
        <div class="card">
          <strong>1. write</strong>
          <p>set cell contents, formulas, and structural changes through the engine.</p>
        </div>
        <div class="card">
          <strong>2. recalc</strong>
          <p>read computed values from workbook state, not from a rendered screenshot.</p>
        </div>
        <div class="card">
          <strong>3. restore</strong>
          <p>serialize the document, reload it, and verify formulas still produce the same answer.</p>
        </div>
      </div>
      <section class="panel terminal">
        <div class="panel-title">
          <span>verification log</span>
          <span>deterministic</span>
        </div>
        <div class="terminal-body">
          <div><span class="muted">before:</span> arr = 172,800</div>
          <div><span class="muted">after:</span> arr = 224,640</div>
          <div class="ok">ok formula cell preserved: =B4*240*1.2*12</div>
          <div class="ok">ok restored workbook matches computed output</div>
        </div>
      </section>
    `,
  },
  {
    fileName: 'product-hunt-gallery-03-node-service.png',
    accent: '#a65f21',
    eyebrow: 'node services',
    title: 'embed formula-backed summaries behind an api',
    subtitle: 'turn csv-shaped inputs into workbook documents your service can compute, persist, and inspect.',
    visual: String.raw`
      <section class="panel terminal">
        <div class="panel-title">
          <span>http handler</span>
          <span>@bilig/headless</span>
        </div>
        <div class="terminal-body">
          <div><span class="prompt">const</span> workbook = WorkPaper.buildFromSheets(data)</div>
          <div><span class="prompt">const</span> total = workbook.getCellValue(summaryCell)</div>
          <div><span class="prompt">return</span> { total, formulas, document }</div>
        </div>
      </section>
      <div class="metric-row">
        <div class="metric">
          <span>formula cells</span>
          <strong>readable</strong>
        </div>
        <div class="metric">
          <span>restore</span>
          <strong>stable</strong>
        </div>
        <div class="metric">
          <span>runtime</span>
          <strong>node</strong>
        </div>
      </div>
      <div class="pill-row">
        <span class="pill">formulas</span>
        <span class="pill">xlsx caveats documented</span>
        <span class="pill">benchmark evidence</span>
        <span class="pill">starter issues</span>
      </div>
    `,
  },
]

function thumbnailHtml(): string {
  return String.raw`<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <style>
      * {
        box-sizing: border-box;
      }

      body {
        display: grid;
        width: ${thumbnailSize}px;
        height: ${thumbnailSize}px;
        place-items: center;
        margin: 0;
        background: #f2f7f4;
        font-family: Inter, Arial, Helvetica, sans-serif;
      }

      .thumb {
        display: grid;
        width: 188px;
        height: 188px;
        grid-template-columns: repeat(2, 1fr);
        grid-template-rows: repeat(2, 1fr);
        gap: 22px;
        border: 2px solid #b7c8d1;
        border-radius: 44px;
        background: #101820;
        box-shadow: 0 22px 48px rgba(15, 23, 42, 0.18);
        padding: 38px;
      }

      .thumb span {
        border-radius: 10px;
        background: #7ed894;
      }
    </style>
  </head>
  <body>
    <div class="thumb" aria-label="bilig icon">
      <span></span>
      <span></span>
      <span></span>
      <span></span>
    </div>
  </body>
</html>`
}

const assets: AssetSpec[] = [
  ...gallerySpecs.map((spec) => ({
    fileName: spec.fileName,
    width: galleryWidth,
    height: galleryHeight,
    html: galleryHtml(spec),
  })),
  {
    fileName: 'product-hunt-thumbnail.png',
    width: thumbnailSize,
    height: thumbnailSize,
    html: thumbnailHtml(),
  },
]

async function renderAsset(asset: AssetSpec): Promise<Buffer> {
  const browser = await chromium.launch()
  try {
    const page = await browser.newPage({
      viewport: { width: asset.width, height: asset.height },
      deviceScaleFactor: 1,
      colorScheme: 'light',
    })

    await page.setContent(asset.html, { waitUntil: 'load' })
    return await page.screenshot({
      type: 'png',
      clip: {
        x: 0,
        y: 0,
        width: asset.width,
        height: asset.height,
      },
    })
  } finally {
    await browser.close()
  }
}

await mkdir(outputDir, { recursive: true })

await Promise.all(
  assets.map(async (asset) => {
    const outputPath = join(outputDir, asset.fileName)
    const rendered = await renderAsset(asset)

    if (checkMode) {
      const existing = await readFile(outputPath)
      if (!existing.equals(rendered)) {
        throw new Error(`${outputPath} is stale. Run pnpm docs:launch-assets:generate.`)
      }
      return
    }

    await writeFile(outputPath, rendered)
    console.log(`wrote ${outputPath}`)
  }),
)

if (checkMode) {
  console.log(`launch assets are current: ${outputDir}`)
}
