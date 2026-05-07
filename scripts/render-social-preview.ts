import { chromium } from '@playwright/test'
import type { Buffer } from 'node:buffer'
import { mkdir, readFile, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const previewWidth = 1280
const previewHeight = 640
const repoRoot = join(dirname(fileURLToPath(import.meta.url)), '..')
const outputPath = join(repoRoot, 'docs', 'assets', 'github-social-preview.png')
const checkMode = process.argv.includes('--check')

const html = String.raw`<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <style>
      * {
        box-sizing: border-box;
      }

      body {
        margin: 0;
        background: #f4f7f9;
        color: #101720;
        font-family: Inter, Arial, Helvetica, sans-serif;
      }

      .frame {
        position: relative;
        width: ${previewWidth}px;
        height: ${previewHeight}px;
        overflow: hidden;
        padding: 46px 64px;
        background:
          linear-gradient(rgba(24, 34, 45, 0.05) 1px, transparent 1px),
          linear-gradient(90deg, rgba(24, 34, 45, 0.05) 1px, transparent 1px),
          #f4f7f9;
        background-size: 32px 32px;
      }

      .shell {
        display: grid;
        grid-template-columns: 0.88fr 1.12fr;
        gap: 34px;
        height: 100%;
        border: 1px solid #cfd9e2;
        border-radius: 22px;
        background: #ffffff;
        box-shadow: 0 26px 70px rgba(15, 23, 42, 0.13);
        padding: 30px;
      }

      .left {
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        min-width: 0;
        padding: 10px 0 4px;
      }

      .mark-row {
        display: flex;
        align-items: center;
        gap: 18px;
      }

      .mark {
        display: grid;
        width: 60px;
        height: 60px;
        grid-template-columns: repeat(2, 1fr);
        grid-template-rows: repeat(2, 1fr);
        gap: 8px;
        border: 1px solid #b7c7d4;
        border-radius: 14px;
        background: #111820;
        padding: 11px;
      }

      .mark span {
        border-radius: 3px;
        background: #78c889;
      }

      .eyebrow {
        color: #287346;
        font-size: 22px;
        font-weight: 760;
        letter-spacing: 0;
      }

      h1 {
        margin: 34px 0 0;
        color: #111820;
        font-size: 82px;
        font-weight: 790;
        letter-spacing: 0;
        line-height: 0.94;
      }

      .subtitle {
        max-width: 500px;
        margin: 28px 0 0;
        color: #475569;
        font-size: 30px;
        font-weight: 610;
        letter-spacing: 0;
        line-height: 1.22;
      }

      .badges {
        display: flex;
        flex-wrap: wrap;
        gap: 12px;
        margin-top: 30px;
      }

      .badge {
        display: inline-flex;
        min-height: 38px;
        align-items: center;
        border: 1px solid #cbd7e0;
        border-radius: 10px;
        background: #f8fafc;
        padding: 7px 12px;
        color: #253241;
        font-size: 20px;
        font-weight: 700;
      }

      .url {
        color: #526173;
        font-family: "SFMono-Regular", Menlo, Consolas, monospace;
        font-size: 22px;
        letter-spacing: 0;
      }

      .right {
        display: grid;
        grid-template-rows: 310px 140px;
        gap: 18px;
        min-width: 0;
      }

      .panel {
        overflow: hidden;
        border: 1px solid #ccd8e2;
        border-radius: 16px;
        background: #f9fbfd;
      }

      .panel-title {
        display: flex;
        min-height: 44px;
        align-items: center;
        justify-content: space-between;
        border-bottom: 1px solid #d7e1ea;
        padding: 0 18px;
        background: #ffffff;
        color: #516173;
        font-size: 15px;
        font-weight: 780;
      }

      .grid {
        width: 100%;
        border-collapse: collapse;
        font-size: 16px;
      }

      .grid th,
      .grid td {
        height: 35px;
        border-right: 1px solid #d7e1ea;
        border-bottom: 1px solid #d7e1ea;
        padding: 0 14px;
        text-align: left;
        vertical-align: middle;
      }

      .grid th {
        background: #eef4f8;
        color: #607286;
        font-size: 14px;
        font-weight: 800;
        text-transform: uppercase;
      }

      .grid td:first-child {
        width: 42%;
        color: #334155;
        font-weight: 700;
      }

      .grid td:nth-child(2) {
        width: 20%;
        color: #1a7242;
        font-family: "SFMono-Regular", Menlo, Consolas, monospace;
      }

      .grid td:nth-child(3) {
        color: #1f2b3a;
        font-family: "SFMono-Regular", Menlo, Consolas, monospace;
      }

      .result {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 20px;
        padding: 14px 20px;
        background: #102033;
        color: #f8fafc;
      }

      .result small {
        display: block;
        color: #98aec1;
        font-size: 17px;
        font-weight: 730;
      }

      .result strong {
        display: block;
        margin-top: 1px;
        font-size: 30px;
        line-height: 1.05;
      }

      .verified {
        border-radius: 999px;
        background: #e5f7eb;
        padding: 8px 12px;
        color: #126a3b;
        font-size: 16px;
        font-weight: 800;
        white-space: nowrap;
      }

      .terminal {
        display: grid;
        grid-template-rows: 44px 1fr;
        background: #111820;
      }

      .dots {
        display: flex;
        align-items: center;
        gap: 8px;
      }

      .dot {
        width: 12px;
        height: 12px;
        border-radius: 999px;
      }

      .dot.red {
        background: #e96357;
      }

      .dot.yellow {
        background: #e3a728;
      }

      .dot.green {
        background: #71bf75;
      }

      .terminal-body {
        padding: 12px 20px;
        color: #d8e6f1;
        font-family: "SFMono-Regular", Menlo, Consolas, monospace;
        font-size: 18px;
        line-height: 1.3;
      }

      .prompt {
        color: #78c889;
      }

      .muted {
        color: #8fa1b3;
      }

      .ok {
        color: #79d596;
      }
    </style>
  </head>
  <body>
    <main class="frame">
      <section class="shell" aria-label="bilig social preview">
        <div class="left">
          <div>
            <div class="mark-row">
              <div class="mark" aria-hidden="true">
                <span></span>
                <span></span>
                <span></span>
                <span></span>
              </div>
              <div class="eyebrow">npm i @bilig/headless</div>
            </div>
            <h1>bilig</h1>
            <p class="subtitle">spreadsheet formulas without a spreadsheet ui</p>
            <div class="badges" aria-label="package badges">
              <span class="badge">formula readback</span>
              <span class="badge">snapshot restore</span>
              <span class="badge">node services</span>
            </div>
          </div>
          <div class="url">github.com/proompteng/bilig</div>
        </div>

        <div class="right">
          <section class="panel" aria-label="workpaper grid">
            <div class="panel-title">
              <span>write data -> verify formulas</span>
              <span>after agent edit</span>
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
                  <td>650</td>
                  <td>input</td>
                </tr>
                <tr>
                  <td>conversion</td>
                  <td>10%</td>
                  <td>input</td>
                </tr>
                <tr>
                  <td>customers</td>
                  <td>65</td>
                  <td>=B2*B3</td>
                </tr>
                <tr>
                  <td>annual arr</td>
                  <td>224640</td>
                  <td>=B4*240*1.2*12</td>
                </tr>
              </tbody>
            </table>
            <div class="result">
              <div>
                <small>restored formula result</small>
                <strong>$224,640 ARR</strong>
              </div>
              <span class="verified">formulas persisted</span>
            </div>
          </section>

          <section class="panel terminal" aria-label="terminal proof">
            <div class="panel-title">
              <div class="dots" aria-hidden="true">
                <span class="dot red"></span>
                <span class="dot yellow"></span>
                <span class="dot green"></span>
              </div>
              <span>node proof</span>
            </div>
            <div class="terminal-body">
              <div><span class="prompt">&gt;</span> node eval.mjs</div>
              <div class="ok">ok recalculated after edit</div>
              <div class="ok">ok persisted after restore</div>
            </div>
          </section>
        </div>
      </section>
    </main>
  </body>
</html>`

async function renderPreview(): Promise<Buffer> {
  const browser = await chromium.launch()
  try {
    const page = await browser.newPage({
      viewport: { width: previewWidth, height: previewHeight },
      deviceScaleFactor: 1,
      colorScheme: 'light',
    })

    await page.setContent(html, { waitUntil: 'load' })
    return await page.screenshot({
      type: 'png',
      clip: {
        x: 0,
        y: 0,
        width: previewWidth,
        height: previewHeight,
      },
    })
  } finally {
    await browser.close()
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
