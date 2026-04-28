import type { GridTextItem, GridTextScene } from '../gridTextScene.js'
import type { GridRenderTileTextRun } from './render-tile-source.js'

export const GRID_TEXT_METRIC_FLOAT_COUNT_V3 = 8

export interface PackedGridTextBufferV3 {
  readonly textMetrics: Float32Array
  readonly textRuns: readonly GridRenderTileTextRun[]
  readonly textCount: number
  readonly textSignature: string
}

export function packGridTextBufferV3(scene: GridTextScene): PackedGridTextBufferV3 {
  const textRuns = scene.items.map(mapTextRunV3)
  return {
    textCount: scene.items.length,
    textMetrics: packTextMetricsV3(scene.items),
    textRuns,
    textSignature: resolveGridTextSignatureV3(textRuns),
  }
}

export function packTextMetricsV3(items: readonly GridTextItem[]): Float32Array {
  const floats = new Float32Array(Math.max(1, items.length) * GRID_TEXT_METRIC_FLOAT_COUNT_V3)
  items.forEach((item, index) => {
    const offset = index * GRID_TEXT_METRIC_FLOAT_COUNT_V3
    floats[offset + 0] = item.x
    floats[offset + 1] = item.y
    floats[offset + 2] = item.width
    floats[offset + 3] = item.height
    floats[offset + 4] = item.clipInsetTop
    floats[offset + 5] = item.clipInsetRight
    floats[offset + 6] = item.clipInsetBottom
    floats[offset + 7] = item.clipInsetLeft
  })
  return floats
}

export function mapTextRunV3(item: GridTextItem): GridRenderTileTextRun {
  return {
    align: item.align,
    col: item.col,
    clipHeight: Math.max(0, item.height - item.clipInsetTop - item.clipInsetBottom),
    clipWidth: Math.max(0, item.width - item.clipInsetLeft - item.clipInsetRight),
    clipX: item.x + item.clipInsetLeft,
    clipY: item.y + item.clipInsetTop,
    color: item.color,
    font: item.font,
    fontSize: item.fontSize,
    height: item.height,
    row: item.row,
    strike: item.strike,
    text: item.text,
    underline: item.underline,
    width: item.width,
    wrap: item.wrap,
    x: item.x,
    y: item.y,
  }
}

export function resolveGridTextSignatureV3(textRuns: readonly GridRenderTileTextRun[]): string {
  let hash = createHash()
  hash = mixNumber(hash, textRuns.length)
  for (const run of textRuns) {
    hash = mixString(hash, run.text)
    hash = mixNumber(hash, run.x)
    hash = mixNumber(hash, run.y)
    hash = mixNumber(hash, run.width)
    hash = mixNumber(hash, run.height)
    hash = mixNumber(hash, run.clipX)
    hash = mixNumber(hash, run.clipY)
    hash = mixNumber(hash, run.clipWidth)
    hash = mixNumber(hash, run.clipHeight)
    hash = mixString(hash, run.align ?? 'left')
    hash = mixNumber(hash, run.wrap ? 1 : 0)
    hash = mixString(hash, run.font)
    hash = mixNumber(hash, run.fontSize)
    hash = mixString(hash, run.color)
    hash = mixNumber(hash, run.underline ? 1 : 0)
    hash = mixNumber(hash, run.strike ? 1 : 0)
  }
  return hash.toString(36)
}

function createHash(): number {
  return 2_166_136_261
}

function mixString(hash: number, value: string): number {
  let next = hash
  for (let index = 0; index < value.length; index += 1) {
    next = mixInteger(next, value.charCodeAt(index))
  }
  return next
}

function mixNumber(hash: number, value: number): number {
  return mixInteger(hash, Math.round(value * 1_000))
}

function mixInteger(hash: number, value: number): number {
  return Math.imul((hash ^ value) >>> 0, 16_777_619) >>> 0
}
