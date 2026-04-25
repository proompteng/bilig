import type { FontKey } from './gridTextPacket.js'
import type { GridTextMetricsProvider } from './gridTextMetrics.js'

export function wrapTextByWidth(input: {
  readonly text: string
  readonly fontKey: FontKey
  readonly maxWidth: number
  readonly metrics: GridTextMetricsProvider
}): readonly string[] {
  if (input.text.length === 0) {
    return []
  }
  if (input.maxWidth <= 0) {
    return ['']
  }
  const words = segmentWords(input.text)
  const lines: string[] = []
  let current = ''
  for (const word of words) {
    const candidate = current.length === 0 ? word : `${current}${word}`
    if (input.metrics.measure(candidate, input.fontKey).advance <= input.maxWidth) {
      current = candidate
      continue
    }
    if (current.length > 0) {
      lines.push(current.trimEnd())
      current = ''
    }
    if (input.metrics.measure(word, input.fontKey).advance <= input.maxWidth) {
      current = word.trimStart()
      continue
    }
    const broken = breakGraphemesByWidth(word.trim(), input.fontKey, input.maxWidth, input.metrics)
    lines.push(...broken.slice(0, -1))
    current = broken.at(-1) ?? ''
  }
  if (current.length > 0) {
    lines.push(current.trimEnd())
  }
  return lines
}

export function segmentGraphemes(text: string): readonly string[] {
  const segmenter = createSegmenter('grapheme')
  if (segmenter) {
    return [...segmenter.segment(text)].map((part) => part.segment)
  }
  return Array.from(text)
}

function segmentWords(text: string): readonly string[] {
  const segmenter = createSegmenter('word')
  if (segmenter) {
    return [...segmenter.segment(text)].map((part) => part.segment)
  }
  return text.match(/\S+\s*/gu) ?? [text]
}

function createSegmenter(granularity: 'grapheme' | 'word'): Intl.Segmenter | null {
  return typeof Intl !== 'undefined' && 'Segmenter' in Intl ? new Intl.Segmenter(undefined, { granularity }) : null
}

function breakGraphemesByWidth(text: string, fontKey: FontKey, maxWidth: number, metrics: GridTextMetricsProvider): readonly string[] {
  const lines: string[] = []
  let current = ''
  for (const glyph of segmentGraphemes(text)) {
    const candidate = `${current}${glyph}`
    if (current.length === 0 || metrics.measure(candidate, fontKey).advance <= maxWidth) {
      current = candidate
      continue
    }
    lines.push(current)
    current = glyph
  }
  if (current.length > 0) {
    lines.push(current)
  }
  return lines
}
