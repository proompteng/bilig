import type { FontKey } from './gridTextPacket.js'

export interface GridTextMetrics {
  readonly advance: number
  readonly ascent: number
  readonly descent: number
  readonly lineHeight: number
}

export interface GridTextMetricsProvider {
  measure(text: string, fontKey: FontKey): GridTextMetrics
}

export function createFallbackTextMetricsProvider(): GridTextMetricsProvider {
  return {
    measure(text, fontKey) {
      const size = fontKey.sizeCssPx
      const advance = estimateTextAdvance(text, size)
      const ascent = size * 0.8
      const descent = size * 0.2
      return {
        advance,
        ascent,
        descent,
        lineHeight: Math.max(size * 1.2, ascent + descent),
      }
    },
  }
}

export function fontKeyToCssFont(fontKey: FontKey): string {
  const style = fontKey.style === 'italic' ? 'italic ' : ''
  return `${style}${fontKey.weight} ${fontKey.sizeCssPx}px ${fontKey.family}`
}

function estimateTextAdvance(text: string, size: number): number {
  let width = 0
  for (const char of text) {
    if (/\s/u.test(char)) {
      width += size * 0.32
    } else if (/[0-9]/u.test(char)) {
      width += size * 0.56
    } else if (/[A-Z]/u.test(char)) {
      width += size * 0.64
    } else {
      width += size * 0.58
    }
  }
  return width
}
