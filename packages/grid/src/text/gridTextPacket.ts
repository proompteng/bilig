import type { Rectangle } from '../gridTypes.js'
import type { GridGpuColor } from '../gridGpuScene.js'

export const CELL_TEXT_PADDING_X = 8
export const CELL_TEXT_PADDING_Y = 3

export type TextHorizontalAlign = 'left' | 'center' | 'right'
export type TextVerticalAlign = 'top' | 'middle' | 'bottom'
export type TextOverflowMode = 'clip' | 'overflow'

export interface FontKey {
  readonly family: string
  readonly sizeCssPx: number
  readonly weight: number | string
  readonly style: 'normal' | 'italic'
  readonly stretch?: string | undefined
  readonly variant?: string | undefined
  readonly dprBucket: number
  readonly fontEpoch: number
}

export interface ResolvedGlyphPlacement {
  readonly glyphId: number
  readonly glyph: string
  readonly atlasGlyphKey: string
  readonly worldX: number
  readonly worldY: number
  readonly width: number
  readonly height: number
  readonly advance: number
  readonly uvPadding: number
}

export interface ResolvedTextLine {
  readonly text: string
  readonly worldX: number
  readonly baselineWorldY: number
  readonly advance: number
  readonly ascent: number
  readonly descent: number
  readonly glyphs: readonly ResolvedGlyphPlacement[]
}

export interface TextDecorationRect {
  readonly x: number
  readonly y: number
  readonly width: number
  readonly height: number
  readonly color: GridGpuColor
}

export interface ResolvedCellTextLayout {
  readonly cell: {
    readonly col: number
    readonly row: number
  }
  readonly text: string
  readonly displayText: string
  readonly fontKey: FontKey
  readonly color: GridGpuColor
  readonly horizontalAlign: TextHorizontalAlign
  readonly verticalAlign: TextVerticalAlign
  readonly wrap: boolean
  readonly overflow: TextOverflowMode
  readonly cellWorldRect: Rectangle
  readonly textClipWorldRect: Rectangle
  readonly overflowWorldRect: Rectangle
  readonly lines: readonly ResolvedTextLine[]
  readonly decorations: readonly TextDecorationRect[]
  readonly generation: number
}
