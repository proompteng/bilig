import type { CSSProperties } from 'react'
import type { GridTextScene } from './gridTextScene.js'

interface GridTextOverlayProps {
  readonly active: boolean
  readonly scene: GridTextScene
}

export function GridTextOverlay({ active, scene }: GridTextOverlayProps) {
  if (!active) {
    return null
  }

  return (
    <div aria-hidden="true" className="pointer-events-none absolute inset-0 z-20 overflow-hidden" data-testid="grid-text-overlay">
      {scene.items.map((item) => (
        <div
          className="absolute box-border overflow-hidden px-2"
          key={`${item.x}-${item.y}-${item.width}-${item.height}-${item.text}`}
          style={getTextItemStyle(item)}
        >
          <span
            className="block w-full"
            style={{
              clipPath: `inset(${item.clipInsetTop}px ${item.clipInsetRight}px ${item.clipInsetBottom}px ${item.clipInsetLeft}px)`,
              lineHeight: 1.2,
              overflow: 'hidden',
              textDecorationLine: [item.underline ? 'underline' : null, item.strike ? 'line-through' : null].filter(Boolean).join(' '),
              textOverflow: item.wrap ? undefined : 'clip',
              whiteSpace: item.wrap ? 'pre-wrap' : 'pre',
              wordBreak: item.wrap ? 'break-word' : 'normal',
            }}
          >
            {item.text}
          </span>
        </div>
      ))}
    </div>
  )
}

function getTextItemStyle(item: GridTextScene['items'][number]): CSSProperties {
  return {
    alignItems: item.wrap ? 'flex-start' : 'center',
    color: item.color,
    display: 'flex',
    font: item.font,
    fontSize: item.fontSize,
    height: item.height,
    justifyContent: item.align === 'right' ? 'flex-end' : item.align === 'center' ? 'center' : 'flex-start',
    left: item.x,
    paddingTop: item.wrap ? 4 : 0,
    textAlign: item.align,
    top: item.y,
    width: item.width,
  }
}
