import { describe, expect, it } from 'vitest'
import { getPaneFrame, resolvePaneLayout } from '../renderer-v2/pane-layout.js'

describe('pane-layout', () => {
  it('returns body top left and corner frames for frozen panes', () => {
    const layout = resolvePaneLayout({
      hostWidth: 960,
      hostHeight: 640,
      rowMarkerWidth: 46,
      headerHeight: 24,
      frozenColumnWidth: 208,
      frozenRowHeight: 44,
    })

    expect(layout.body.frame).toEqual({ x: 254, y: 68, width: 706, height: 572 })
    expect(layout.top.frame).toEqual({ x: 254, y: 24, width: 706, height: 44 })
    expect(layout.left.frame).toEqual({ x: 46, y: 68, width: 208, height: 572 })
    expect(layout.corner.frame).toEqual({ x: 46, y: 24, width: 208, height: 44 })
  })

  it('returns a pane frame by pane id', () => {
    const layout = resolvePaneLayout({
      hostWidth: 960,
      hostHeight: 640,
      rowMarkerWidth: 46,
      headerHeight: 24,
      frozenColumnWidth: 208,
      frozenRowHeight: 44,
    })

    expect(getPaneFrame(layout, 'body')).toEqual(layout.body.frame)
    expect(getPaneFrame(layout, 'top')).toEqual(layout.top.frame)
    expect(getPaneFrame(layout, 'left')).toEqual(layout.left.frame)
    expect(getPaneFrame(layout, 'corner')).toEqual(layout.corner.frame)
  })
})
