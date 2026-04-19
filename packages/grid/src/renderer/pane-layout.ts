import type { Rectangle } from '../gridTypes.js'
import type { WorkbookPaneId } from './pane-scene-types.js'

export interface WorkbookPaneLayout {
  readonly body: {
    readonly frame: Rectangle
  }
  readonly top: {
    readonly frame: Rectangle
  }
  readonly left: {
    readonly frame: Rectangle
  }
  readonly corner: {
    readonly frame: Rectangle
  }
}

function clampLength(value: number): number {
  return Math.max(0, Math.round(value))
}

function createFrame(x: number, y: number, width: number, height: number): Rectangle {
  return {
    x: clampLength(x),
    y: clampLength(y),
    width: clampLength(width),
    height: clampLength(height),
  }
}

export function resolvePaneLayout(input: {
  hostWidth: number
  hostHeight: number
  rowMarkerWidth: number
  headerHeight: number
  frozenColumnWidth: number
  frozenRowHeight: number
}): WorkbookPaneLayout {
  const bodyX = input.rowMarkerWidth + input.frozenColumnWidth
  const bodyY = input.headerHeight + input.frozenRowHeight
  return {
    body: {
      frame: createFrame(bodyX, bodyY, input.hostWidth - bodyX, input.hostHeight - bodyY),
    },
    top: {
      frame: createFrame(bodyX, input.headerHeight, input.hostWidth - bodyX, input.frozenRowHeight),
    },
    left: {
      frame: createFrame(input.rowMarkerWidth, bodyY, input.frozenColumnWidth, input.hostHeight - bodyY),
    },
    corner: {
      frame: createFrame(input.rowMarkerWidth, input.headerHeight, input.frozenColumnWidth, input.frozenRowHeight),
    },
  }
}

export function getPaneFrame(layout: WorkbookPaneLayout, paneId: WorkbookPaneId): Rectangle {
  switch (paneId) {
    case 'body':
      return layout.body.frame
    case 'top':
      return layout.top.frame
    case 'left':
      return layout.left.frame
    case 'corner':
      return layout.corner.frame
  }
}
