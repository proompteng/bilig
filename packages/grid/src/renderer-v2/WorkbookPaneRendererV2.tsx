import { memo } from 'react'
import type { GridGeometrySnapshot } from '../gridGeometry.js'

export interface WorkbookPaneRendererV2Props {
  readonly active: boolean
  readonly host: HTMLDivElement | null
  readonly geometry: GridGeometrySnapshot | null
}

export const WorkbookPaneRendererV2 = memo(function WorkbookPaneRendererV2({ active, geometry, host }: WorkbookPaneRendererV2Props) {
  if (!active || !host) {
    return null
  }

  return (
    <canvas
      aria-hidden="true"
      className="pointer-events-none absolute inset-0 z-10"
      data-pane-renderer="workbook-pane-renderer-v2"
      data-renderer-mode="typegpu-v2"
      data-testid="grid-pane-renderer"
      data-v2-body-world-x={geometry?.camera.bodyWorldX ?? 0}
      data-v2-body-world-y={geometry?.camera.bodyWorldY ?? 0}
      style={{ contain: 'strict' }}
    />
  )
})
