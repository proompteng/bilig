import type { GridGeometrySnapshot } from './gridGeometry.js'
import type { Rectangle } from './gridTypes.js'

export function applyEditorOverlayBounds(bounds: Rectangle): void {
  if (typeof document === 'undefined') {
    return
  }
  const element = document.querySelector('[data-testid="cell-editor-overlay"]')
  if (!(element instanceof HTMLElement)) {
    return
  }
  element.style.height = `${bounds.height}px`
  element.style.left = `${bounds.x}px`
  element.style.top = `${bounds.y}px`
  element.style.width = `${bounds.width}px`
}

export function resolveEditorOverlayScreenBounds(input: {
  readonly col: number
  readonly row: number
  readonly geometry: GridGeometrySnapshot | null
  readonly hostElement: HTMLElement | null
  readonly getCellLocalBounds: (col: number, row: number) => Rectangle | undefined
}): Rectangle | null {
  const localBounds = input.geometry?.editorScreenRect(input.col, input.row) ?? input.getCellLocalBounds(input.col, input.row)
  const hostBounds = input.hostElement?.getBoundingClientRect()
  if (!localBounds || !hostBounds) {
    return null
  }
  return {
    height: localBounds.height,
    width: localBounds.width,
    x: hostBounds.left + localBounds.x,
    y: hostBounds.top + localBounds.y,
  }
}
