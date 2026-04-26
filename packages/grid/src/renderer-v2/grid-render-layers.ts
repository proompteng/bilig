export const GRID_RENDER_LAYERS = Object.freeze([
  'sheet-background',
  'cell-fills',
  'grid-lines',
  'semantic-fills',
  'selection-fill',
  'body-text',
  'text-decorations',
  'cell-borders',
  'active-cell-border',
  'fill-handle',
  'resize-guides',
  'frozen-pane-separators',
  'editor-and-accessibility-overlays',
] as const)

export type GridRenderLayer = (typeof GRID_RENDER_LAYERS)[number]

export function resolveGridRenderLayerIndex(layer: GridRenderLayer): number {
  return GRID_RENDER_LAYERS.indexOf(layer)
}
