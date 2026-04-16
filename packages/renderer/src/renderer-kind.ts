export const RENDERER_KIND_PROP = '__biligRendererKind'

export type RendererKind = 'Workbook' | 'Sheet' | 'Cell'

export type RendererComponent = {
  [RENDERER_KIND_PROP]?: RendererKind
}
