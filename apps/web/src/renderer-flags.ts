export const WEB_WORKBOOK_RENDERER_MODES = ['typegpu-v1', 'typegpu-v2', 'canvas-fallback'] as const

export type WebWorkbookRendererMode = (typeof WEB_WORKBOOK_RENDERER_MODES)[number]

function isWebWorkbookRendererMode(value: string | null | undefined): value is WebWorkbookRendererMode {
  return value === 'typegpu-v1' || value === 'typegpu-v2' || value === 'canvas-fallback'
}

export function resolveWebWorkbookRendererMode(search = window.location.search): WebWorkbookRendererMode {
  const searchParams = new URLSearchParams(search.startsWith('?') ? search.slice(1) : search)
  const queryValue = searchParams.get('workbookRenderer') ?? searchParams.get('renderer')
  if (isWebWorkbookRendererMode(queryValue)) {
    return queryValue
  }
  const envValue = import.meta.env['VITE_WORKBOOK_RENDERER']
  return isWebWorkbookRendererMode(envValue) ? envValue : 'typegpu-v1'
}
