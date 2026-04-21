export const WORKBOOK_RENDERER_MODES = ['typegpu-v1', 'typegpu-v2', 'canvas-fallback'] as const

export type WorkbookRendererMode = (typeof WORKBOOK_RENDERER_MODES)[number]

export const DEFAULT_WORKBOOK_RENDERER_MODE: WorkbookRendererMode = 'typegpu-v1'

export function isWorkbookRendererMode(value: string | null | undefined): value is WorkbookRendererMode {
  return value === 'typegpu-v1' || value === 'typegpu-v2' || value === 'canvas-fallback'
}

export function resolveWorkbookRendererMode(
  input: {
    readonly explicit?: string | null | undefined
    readonly env?: string | null | undefined
    readonly search?: string | URLSearchParams | null | undefined
  } = {},
): WorkbookRendererMode {
  if (isWorkbookRendererMode(input.explicit)) {
    return input.explicit
  }

  const searchParams =
    typeof input.search === 'string'
      ? new URLSearchParams(input.search.startsWith('?') ? input.search.slice(1) : input.search)
      : input.search instanceof URLSearchParams
        ? input.search
        : null
  const queryValue = searchParams?.get('workbookRenderer') ?? searchParams?.get('renderer')
  if (isWorkbookRendererMode(queryValue)) {
    return queryValue
  }

  if (isWorkbookRendererMode(input.env)) {
    return input.env
  }

  return DEFAULT_WORKBOOK_RENDERER_MODE
}
