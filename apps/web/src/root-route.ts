export const ISOLATED_WORKBOOK_PANE_RENDERER_PATH = '/debug/workbook-pane-renderer'

export type WebEntryRoute = 'app' | 'models' | 'isolated-workbook-pane-renderer'

export function resolveWebEntryRoute(pathname: string, search = ''): WebEntryRoute {
  const normalizedPathname = normalizePathname(pathname)
  if (normalizedPathname === ISOLATED_WORKBOOK_PANE_RENDERER_PATH) {
    return 'isolated-workbook-pane-renderer'
  }
  if (normalizedPathname === '/models' || (normalizedPathname === '/' && !search)) return 'models'
  return 'app'
}

function normalizePathname(pathname: string): string {
  if (pathname.length > 1 && pathname.endsWith('/')) {
    return pathname.slice(0, -1)
  }
  return pathname
}
