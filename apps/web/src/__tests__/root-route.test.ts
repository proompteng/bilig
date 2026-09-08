import { describe, expect, it } from 'vitest'
import { ISOLATED_WORKBOOK_PANE_RENDERER_PATH, resolveWebEntryRoute } from '../root-route.js'

describe('resolveWebEntryRoute', () => {
  it('routes the isolated renderer path to the standalone renderer entry', () => {
    expect(resolveWebEntryRoute(ISOLATED_WORKBOOK_PANE_RENDERER_PATH)).toBe('isolated-workbook-pane-renderer')
    expect(resolveWebEntryRoute(`${ISOLATED_WORKBOOK_PANE_RENDERER_PATH}/`)).toBe('isolated-workbook-pane-renderer')
  })

  it('opens models at the root and preserves workbook and debug links', () => {
    expect(resolveWebEntryRoute('/')).toBe('models')
    expect(resolveWebEntryRoute('/models/')).toBe('models')
    expect(resolveWebEntryRoute('/models', '?model=example')).toBe('models')
    expect(resolveWebEntryRoute('/workbook')).toBe('app')
    expect(resolveWebEntryRoute('/', '?document=example')).toBe('app')
    expect(resolveWebEntryRoute('/', '?persist=0')).toBe('app')
    expect(resolveWebEntryRoute('/workbooks/demo')).toBe('app')
  })
})
