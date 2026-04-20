import { describe, expect, it } from 'vitest'
import { ISOLATED_WORKBOOK_PANE_RENDERER_PATH, resolveWebEntryRoute } from '../root-route.js'

describe('resolveWebEntryRoute', () => {
  it('routes the isolated renderer path to the standalone renderer entry', () => {
    expect(resolveWebEntryRoute(ISOLATED_WORKBOOK_PANE_RENDERER_PATH)).toBe('isolated-workbook-pane-renderer')
    expect(resolveWebEntryRoute(`${ISOLATED_WORKBOOK_PANE_RENDERER_PATH}/`)).toBe('isolated-workbook-pane-renderer')
  })

  it('defaults all other paths to the workbook app entry', () => {
    expect(resolveWebEntryRoute('/')).toBe('app')
    expect(resolveWebEntryRoute('/workbooks/demo')).toBe('app')
  })
})
