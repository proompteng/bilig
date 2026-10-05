import { describe, expect, it } from 'vitest'
import { ISOLATED_WORKBOOK_PANE_RENDERER_PATH, resolveWebEntryRoute } from '../root-route.js'

describe('resolveWebEntryRoute', () => {
  it('routes the isolated renderer path to the standalone renderer entry', () => {
    expect(resolveWebEntryRoute(ISOLATED_WORKBOOK_PANE_RENDERER_PATH)).toBe('isolated-workbook-pane-renderer')
    expect(resolveWebEntryRoute(`${ISOLATED_WORKBOOK_PANE_RENDERER_PATH}/`)).toBe('isolated-workbook-pane-renderer')
  })

  it('opens the spreadsheet at the root and workbook paths', () => {
    expect(resolveWebEntryRoute('/')).toBe('app')
    expect(resolveWebEntryRoute('/workbook')).toBe('app')
    expect(resolveWebEntryRoute('/workbooks/demo')).toBe('app')
  })
})
