import { describe, expect, it } from 'vitest'
import { collectDeleteOps, collectMountOps, collectSheetOrderOps, normalizeCommitOps } from '../commit-log.js'
import type { CellDescriptor, SheetDescriptor, WorkbookDescriptor } from '../descriptors.js'

function cell(props: CellDescriptor['props']): CellDescriptor {
  return {
    kind: 'Cell',
    props,
    parent: null,
    container: null,
  }
}

function sheet(name: string, children: CellDescriptor[]): SheetDescriptor {
  const descriptor: SheetDescriptor = {
    kind: 'Sheet',
    props: { name },
    children,
    parent: null,
    container: null,
  }
  children.forEach((child) => {
    child.parent = descriptor
  })
  return descriptor
}

function workbook(name: string, children: SheetDescriptor[]): WorkbookDescriptor {
  const descriptor: WorkbookDescriptor = {
    kind: 'Workbook',
    props: { name },
    children,
    parent: null,
    container: null,
  }
  children.forEach((child) => {
    child.parent = descriptor
  })
  return descriptor
}

describe('renderer commit log', () => {
  it('collects mount, delete, and order ops for workbook trees', () => {
    const root = workbook('book', [
      sheet('Sheet1', [cell({ addr: 'A1', value: 10, format: 'currency-usd' }), cell({ addr: 'B1', formula: 'A1*2' })]),
      sheet('Sheet2', [cell({ addr: 'C3', value: true })]),
    ])

    expect(collectMountOps(root)).toEqual([
      { kind: 'upsertWorkbook', name: 'book' },
      { kind: 'upsertSheet', name: 'Sheet1', order: 0 },
      { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'A1', value: 10, format: 'currency-usd' },
      { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'B1', formula: 'A1*2' },
      { kind: 'upsertSheet', name: 'Sheet2', order: 1 },
      { kind: 'upsertCell', sheetName: 'Sheet2', addr: 'C3', value: true },
    ])

    expect(collectDeleteOps(root)).toEqual([
      { kind: 'deleteSheet', name: 'Sheet2' },
      { kind: 'deleteSheet', name: 'Sheet1' },
    ])
    const firstSheet = root.children[0]
    expect(collectDeleteOps(firstSheet)).toEqual([{ kind: 'deleteSheet', name: 'Sheet1' }])
    expect(collectDeleteOps(firstSheet.children[0])).toEqual([{ kind: 'deleteCell', sheetName: 'Sheet1', addr: 'A1' }])
    expect(collectSheetOrderOps(root)).toEqual([
      { kind: 'upsertSheet', name: 'Sheet1', order: 0 },
      { kind: 'upsertSheet', name: 'Sheet2', order: 1 },
    ])
  })

  it('normalizes semantic commit streams without a shadow snapshot diff', () => {
    expect(
      normalizeCommitOps([
        { kind: 'upsertWorkbook', name: 'old-book' },
        { kind: 'upsertWorkbook', name: 'new-book' },
        { kind: 'upsertSheet', name: 'Sheet1', order: 0 },
        { kind: 'deleteCell', sheetName: 'Sheet1', addr: 'A1' },
        { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'A1', value: 42 },
        { kind: 'deleteSheet', name: 'Sheet2' },
      ]),
    ).toEqual([
      { kind: 'upsertWorkbook', name: 'new-book' },
      { kind: 'upsertSheet', name: 'Sheet1', order: 0 },
      { kind: 'upsertCell', sheetName: 'Sheet1', addr: 'A1', value: 42 },
      { kind: 'deleteSheet', name: 'Sheet2' },
    ])
  })

  it('returns no delete ops for orphaned cells', () => {
    expect(collectDeleteOps(cell({ addr: 'A1', value: 1 }))).toEqual([])
  })

  it('returns no sheet order ops without a workbook root', () => {
    expect(collectSheetOrderOps(null)).toEqual([])
  })
})
