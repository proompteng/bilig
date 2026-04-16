import { describe, expect, it, vi } from 'vitest'
import { SpreadsheetEngine } from '@bilig/core'
import type { CellProps, SheetProps, WorkbookDescriptor, WorkbookProps } from '../descriptors.js'
import { workbookHostConfig, type WorkbookContainer } from '../host-config.js'

function createContainer(engine: SpreadsheetEngine): WorkbookContainer {
  return {
    engine,
    root: null,
    pendingOps: [],
    shouldSyncSheetOrders: false,
    lastError: null,
  }
}

function isWorkbookDescriptor(descriptor: unknown): descriptor is WorkbookDescriptor {
  return typeof descriptor === 'object' && descriptor !== null && 'kind' in descriptor && descriptor.kind === 'Workbook'
}

function expectWorkbookDescriptor(descriptor: unknown): WorkbookDescriptor {
  expect(isWorkbookDescriptor(descriptor)).toBe(true)
  if (!isWorkbookDescriptor(descriptor)) {
    throw new Error('Expected a workbook descriptor')
  }
  return descriptor
}

describe('workbook host config', () => {
  it('exposes stable renderer capabilities and scheduler hooks', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'host-config-basics' })
    await engine.ready()
    const container = createContainer(engine)

    expect(workbookHostConfig.supportsMutation).toBe(true)
    expect(workbookHostConfig.supportsPersistence).toBe(false)
    expect(workbookHostConfig.supportsHydration).toBe(false)
    expect(workbookHostConfig.supportsMicrotasks).toBe(true)
    expect(workbookHostConfig.isPrimaryRenderer).toBe(false)
    expect(workbookHostConfig.getRootHostContext()).toEqual({ kind: 'workbook-root' })
    expect(workbookHostConfig.getChildHostContext()).toEqual({ kind: 'workbook-root' })
    expect(workbookHostConfig.now()).toBeTypeOf('number')
    expect(workbookHostConfig.noTimeout).toBe(-1)
    expect(workbookHostConfig.shouldSetTextContent()).toBe(false)
    expect(workbookHostConfig.getCurrentEventPriority()).toBe(workbookHostConfig.resolveUpdatePriority())
    workbookHostConfig.setCurrentUpdatePriority(42)
    expect(workbookHostConfig.getCurrentUpdatePriority()).toBe(42)
    expect(workbookHostConfig.resolveUpdatePriority()).toBe(42)
    expect(workbookHostConfig.resolveEventType()).toBeNull()
    expect(workbookHostConfig.resolveEventTimeStamp()).toBeTypeOf('number')
    expect(workbookHostConfig.shouldAttemptEagerTransition()).toBe(false)
    expect(workbookHostConfig.maySuspendCommit()).toBe(false)
    expect(workbookHostConfig.maySuspendCommitOnUpdate()).toBe(false)
    expect(workbookHostConfig.maySuspendCommitInSyncRender()).toBe(false)
    expect(workbookHostConfig.waitForCommitToBeReady()).toBeNull()
    expect(workbookHostConfig.NotPendingTransition).toBeNull()
    const workbook = workbookHostConfig.createInstance('Workbook', { name: 'Public' } satisfies WorkbookProps, container)
    expect(workbookHostConfig.getPublicInstance(workbook)).toBe(workbook)

    const logger = workbookHostConfig.bindToConsole()
    expect(logger).toBeTypeOf('function')

    expect(() => workbookHostConfig.createTextInstance()).toThrow('Workbook DSL does not support text nodes.')

    workbookHostConfig.preparePortalMount()
    workbookHostConfig.preloadInstance()
    workbookHostConfig.startSuspendingCommit()
    workbookHostConfig.suspendInstance()
    workbookHostConfig.trackSchedulerEvent()
    workbookHostConfig.detachDeletedInstance()
    workbookHostConfig.commitTextUpdate()
    workbookHostConfig.commitMount()
    workbookHostConfig.resetTextContent()
    workbookHostConfig.hideInstance()
    workbookHostConfig.hideTextInstance()
    workbookHostConfig.unhideInstance()
    workbookHostConfig.unhideTextInstance()
    workbookHostConfig.resetFormInstance()
    workbookHostConfig.scheduleMicrotask(() => {})
    const timeout = workbookHostConfig.scheduleTimeout(() => {}, 0)
    workbookHostConfig.cancelTimeout(timeout)
  })

  it('creates, updates, reorders, and clears descriptors through commit boundaries', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'host-config-ops' })
    await engine.ready()
    const container = createContainer(engine)

    const workbook = workbookHostConfig.createInstance('Workbook', { name: 'Book' } satisfies WorkbookProps, container)
    const sheet = workbookHostConfig.createInstance('Sheet', { name: 'Sheet1' } satisfies SheetProps, container)
    const cell = workbookHostConfig.createInstance('Cell', { addr: 'A1', value: 10 } satisfies CellProps, container)
    const movedCell = workbookHostConfig.createInstance('Cell', { addr: 'B1', formula: 'A1*2' } satisfies CellProps, container)

    workbookHostConfig.appendInitialChild(workbook, sheet)
    workbookHostConfig.appendInitialChild(sheet, cell)
    workbookHostConfig.finalizeInitialChildren()

    workbookHostConfig.prepareForCommit(container)
    workbookHostConfig.appendChildToContainer(container, expectWorkbookDescriptor(workbook))
    workbookHostConfig.resetAfterCommit(container)
    expect(engine.getCellValue('Sheet1', 'A1')).toEqual({ tag: 1, value: 10 })

    workbookHostConfig.prepareForCommit(container)
    workbookHostConfig.appendChild(sheet, movedCell)
    workbookHostConfig.commitUpdate(
      movedCell,
      'Cell',
      { addr: 'B1', formula: 'A1*2' } satisfies CellProps,
      { addr: 'B2', formula: 'A1*3', format: 'currency-usd' } satisfies CellProps,
    )
    workbookHostConfig.commitUpdate(sheet, 'Sheet', { name: 'Sheet1' } satisfies SheetProps, { name: 'Renamed' } satisfies SheetProps)
    workbookHostConfig.resetAfterCommit(container)
    expect(engine.getCell('Renamed', 'B2').format).toBe('currency-usd')
    expect(engine.getCellValue('Renamed', 'B2')).toEqual({ tag: 1, value: 30 })

    const secondSheet = workbookHostConfig.createInstance('Sheet', { name: 'Sheet2' } satisfies SheetProps, container)
    workbookHostConfig.prepareForCommit(container)
    workbookHostConfig.insertBefore(workbook, secondSheet, sheet)
    workbookHostConfig.resetAfterCommit(container)
    expect(engine.exportSnapshot().sheets.map((entry) => `${entry.order}:${entry.name}`)).toEqual(['0:Sheet2', '1:Renamed'])

    workbookHostConfig.prepareForCommit(container)
    workbookHostConfig.removeChild(sheet, cell)
    workbookHostConfig.resetAfterCommit(container)
    expect(engine.getCellValue('Renamed', 'A1')).toEqual({ tag: 0 })

    workbookHostConfig.prepareForCommit(container)
    workbookHostConfig.removeChildFromContainer(container, expectWorkbookDescriptor(workbook))
    workbookHostConfig.resetAfterCommit(container)
    expect(engine.exportSnapshot().sheets).toEqual([])

    workbookHostConfig.clearContainer(container)
    expect(container.root).toBeNull()
  })

  it('surfaces commit-collection errors through the container error slot', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'host-config-errors' })
    await engine.ready()
    const container = createContainer(engine)
    const workbook = workbookHostConfig.createInstance('Workbook', { name: 'BadBook' } satisfies WorkbookProps, container)
    const badSheet = workbookHostConfig.createInstance('Sheet', { name: 'Sheet1' } satisfies SheetProps, container)
    const badCell = workbookHostConfig.createInstance('Cell', { addr: 'A1', value: 1, formula: 'B1' } satisfies CellProps, container)
    workbookHostConfig.appendInitialChild(workbook, badSheet)
    workbookHostConfig.appendInitialChild(badSheet, badCell)

    workbookHostConfig.prepareForCommit(container)
    workbookHostConfig.appendChildToContainer(container, expectWorkbookDescriptor(workbook))
    expect(container.lastError).toBeInstanceOf(Error)

    const consoleError = vi.spyOn(console, 'error').mockImplementation(() => {})
    try {
      workbookHostConfig.resetAfterCommit(container)
      expect(container.lastError?.message).toContain('<Cell> cannot specify both value and formula.')
    } finally {
      consoleError.mockRestore()
    }

    expect(workbookHostConfig.prepareUpdate(workbook, 'Workbook', workbook.props, workbook.props)).toBeNull()
    expect(
      workbookHostConfig.prepareUpdate(workbook, 'Workbook', workbook.props, {
        name: 'Changed',
      } satisfies WorkbookProps),
    ).toBe(true)
    expect(() => workbookHostConfig.createInstance('Unknown', {}, container)).toThrow('Unknown workbook host type')
  })

  it('rejects invalid workbook, sheet, and cell props when creating instances', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'host-config-invalid-props' })
    await engine.ready()
    const container = createContainer(engine)

    expect(() => workbookHostConfig.createInstance('Workbook', { addr: 'A1' } satisfies CellProps, container)).toThrow(
      'Workbook props must not include cell fields.',
    )

    expect(() => workbookHostConfig.createInstance('Sheet', {} satisfies WorkbookProps, container)).toThrow(
      'Sheet props require a sheet name.',
    )

    expect(() => workbookHostConfig.createInstance('Cell', { name: 'Sheet1' } satisfies SheetProps, container)).toThrow(
      'Cell props require an address.',
    )
  })

  it('guards descriptor parenting, container attachment, and prop updates', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'host-config-guards' })
    await engine.ready()
    const container = createContainer(engine)

    const workbook = workbookHostConfig.createInstance('Workbook', { name: 'Book' } satisfies WorkbookProps, container)
    const sheet = workbookHostConfig.createInstance('Sheet', { name: 'Sheet1' } satisfies SheetProps, container)
    const cell = workbookHostConfig.createInstance('Cell', { addr: 'A1', value: 1 } satisfies CellProps, container)
    const anotherCell = workbookHostConfig.createInstance('Cell', { addr: 'B1', value: 2 } satisfies CellProps, container)

    workbookHostConfig.appendInitialChild(workbook, sheet)
    workbookHostConfig.appendInitialChild(sheet, cell)

    expect(() => workbookHostConfig.appendInitialChild(workbook, anotherCell)).toThrow('Cannot append Cell to Workbook.')
    expect(() =>
      workbookHostConfig.appendInitialChild(
        sheet,
        workbookHostConfig.createInstance('Workbook', { name: 'Nested' } satisfies WorkbookProps, container),
      ),
    ).toThrow('Cannot append Workbook to Sheet.')

    workbookHostConfig.appendInitialChild(cell, anotherCell)
    expect(anotherCell.parent).toBe(workbook)

    const floatingSheet = workbookHostConfig.createInstance('Sheet', { name: 'Sheet2' } satisfies SheetProps, container)
    const floatingCell = workbookHostConfig.createInstance('Cell', { addr: 'C1', value: 3 } satisfies CellProps, container)
    workbookHostConfig.insertBefore(workbook, floatingSheet, floatingCell)
    expect(expectWorkbookDescriptor(workbook).children.at(-1)).toBe(floatingSheet)
    workbookHostConfig.insertBefore(sheet, floatingCell, workbook)
    expect(sheet.children.at(-1)).toBe(floatingCell)

    workbookHostConfig.removeChild(workbook, floatingCell)
    expect(expectWorkbookDescriptor(workbook).children.includes(floatingCell)).toBe(false)
    workbookHostConfig.removeChild(sheet, floatingSheet)
    expect(sheet.children.includes(floatingSheet)).toBe(false)

    const orphanContainer = {
      kind: 'Workbook',
      props: { name: 'Orphan' },
      children: [],
      parent: null,
      container: null,
    }
    expect(() =>
      workbookHostConfig.appendChild(orphanContainer, workbookHostConfig.createInstance('Sheet', { name: 'X' }, container)),
    ).toThrow('Descriptor is not attached to a workbook container.')

    expect(() => workbookHostConfig.appendChildToContainer(container, sheet)).toThrow(
      'Only workbook descriptors can be attached to the root container.',
    )
    expect(() => workbookHostConfig.insertInContainerBefore(container, sheet)).toThrow(
      'Only workbook descriptors can be attached to the root container.',
    )
    expect(() => workbookHostConfig.removeChildFromContainer(container, sheet)).toThrow(
      'Only workbook descriptors can be removed from the root container.',
    )

    workbookHostConfig.prepareForCommit(container)
    workbookHostConfig.insertInContainerBefore(container, expectWorkbookDescriptor(workbook))
    workbookHostConfig.resetAfterCommit(container)
    expect(container.root).toBe(workbook)

    expect(() =>
      workbookHostConfig.commitUpdate(workbook, 'Workbook', { addr: 'A1', value: 1 } satisfies CellProps, workbook.props),
    ).toThrow('Workbook updates require workbook props.')
    expect(() => workbookHostConfig.commitUpdate(sheet, 'Sheet', { addr: 'A1', value: 1 } satisfies CellProps, sheet.props)).toThrow(
      'Sheet updates require sheet props.',
    )
    expect(() => workbookHostConfig.commitUpdate(cell, 'Cell', { name: 'Bad' } satisfies SheetProps, cell.props)).toThrow(
      'Cell updates require cell props.',
    )

    const detachedCell = workbookHostConfig.createInstance('Cell', { addr: 'Z1', value: 9 } satisfies CellProps, container)
    workbookHostConfig.commitUpdate(detachedCell, 'Cell', detachedCell.props, {
      addr: 'Z2',
      value: 10,
    } satisfies CellProps)
    expect(container.pendingOps).toEqual([])

    workbookHostConfig.clearContainer(container)
    expect(container.root).toBeNull()
  })

  it('inserts before an existing sibling using the splice path', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'host-config-insert-before' })
    await engine.ready()
    const container = createContainer(engine)

    const workbook = workbookHostConfig.createInstance('Workbook', { name: 'Book' } satisfies WorkbookProps, container)
    const firstSheet = workbookHostConfig.createInstance('Sheet', { name: 'Sheet1' } satisfies SheetProps, container)
    const secondSheet = workbookHostConfig.createInstance('Sheet', { name: 'Sheet2' } satisfies SheetProps, container)
    const insertedSheet = workbookHostConfig.createInstance('Sheet', { name: 'Inserted' } satisfies SheetProps, container)

    workbookHostConfig.appendInitialChild(workbook, firstSheet)
    workbookHostConfig.appendInitialChild(workbook, secondSheet)
    workbookHostConfig.insertBefore(workbook, insertedSheet, secondSheet)

    expect(expectWorkbookDescriptor(workbook).children.map((sheet) => sheet.props.name)).toEqual(['Sheet1', 'Inserted', 'Sheet2'])

    const firstCell = workbookHostConfig.createInstance('Cell', { addr: 'A1', value: 1 } satisfies CellProps, container)
    const secondCell = workbookHostConfig.createInstance('Cell', { addr: 'B1', value: 2 } satisfies CellProps, container)
    const insertedCell = workbookHostConfig.createInstance('Cell', { addr: 'A2', value: 3 } satisfies CellProps, container)

    workbookHostConfig.appendInitialChild(firstSheet, firstCell)
    workbookHostConfig.appendInitialChild(firstSheet, secondCell)
    workbookHostConfig.insertBefore(firstSheet, insertedCell, secondCell)

    expect(firstSheet.children.map((cell) => cell.props.addr)).toEqual(['A1', 'A2', 'B1'])
  })
})
