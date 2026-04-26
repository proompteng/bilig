import React from 'react'
import { SpreadsheetEngine } from '@bilig/core'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { Cell, Sheet, Workbook } from '../components.js'
import type * as RendererRoot from '../renderer-root.js'

type CompatContainer = {
  lastError: unknown
  root?: unknown
}

type CompatMocks = {
  createFiberRoot: ReturnType<typeof vi.fn>
  updateFiberRoot: ReturnType<typeof vi.fn>
}

function isMockContainer(value: unknown): value is {
  root: unknown
} {
  return typeof value === 'object' && value !== null && 'root' in value
}

function DisplayNamedWorkbook({ children }: { name?: string; children?: React.ReactNode }) {
  return <>{children}</>
}
DisplayNamedWorkbook.displayName = 'Workbook'

function DisplayNamedSheet({ children }: { name?: string; children?: React.ReactNode }) {
  return <>{children}</>
}
DisplayNamedSheet.displayName = 'Sheet'

function DisplayNamedCell(_props: { addr?: string; value?: unknown }) {
  return null
}
DisplayNamedCell.displayName = 'Cell'

type UpdateImpl = (element: React.ReactNode, callback: () => void, container: CompatContainer) => void

let capturedContainer: CompatContainer | undefined
let compatUpdateImpl: UpdateImpl = (_element, callback) => {
  callback()
}

const createFiberRootMock = vi.fn((container: CompatContainer) => {
  capturedContainer = container
  return { kind: 'fiber-root' }
})

const updateFiberRootMock = vi.fn((_root: unknown, element: React.ReactNode, callback: () => void) => {
  compatUpdateImpl(element, callback, capturedContainer ?? { lastError: null })
})

vi.mock('../compat.js', () => ({
  createFiberRoot: createFiberRootMock,
  updateFiberRoot: updateFiberRootMock,
}))

async function loadRendererRootWithCompatMock(updateImpl: UpdateImpl): Promise<
  CompatMocks & {
    createWorkbookRendererRoot: typeof RendererRoot.createWorkbookRendererRoot
    getCapturedContainer: () => CompatContainer | undefined
  }
> {
  compatUpdateImpl = updateImpl
  capturedContainer = undefined
  createFiberRootMock.mockClear()
  updateFiberRootMock.mockClear()
  const { createWorkbookRendererRoot } = await import('../renderer-root.js')
  return {
    createWorkbookRendererRoot,
    createFiberRoot: createFiberRootMock,
    updateFiberRoot: updateFiberRootMock,
    getCapturedContainer: () => capturedContainer,
  }
}

afterEach(() => {
  vi.useRealTimers()
  compatUpdateImpl = (_element, callback) => {
    callback()
  }
  capturedContainer = undefined
  createFiberRootMock.mockClear()
  updateFiberRootMock.mockClear()
})

describe('renderer root error handling', () => {
  it('recognizes display-named workbook DSL wrapper functions during validation', async () => {
    const { createWorkbookRendererRoot } = await loadRendererRootWithCompatMock((_element, callback) => {
      callback()
    })
    const engine = new SpreadsheetEngine({ workbookName: 'renderer-root-display-name-dsl' })
    await engine.ready()
    const root = createWorkbookRendererRoot(engine)

    await expect(
      root.render(
        <DisplayNamedWorkbook name="book">
          <DisplayNamedSheet name="Sheet1">
            <DisplayNamedCell addr="A1" value={1} />
          </DisplayNamedSheet>
        </DisplayNamedWorkbook>,
      ),
    ).resolves.toBeUndefined()
  })

  it('ignores duplicate compat callbacks and deletes committed sheets on unmount', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'renderer-root-double-callback' })
    await engine.ready()
    const renderCommit = vi.spyOn(engine, 'renderCommit')
    const { createWorkbookRendererRoot, getCapturedContainer } = await loadRendererRootWithCompatMock((_element, callback) => {
      callback()
      callback()
    })
    const root = createWorkbookRendererRoot(engine)

    await root.render(
      <Workbook name="book">
        <Sheet name="Sheet1">
          <Cell addr="A1" value={1} />
        </Sheet>
      </Workbook>,
    )

    const maybeContainer = getCapturedContainer()
    if (!isMockContainer(maybeContainer)) {
      throw new Error('Expected mocked container')
    }
    const container = maybeContainer
    container.root = {
      kind: 'Workbook',
      parent: null,
      container,
      props: { name: 'book' },
      children: [
        {
          kind: 'Sheet',
          parent: null,
          container,
          props: { name: 'Sheet1' },
          children: [],
        },
      ],
    }

    await root.unmount()

    expect(renderCommit).toHaveBeenCalledWith([{ kind: 'deleteSheet', name: 'Sheet1' }])
  })

  it('unmounts without emitting delete ops when the committed root has no deletable children', async () => {
    const engine = new SpreadsheetEngine({ workbookName: 'renderer-root-empty-unmount' })
    await engine.ready()
    const renderCommit = vi.spyOn(engine, 'renderCommit')
    const { createWorkbookRendererRoot, getCapturedContainer } = await loadRendererRootWithCompatMock((_element, callback) => {
      callback()
    })
    const root = createWorkbookRendererRoot(engine)

    const maybeContainer = getCapturedContainer()
    if (!isMockContainer(maybeContainer)) {
      throw new Error('Expected mocked container')
    }
    const container = maybeContainer
    container.root = {
      kind: 'Workbook',
      parent: null,
      container,
      props: { name: 'book' },
      children: [],
    }

    await root.unmount()

    expect(renderCommit).not.toHaveBeenCalled()
  })

  it('rejects unmount when the compat callback surfaces an async error', async () => {
    const { createWorkbookRendererRoot } = await loadRendererRootWithCompatMock((element, callback, container) => {
      if (element === null) {
        setTimeout(() => {
          container.lastError = new Error('async unmount failed')
        }, 0)
      }
      callback()
    })
    const engine = new SpreadsheetEngine({ workbookName: 'renderer-root-unmount-async-error' })
    await engine.ready()
    const root = createWorkbookRendererRoot(engine)

    await root.render(
      <Workbook name="book">
        <Sheet name="Sheet1">
          <Cell addr="A1" value={1} />
        </Sheet>
      </Workbook>,
    )

    await expect(root.unmount()).rejects.toThrow('async unmount failed')
  })

  it('rejects render when the compat callback surfaces an async error', async () => {
    const { createWorkbookRendererRoot } = await loadRendererRootWithCompatMock((element, callback, container) => {
      callback()
      if (element !== null) {
        setTimeout(() => {
          container.lastError = new Error('async render failed')
        }, 0)
      }
    })
    const engine = new SpreadsheetEngine({ workbookName: 'renderer-root-render-async-error' })
    await engine.ready()
    const root = createWorkbookRendererRoot(engine)

    await expect(
      root.render(
        <Workbook name="book">
          <Sheet name="Sheet1">
            <Cell addr="A1" value={1} />
          </Sheet>
        </Workbook>,
      ),
    ).rejects.toThrow('async render failed')
  })

  it('rejects render and unmount when compat throws synchronously', async () => {
    const syncRenderError = new Error('sync render failed')
    const syncUnmountError = new Error('sync unmount failed')
    const { createWorkbookRendererRoot } = await loadRendererRootWithCompatMock((element) => {
      if (element === null) {
        throw syncUnmountError
      }
      throw syncRenderError
    })
    const engine = new SpreadsheetEngine({ workbookName: 'renderer-root-sync-throws' })
    await engine.ready()
    const root = createWorkbookRendererRoot(engine)

    await expect(
      root.render(
        <Workbook name="book">
          <Sheet name="Sheet1">
            <Cell addr="A1" value={1} />
          </Sheet>
        </Workbook>,
      ),
    ).rejects.toBe(syncRenderError)

    await expect(root.unmount()).rejects.toBe(syncUnmountError)
  })

  it('rejects root text nodes before commit', async () => {
    const { createWorkbookRendererRoot } = await loadRendererRootWithCompatMock(() => {})
    const engine = new SpreadsheetEngine({ workbookName: 'renderer-root-text-node-error' })
    await engine.ready()
    const root = createWorkbookRendererRoot(engine)

    await expect(root.render('bad')).rejects.toThrow('Workbook DSL does not support text nodes.')
  })

  it('rejects invalid workbook DSL shapes before compat commits begin', async () => {
    const { createWorkbookRendererRoot, updateFiberRoot } = await loadRendererRootWithCompatMock(() => {})
    const engine = new SpreadsheetEngine({ workbookName: 'renderer-root-shape-errors' })
    await engine.ready()
    const root = createWorkbookRendererRoot(engine)

    await expect(
      root.render(
        <Sheet name="Sheet1">
          <Cell addr="A1" value={1} />
        </Sheet>,
      ),
    ).rejects.toThrow('Root descriptor must be a Workbook.')

    await expect(
      root.render(
        <Workbook name="book">
          <Cell addr="A1" value={1} />
        </Workbook>,
      ),
    ).rejects.toThrow('Only <Sheet> nodes can exist under <Workbook>.')

    await expect(
      root.render(
        <Workbook name="book">
          <Sheet>
            <Cell addr="A1" value={1} />
          </Sheet>
        </Workbook>,
      ),
    ).rejects.toThrow('<Sheet> requires a name prop.')

    await expect(
      root.render(
        <Workbook name="book">
          <Sheet name="Sheet1">bad</Sheet>
        </Workbook>,
      ),
    ).rejects.toThrow('Workbook DSL does not support text nodes.')

    await expect(
      root.render(
        <Workbook name="book">
          <Sheet name="Sheet1">
            <Sheet name="Nested" />
          </Sheet>
        </Workbook>,
      ),
    ).rejects.toThrow('Only <Cell> can be nested inside <Sheet>.')

    await expect(
      root.render(
        <Workbook name="book">
          <Sheet name="Sheet1">
            <Cell value={1} />
          </Sheet>
        </Workbook>,
      ),
    ).rejects.toThrow('<Cell> requires an addr prop.')

    await expect(
      root.render(
        <Workbook name="book">
          <Sheet name="Sheet1">
            <Cell addr="A1" value={1} formula="A2" />
          </Sheet>
        </Workbook>,
      ),
    ).rejects.toThrow('<Cell> cannot specify both value and formula.')

    expect(updateFiberRoot).not.toHaveBeenCalled()
  })

  it('routes null renders through unmount and surfaces sync compat unmount errors', async () => {
    const { createWorkbookRendererRoot } = await loadRendererRootWithCompatMock((element, callback) => {
      if (element === null) {
        throw new Error('sync compat failure')
      }
      callback()
    })
    const engine = new SpreadsheetEngine({ workbookName: 'renderer-root-sync-error' })
    await engine.ready()
    const root = createWorkbookRendererRoot(engine)

    await expect(root.render(null)).rejects.toThrow('sync compat failure')
  })

  it('treats non-DSL string tags as wrappers during validation', async () => {
    const { createWorkbookRendererRoot } = await loadRendererRootWithCompatMock((_element, callback) => {
      callback()
    })
    const engine = new SpreadsheetEngine({ workbookName: 'renderer-root-dom-wrapper' })
    await engine.ready()
    const root = createWorkbookRendererRoot(engine)

    await expect(
      root.render(
        <Workbook name="book">
          <div>
            <Sheet name="Sheet1">
              <Cell addr="A1" value={1} />
            </Sheet>
          </div>
        </Workbook>,
      ),
    ).resolves.toBeUndefined()
  })

  it('ignores nullish children and rejects workbook text nodes', async () => {
    const { createWorkbookRendererRoot } = await loadRendererRootWithCompatMock((_element, callback) => {
      callback()
    })
    const engine = new SpreadsheetEngine({ workbookName: 'renderer-root-bigint-children' })
    await engine.ready()
    const root = createWorkbookRendererRoot(engine)

    await expect(
      root.render(
        <>
          {null}
          <Workbook name="book">
            {false}
            <Sheet name="Sheet1">
              {undefined}
              <Cell addr="A1" value={1} />
            </Sheet>
          </Workbook>
        </>,
      ),
    ).resolves.toBeUndefined()

    await expect(
      root.render(
        <Workbook name="bad">
          {'bad text'}
          <Sheet name="Sheet1">
            <Cell addr="A1" value={1} />
          </Sheet>
        </Workbook>,
      ),
    ).rejects.toThrow('Workbook DSL does not support text nodes.')
  })
})
