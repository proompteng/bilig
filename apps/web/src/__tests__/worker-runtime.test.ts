import { afterEach, describe, expect, it, vi } from 'vitest'
import { SpreadsheetEngine } from '@bilig/core'
import { formatAddress } from '@bilig/formula'
import type { WorkbookLocalMutationRecord, WorkbookLocalStoreFactory } from '@bilig/storage-browser'
import { WorkbookLocalStoreLockedError, createMemoryWorkbookLocalStoreFactory } from '@bilig/storage-browser'
import { ErrorCode, createCellNumberFormatRecord, formatCellDisplayValue, isWorkbookSnapshot, ValueTag } from '@bilig/protocol'
import {
  decodeRenderTileDeltaBatch,
  decodeViewportPatch,
  decodeWorkbookDeltaBatchV3,
  type RenderTileDeltaBatch,
  type RenderTileReplaceMutation,
} from '@bilig/worker-transport'
import { buildWorkbookLocalAuthoritativeBase } from '../worker-local-base.js'
import { collectChangedCellsBySheet, collectViewportCells } from '../worker-runtime-support.js'
import { WorkbookWorkerRuntime } from '../worker-runtime'

type TestLocalStore = Awaited<ReturnType<WorkbookLocalStoreFactory['open']>>
type TestStoredState = Awaited<ReturnType<TestLocalStore['loadState']>>
type TestStoredStateValue = NonNullable<TestStoredState>
type TestPersistProjectionStateInput = Parameters<TestLocalStore['persistProjectionState']>[0]
type TestIngestAuthoritativeDeltaInput = Parameters<TestLocalStore['ingestAuthoritativeDelta']>[0]
type TestAuthoritativeBase = TestPersistProjectionStateInput['authoritativeBase']
type TestProjectionOverlay = TestPersistProjectionStateInput['projectionOverlay']
type TestAuthoritativeDelta = TestIngestAuthoritativeDeltaInput['authoritativeDelta']

function cloneMutationRecord(mutation: WorkbookLocalMutationRecord): WorkbookLocalMutationRecord {
  const nextMutation = structuredClone(mutation)
  nextMutation.args = [...mutation.args]
  return nextMutation
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null
}

function findRenderTileReplace(batch: RenderTileDeltaBatch, rowTile: number, colTile: number): RenderTileReplaceMutation {
  const mutation = batch.mutations.find(
    (entry) => entry.kind === 'tileReplace' && entry.coord.rowTile === rowTile && entry.coord.colTile === colTile,
  )
  if (!mutation || mutation.kind !== 'tileReplace') {
    throw new Error(`Missing render tile replacement r${rowTile}:c${colTile}`)
  }
  return mutation
}

function buildViewportFromAuthoritativeBase(input: {
  authoritativeBase: NonNullable<TestAuthoritativeBase>
  sheetName: string
  viewport: {
    rowStart: number
    rowEnd: number
    colStart: number
    colEnd: number
  }
}) {
  const { authoritativeBase, sheetName, viewport } = input
  const sheet = authoritativeBase.sheets.find((entry) => entry.name === sheetName)
  if (!sheet) {
    throw new Error(`Missing authoritative sheet ${sheetName}`)
  }
  const styles = authoritativeBase.styles.filter((style) => {
    return (
      style.id === 'style-0' ||
      authoritativeBase.cellRenders.some((cell) => {
        return (
          cell.sheetName === sheetName &&
          cell.styleId === style.id &&
          cell.rowNum >= viewport.rowStart &&
          cell.rowNum <= viewport.rowEnd &&
          cell.colNum >= viewport.colStart &&
          cell.colNum <= viewport.colEnd
        )
      })
    )
  })
  return {
    sheetId: sheet.sheetId,
    sheetName,
    cells: authoritativeBase.cellRenders
      .filter((cell) => {
        return (
          cell.sheetName === sheetName &&
          cell.rowNum >= viewport.rowStart &&
          cell.rowNum <= viewport.rowEnd &&
          cell.colNum >= viewport.colStart &&
          cell.colNum <= viewport.colEnd
        )
      })
      .map((cell) => {
        const inputRecord = authoritativeBase.cellInputs.find(
          (entry) => entry.sheetName === cell.sheetName && entry.address === cell.address,
        )
        return {
          row: cell.rowNum,
          col: cell.colNum,
          snapshot: {
            sheetName: cell.sheetName,
            address: cell.address,
            value: structuredClone(cell.value),
            flags: cell.flags,
            version: cell.version,
            styleId: cell.styleId,
            numberFormatId: cell.numberFormatId,
            input: inputRecord?.input,
            formula: inputRecord?.formula,
            format: inputRecord?.format,
          },
        }
      }),
    rowAxisEntries: authoritativeBase.rowAxisEntries
      .filter((entry) => entry.sheetName === sheetName)
      .map((entry) => structuredClone(entry.entry)),
    columnAxisEntries: authoritativeBase.columnAxisEntries
      .filter((entry) => entry.sheetName === sheetName)
      .map((entry) => structuredClone(entry.entry)),
    styles: structuredClone(styles),
  }
}

function mergeViewportWithProjectionOverlay(input: {
  baseViewport: ReturnType<typeof buildViewportFromAuthoritativeBase>
  projectionOverlay: TestProjectionOverlay
  sheetName: string
  viewport: {
    rowStart: number
    rowEnd: number
    colStart: number
    colEnd: number
  }
}) {
  const { baseViewport, projectionOverlay, sheetName, viewport } = input
  if (!projectionOverlay) {
    return baseViewport
  }

  const cells = new Map(baseViewport.cells.map((cell) => [cell.snapshot.address, cell]))
  projectionOverlay.cells
    .filter((cell) => {
      return (
        cell.sheetName === sheetName &&
        cell.rowNum >= viewport.rowStart &&
        cell.rowNum <= viewport.rowEnd &&
        cell.colNum >= viewport.colStart &&
        cell.colNum <= viewport.colEnd
      )
    })
    .forEach((cell) => {
      cells.set(cell.address, {
        row: cell.rowNum,
        col: cell.colNum,
        snapshot: {
          sheetName: cell.sheetName,
          address: cell.address,
          value: structuredClone(cell.value),
          flags: cell.flags,
          version: cell.version,
          input: cell.input,
          formula: cell.formula,
          format: cell.format,
          styleId: cell.styleId,
          numberFormatId: cell.numberFormatId,
        },
      })
    })

  const rowAxisEntries = new Map(baseViewport.rowAxisEntries.map((entry) => [entry.index, entry]))
  projectionOverlay.rowAxisEntries
    .filter((entry) => entry.sheetName === sheetName)
    .forEach((entry) => {
      rowAxisEntries.set(entry.entry.index, structuredClone(entry.entry))
    })

  const columnAxisEntries = new Map(baseViewport.columnAxisEntries.map((entry) => [entry.index, entry]))
  projectionOverlay.columnAxisEntries
    .filter((entry) => entry.sheetName === sheetName)
    .forEach((entry) => {
      columnAxisEntries.set(entry.entry.index, structuredClone(entry.entry))
    })

  const styles = new Map(baseViewport.styles.map((style) => [style.id, style]))
  projectionOverlay.styles.forEach((style) => {
    styles.set(style.id, structuredClone(style))
  })

  return {
    ...baseViewport,
    cells: [...cells.values()].toSorted((left, right) => left.row - right.row || left.col - right.col),
    rowAxisEntries: [...rowAxisEntries.values()].toSorted((left, right) => left.index - right.index),
    columnAxisEntries: [...columnAxisEntries.values()].toSorted((left, right) => left.index - right.index),
    styles: [...styles.values()],
  }
}

function mergeAuthoritativeBaseDelta(input: {
  currentBase: TestAuthoritativeBase
  authoritativeDelta: TestAuthoritativeDelta
}): NonNullable<TestAuthoritativeBase> {
  const { currentBase, authoritativeDelta } = input
  if (authoritativeDelta.replaceAll || currentBase === null) {
    return structuredClone(authoritativeDelta.base)
  }

  const replacedSheetIds = new Set(authoritativeDelta.replacedSheetIds)
  return {
    sheets: [
      ...currentBase.sheets.filter((sheet) => !replacedSheetIds.has(sheet.sheetId)),
      ...structuredClone(authoritativeDelta.base.sheets),
    ].toSorted((left, right) => left.sortOrder - right.sortOrder),
    cellInputs: [
      ...currentBase.cellInputs.filter((cell) => !replacedSheetIds.has(cell.sheetId)),
      ...structuredClone(authoritativeDelta.base.cellInputs),
    ],
    cellRenders: [
      ...currentBase.cellRenders.filter((cell) => !replacedSheetIds.has(cell.sheetId)),
      ...structuredClone(authoritativeDelta.base.cellRenders),
    ],
    rowAxisEntries: [
      ...currentBase.rowAxisEntries.filter((entry) => !replacedSheetIds.has(entry.sheetId)),
      ...structuredClone(authoritativeDelta.base.rowAxisEntries),
    ],
    columnAxisEntries: [
      ...currentBase.columnAxisEntries.filter((entry) => !replacedSheetIds.has(entry.sheetId)),
      ...structuredClone(authoritativeDelta.base.columnAxisEntries),
    ],
    styles: structuredClone(authoritativeDelta.base.styles),
  }
}

function createMemoryLocalStoreFactory(seed?: {
  state?: TestStoredState
  pendingMutations?: readonly WorkbookLocalMutationRecord[]
  onPersistProjectionState?: (state: TestStoredStateValue) => Promise<void> | void
  onIngestAuthoritativeDelta?: (state: TestStoredStateValue, delta: TestAuthoritativeDelta) => Promise<void> | void
  onReadViewportProjection?: (
    sheetName: string,
    viewport: {
      rowStart: number
      rowEnd: number
      colStart: number
      colEnd: number
    },
  ) => void
  authoritativeBase?: TestAuthoritativeBase
  projectionOverlay?: TestProjectionOverlay
}): WorkbookLocalStoreFactory {
  let currentState = seed?.state ? structuredClone(seed.state) : null
  let currentMutationJournal = (seed?.pendingMutations ?? []).map(cloneMutationRecord)
  let currentAuthoritativeBase = seed?.authoritativeBase ? structuredClone(seed.authoritativeBase) : null
  let currentProjectionOverlay = seed?.projectionOverlay ? structuredClone(seed.projectionOverlay) : null
  return {
    async open() {
      if (currentState && currentAuthoritativeBase === null) {
        const snapshot = currentState.snapshot
        if (isWorkbookSnapshot(snapshot)) {
          const engine = new SpreadsheetEngine({ workbookName: 'derived', replicaId: 'derived' })
          await engine.ready()
          engine.importSnapshot(snapshot)
          currentAuthoritativeBase = buildWorkbookLocalAuthoritativeBase(engine)
          currentProjectionOverlay ??= {
            cells: [],
            rowAxisEntries: [],
            columnAxisEntries: [],
            styles: [],
          }
        }
      }
      return {
        async loadBootstrapState() {
          const sheetNames =
            currentAuthoritativeBase?.sheets.map((sheet) => sheet.name) ??
            (isRecord(currentState?.snapshot) && Array.isArray(currentState.snapshot['sheets'])
              ? currentState.snapshot['sheets'].flatMap((sheet) =>
                  isRecord(sheet) && typeof sheet['name'] === 'string' ? [sheet['name']] : [],
                )
              : [])
          const workbookName =
            isRecord(currentState?.snapshot) &&
            isRecord(currentState.snapshot['workbook']) &&
            typeof currentState.snapshot['workbook']['name'] === 'string'
              ? currentState.snapshot['workbook']['name']
              : 'Sheet1'
          const state = currentState
            ? {
                workbookName,
                sheetNames,
                materializedCellCount:
                  currentAuthoritativeBase?.cellRenders.length ??
                  (isRecord(currentState.snapshot) && Array.isArray(currentState.snapshot['sheets'])
                    ? currentState.snapshot['sheets'].reduce((count, sheet) => {
                        if (!isRecord(sheet) || !Array.isArray(sheet['cells'])) {
                          return count
                        }
                        return count + sheet['cells'].length
                      }, 0)
                    : 0),
                authoritativeRevision: currentState.authoritativeRevision,
                appliedPendingLocalSeq: currentState.appliedPendingLocalSeq,
              }
            : null
          return state ? structuredClone(state) : null
        },
        async loadState() {
          return currentState ? structuredClone(currentState) : null
        },
        async persistProjectionState(input) {
          currentState = structuredClone(input.state)
          currentAuthoritativeBase = structuredClone(input.authoritativeBase)
          currentProjectionOverlay = structuredClone(input.projectionOverlay)
          await seed?.onPersistProjectionState?.(input.state)
        },
        async ingestAuthoritativeDelta(input) {
          currentState = structuredClone(input.state)
          currentAuthoritativeBase = mergeAuthoritativeBaseDelta({
            currentBase: currentAuthoritativeBase,
            authoritativeDelta: input.authoritativeDelta,
          })
          currentProjectionOverlay = structuredClone(input.projectionOverlay)
          if ((input.removePendingMutationIds?.length ?? 0) > 0) {
            const removedIds = new Set(input.removePendingMutationIds)
            currentMutationJournal = currentMutationJournal.map((mutation) =>
              removedIds.has(mutation.id)
                ? {
                    ...cloneMutationRecord(mutation),
                    ackedAtUnixMs: Date.now(),
                    status: 'acked',
                  }
                : mutation,
            )
          }
          await seed?.onIngestAuthoritativeDelta?.(input.state, input.authoritativeDelta)
        },
        async listPendingMutations() {
          return currentMutationJournal.filter((mutation) => mutation.status !== 'acked').map(cloneMutationRecord)
        },
        async listMutationJournalEntries() {
          return currentMutationJournal.map(cloneMutationRecord)
        },
        async appendPendingMutation(mutation) {
          currentMutationJournal.push(cloneMutationRecord(mutation))
        },
        async updatePendingMutation(mutation) {
          currentMutationJournal = currentMutationJournal.map((entry) => (entry.id === mutation.id ? cloneMutationRecord(mutation) : entry))
        },
        async removePendingMutation(id) {
          currentMutationJournal = currentMutationJournal.filter((mutation) => mutation.id !== id)
        },
        readViewportProjection(sheetName, viewport) {
          const authoritativeBase = currentAuthoritativeBase
          if (!authoritativeBase) {
            return null
          }
          seed?.onReadViewportProjection?.(sheetName, viewport)
          return mergeViewportWithProjectionOverlay({
            baseViewport: buildViewportFromAuthoritativeBase({
              authoritativeBase,
              sheetName,
              viewport,
            }),
            projectionOverlay: currentProjectionOverlay,
            sheetName,
            viewport,
          })
        },
        close() {},
      }
    },
  }
}

describe('WorkbookWorkerRuntime', () => {
  afterEach(() => {
    vi.useRealTimers()
  })

  it('restores persisted workbook state and emits viewport patches for visible edits', async () => {
    const seedEngine = new SpreadsheetEngine({ workbookName: 'phase3-doc', replicaId: 'seed' })
    seedEngine.createSheet('Sheet1')
    seedEngine.setCellValue('Sheet1', 'A1', 7)

    const localStoreFactory = createMemoryLocalStoreFactory({
      state: {
        snapshot: seedEngine.exportSnapshot(),
        replica: seedEngine.exportReplicaSnapshot(),
        authoritativeRevision: 0,
        appliedPendingLocalSeq: 0,
      },
    })

    const runtime = new WorkbookWorkerRuntime({ localStoreFactory })
    await runtime.bootstrap({
      documentId: 'phase3-doc',
      replicaId: 'browser:test',
      persistState: true,
    })

    const received = new Array<ReturnType<typeof decodeViewportPatch>>()
    runtime.subscribeViewportPatches(
      {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 1,
        colStart: 0,
        colEnd: 1,
      },
      (bytes) => {
        received.push(decodeViewportPatch(bytes))
      },
    )

    expect(received[0]?.full).toBe(true)
    expect(received[0]?.cells.find((cell) => cell.snapshot.address === 'A1')?.displayText).toBe('7')

    await runtime.setCellFormula('Sheet1', 'B1', 'A1*2')

    expect(received).toHaveLength(2)
    expect(received[1]?.full).toBe(false)
    expect(received[1]?.cells).toHaveLength(1)
    expect(received[1]?.cells.find((cell) => cell.snapshot.address === 'B1')?.displayText).toBe('14')
  })

  it('emits invalid formulas through incremental viewport patches as #VALUE!', async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    })
    await runtime.bootstrap({
      documentId: 'invalid-formula-doc',
      replicaId: 'browser:test',
      persistState: false,
    })

    const received = new Array<ReturnType<typeof decodeViewportPatch>>()
    runtime.subscribeViewportPatches(
      {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 0,
        colStart: 0,
        colEnd: 0,
      },
      (bytes) => {
        received.push(decodeViewportPatch(bytes))
      },
    )

    await runtime.setCellFormula('Sheet1', 'A1', '1+')

    expect(received).toHaveLength(2)
    expect(received[1]?.full).toBe(false)
    expect(received[1]?.cells).toHaveLength(1)
    expect(received[1]?.cells[0]?.displayText).toBe('#VALUE!')
    expect(received[1]?.cells[0]?.editorText).toBe('#VALUE!')
  })

  it('skips persistence restore when bootstrapped in ephemeral mode', async () => {
    const seedEngine = new SpreadsheetEngine({ workbookName: 'phase3-doc', replicaId: 'seed' })
    seedEngine.createSheet('Sheet1')
    seedEngine.setCellValue('Sheet1', 'A1', 99)

    const localStoreFactory = createMemoryLocalStoreFactory({
      state: {
        snapshot: seedEngine.exportSnapshot(),
        replica: seedEngine.exportReplicaSnapshot(),
        authoritativeRevision: 0,
        appliedPendingLocalSeq: 0,
      },
    })

    const runtime = new WorkbookWorkerRuntime({ localStoreFactory })
    await runtime.bootstrap({
      documentId: 'phase3-doc',
      replicaId: 'browser:test',
      persistState: false,
    })

    expect(runtime.getCell('Sheet1', 'A1').value).toEqual({ tag: ValueTag.Empty })
  })

  it('falls back to ephemeral runtime state when the local sqlite store is locked by another tab', async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: {
        async open() {
          throw new WorkbookLocalStoreLockedError('locked')
        },
      },
    })

    const bootstrap = await runtime.bootstrap({
      documentId: 'locked-doc',
      replicaId: 'browser:test',
      persistState: true,
    })

    expect(bootstrap.restoredFromPersistence).toBe(false)
    expect(bootstrap.requiresAuthoritativeHydrate).toBe(false)
    expect(bootstrap.localPersistenceMode).toBe('follower')
    expect(runtime.getCell('Sheet1', 'A1').value).toEqual({ tag: ValueTag.Empty })
  })

  it('publishes viewport style dictionaries and stable style ids', async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    })
    await runtime.bootstrap({
      documentId: 'style-doc',
      replicaId: 'browser:test',
      persistState: false,
    })

    const received = new Array<ReturnType<typeof decodeViewportPatch>>()
    runtime.subscribeViewportPatches(
      {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 0,
        colStart: 0,
        colEnd: 0,
      },
      (bytes) => {
        received.push(decodeViewportPatch(bytes))
      },
    )

    await runtime.setRangeStyle(
      { sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' },
      { fill: { backgroundColor: '#336699' }, font: { family: 'Fira Sans' } },
    )

    const patch = received.at(-1)
    expect(patch?.full).toBe(false)
    expect(patch?.styles).toHaveLength(1)
    expect(patch?.styles[0]).toMatchObject({
      fill: { backgroundColor: '#336699' },
      font: { family: 'Fira Sans' },
    })
    expect(patch?.cells[0]?.styleId).toBe(patch?.styles[0]?.id)
  })

  it('publishes render tile replacements after style-only changes', async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    })
    await runtime.bootstrap({
      documentId: 'render-tile-style-doc',
      replicaId: 'browser:test',
      persistState: false,
    })

    const batches: RenderTileDeltaBatch[] = []
    const resolvers: Array<() => void> = []
    const waitForNextBatch = () =>
      new Promise<void>((resolve) => {
        resolvers.push(resolve)
      })
    const firstBatch = waitForNextBatch()
    const unsubscribe = runtime.subscribeRenderTileDeltas(
      {
        sheetId: 1,
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 31,
        colStart: 0,
        colEnd: 127,
        dprBucket: 1,
        initialDelta: 'full',
      },
      (bytes) => {
        batches.push(decodeRenderTileDeltaBatch(bytes))
        resolvers.shift()?.()
      },
    )

    await firstBatch
    const secondBatch = waitForNextBatch()
    await runtime.setRangeStyle({ sheetName: 'Sheet1', startAddress: 'B2', endAddress: 'B2' }, { fill: { backgroundColor: '#cfe2f3' } })
    await secondBatch
    unsubscribe()

    const initialTile = findRenderTileReplace(batches[0], 0, 0)
    const styledTile = findRenderTileReplace(batches[1], 0, 0)
    expect(styledTile.version.styles).toBeGreaterThan(initialTile.version.styles)
    expect(styledTile.rectCount).toBeGreaterThan(initialTile.rectCount)
  })

  it('publishes sheet-level workbook deltas for renderer damage consumers', async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    })
    await runtime.bootstrap({
      documentId: 'workbook-delta-doc',
      replicaId: 'browser:test',
      persistState: false,
    })

    const deltas: ReturnType<typeof decodeWorkbookDeltaBatchV3>[] = []
    const resolvers: Array<() => void> = []
    const waitForNextDelta = () =>
      new Promise<void>((resolve) => {
        resolvers.push(resolve)
      })
    const nextDelta = waitForNextDelta()
    const unsubscribe = runtime.subscribeWorkbookDeltas((bytes) => {
      deltas.push(decodeWorkbookDeltaBatchV3(bytes))
      resolvers.shift()?.()
    })

    await runtime.setCellValue('Sheet1', 'B2', 'visible')
    await nextDelta
    unsubscribe()

    const delta = deltas.at(-1)
    expect(delta).toMatchObject({
      source: 'workerAuthoritative',
      sheetId: 1,
      sheetOrdinal: 1,
    })
    const ranges = Array.from(delta?.dirty.cellRanges ?? [])
    expect(ranges.some((value, index) => index % 5 === 0 && value === 1 && ranges[index + 2] === 1)).toBe(true)
  })

  it('publishes numeric cell font-size style updates through viewport patches', async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    })
    await runtime.bootstrap({
      documentId: 'numeric-font-doc',
      replicaId: 'browser:test',
      persistState: false,
    })
    await runtime.setCellValue('Sheet1', 'G8', 1200)

    const received = new Array<ReturnType<typeof decodeViewportPatch>>()
    runtime.subscribeViewportPatches(
      {
        sheetName: 'Sheet1',
        rowStart: 7,
        rowEnd: 7,
        colStart: 6,
        colEnd: 6,
      },
      (bytes) => {
        received.push(decodeViewportPatch(bytes))
      },
    )

    await runtime.setRangeStyle({ sheetName: 'Sheet1', startAddress: 'G8', endAddress: 'G8' }, { font: { size: 20 } })

    const patch = received.at(-1)
    expect(patch?.full).toBe(false)
    expect(patch?.cells[0]?.snapshot.address).toBe('G8')
    expect(patch?.cells[0]?.displayText).toBe('1200')
    expect(patch?.cells[0]?.styleId).toBe(patch?.styles[0]?.id)
    expect(patch?.styles[0]).toMatchObject({
      font: { size: 20 },
    })
  })

  it('builds the initial full viewport patch from local projection without replaying it after materialization', async () => {
    const seedEngine = new SpreadsheetEngine({ workbookName: 'base-doc', replicaId: 'seed' })
    seedEngine.createSheet('Sheet1')
    let viewportReadCount = 0
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: {
        async open() {
          return {
            async loadBootstrapState() {
              return {
                workbookName: 'base-doc',
                sheetNames: ['Sheet1'],
                materializedCellCount: 250_000,
                authoritativeRevision: 0,
                appliedPendingLocalSeq: 0,
              }
            },
            async loadState() {
              return {
                snapshot: seedEngine.exportSnapshot(),
                replica: seedEngine.exportReplicaSnapshot(),
                authoritativeRevision: 0,
                appliedPendingLocalSeq: 0,
              }
            },
            async persistProjectionState() {},
            async ingestAuthoritativeDelta() {},
            async listPendingMutations() {
              return []
            },
            async listMutationJournalEntries() {
              return []
            },
            async appendPendingMutation() {},
            async updatePendingMutation() {},
            async removePendingMutation() {},
            readViewportProjection() {
              viewportReadCount += 1
              return {
                sheetName: 'Sheet1',
                cells: [
                  {
                    row: 0,
                    col: 0,
                    snapshot: {
                      sheetName: 'Sheet1',
                      address: 'A1',
                      value: { tag: ValueTag.Number, value: 42 },
                      flags: 0,
                      version: 1,
                    },
                  },
                ],
                rowAxisEntries: [],
                columnAxisEntries: [],
                styles: [{ id: 'style-0' }],
              }
            },
            close() {},
          }
        },
      },
    })

    await runtime.bootstrap({
      documentId: 'base-doc',
      replicaId: 'browser:test',
      persistState: true,
    })

    const received = new Array<ReturnType<typeof decodeViewportPatch>>()
    runtime.subscribeViewportPatches(
      {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 0,
        colStart: 0,
        colEnd: 0,
      },
      (bytes) => {
        received.push(decodeViewportPatch(bytes))
      },
    )

    expect(viewportReadCount).toBe(1)
    expect(received[0]?.cells[0]?.displayText).toBe('42')
    await runtime.materializeProjectionEngine()
    await Promise.resolve()

    expect(viewportReadCount).toBeGreaterThanOrEqual(2)
    expect(received).toHaveLength(1)
  })

  it('prefers installed engine calculations over stale persisted local projection on first viewport patch', async () => {
    const seedEngine = new SpreadsheetEngine({ workbookName: 'formula-doc', replicaId: 'seed' })
    seedEngine.createSheet('Sheet1')
    seedEngine.setCellValue('Sheet1', 'A1', 12)
    seedEngine.setCellFormula('Sheet1', 'B1', 'A1/2')
    let viewportReadCount = 0

    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory({
        state: {
          snapshot: seedEngine.exportSnapshot(),
          replica: seedEngine.exportReplicaSnapshot(),
          authoritativeRevision: 0,
          appliedPendingLocalSeq: 0,
        },
        onReadViewportProjection() {
          viewportReadCount += 1
        },
        projectionOverlay: {
          cells: [
            {
              sheetId: 1,
              sheetName: 'Sheet1',
              address: 'B1',
              rowNum: 0,
              colNum: 1,
              value: { tag: ValueTag.Error, code: ErrorCode.Div0 },
              flags: 0,
              version: 2,
              input: '=A1/2',
              formula: 'A1/2',
              format: undefined,
              styleId: undefined,
              numberFormatId: undefined,
            },
          ],
          rowAxisEntries: [],
          columnAxisEntries: [],
          styles: [],
        },
      }),
    })

    await runtime.bootstrap({
      documentId: 'formula-doc',
      replicaId: 'browser:test',
      persistState: true,
    })

    const received = new Array<ReturnType<typeof decodeViewportPatch>>()
    runtime.subscribeViewportPatches(
      {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 0,
        colStart: 0,
        colEnd: 1,
      },
      (bytes) => {
        received.push(decodeViewportPatch(bytes))
      },
    )

    expect(viewportReadCount).toBe(1)
    expect(received[0]?.cells.find((cell) => cell.snapshot.address === 'B1')?.displayText).toBe('6')
  })

  it('renders date-formatted cells from persisted local projection on the first viewport patch', async () => {
    const seedEngine = new SpreadsheetEngine({ workbookName: 'date-doc', replicaId: 'seed' })
    seedEngine.createSheet('Sheet1')
    seedEngine.setCellValue('Sheet1', 'A1', 46023)
    seedEngine.workbook.upsertCellNumberFormat(createCellNumberFormatRecord('format-date', 'm/d/yyyy'))
    seedEngine.workbook.setFormatRange({ sheetName: 'Sheet1', startAddress: 'A1', endAddress: 'A1' }, 'format-date')
    const expectedDisplay = formatCellDisplayValue(seedEngine.getCell('Sheet1', 'A1').value, seedEngine.getCell('Sheet1', 'A1').format)
    let viewportReadCount = 0

    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory({
        state: {
          snapshot: seedEngine.exportSnapshot(),
          replica: seedEngine.exportReplicaSnapshot(),
          authoritativeRevision: 0,
          appliedPendingLocalSeq: 0,
        },
        onReadViewportProjection() {
          viewportReadCount += 1
        },
      }),
    })

    await runtime.bootstrap({
      documentId: 'date-doc',
      replicaId: 'browser:test',
      persistState: true,
    })

    const received = new Array<ReturnType<typeof decodeViewportPatch>>()
    runtime.subscribeViewportPatches(
      {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 0,
        colStart: 0,
        colEnd: 0,
      },
      (bytes) => {
        received.push(decodeViewportPatch(bytes))
      },
    )

    expect(viewportReadCount).toBe(1)
    expect(received[0]?.cells[0]?.displayText).toBe(expectedDisplay)
  })

  it('renders readable dates for inferred date cells on the first viewport patch', async () => {
    const seedEngine = new SpreadsheetEngine({ workbookName: 'date-inference-doc', replicaId: 'seed' })
    seedEngine.createSheet('Sheet1')
    seedEngine.setCellValue('Sheet1', 'A1', 'Month')
    seedEngine.setCellFormula('Sheet1', 'A2', 'DATE(2026,12,1)')
    seedEngine.setCellValue('Sheet1', 'B1', 'Start Date')
    seedEngine.setCellValue('Sheet1', 'B2', 46023)
    let viewportReadCount = 0

    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory({
        state: {
          snapshot: seedEngine.exportSnapshot(),
          replica: seedEngine.exportReplicaSnapshot(),
          authoritativeRevision: 0,
          appliedPendingLocalSeq: 0,
        },
        onReadViewportProjection() {
          viewportReadCount += 1
        },
      }),
    })

    await runtime.bootstrap({
      documentId: 'date-inference-doc',
      replicaId: 'browser:test',
      persistState: true,
    })

    const received = new Array<ReturnType<typeof decodeViewportPatch>>()
    runtime.subscribeViewportPatches(
      {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 1,
        colStart: 0,
        colEnd: 1,
      },
      (bytes) => {
        received.push(decodeViewportPatch(bytes))
      },
    )

    expect(viewportReadCount).toBe(1)
    expect(received[0]?.cells.find((cell) => cell.snapshot.address === 'A2')?.displayText).toBe('12/01/2026')
    expect(received[0]?.cells.find((cell) => cell.snapshot.address === 'B2')?.displayText).toBe('01/01/2026')
  })

  it('prefers installed engine date inference over stale persisted local projection on first viewport patch', async () => {
    const seedEngine = new SpreadsheetEngine({ workbookName: 'date-projection-doc', replicaId: 'seed' })
    seedEngine.createSheet('Sheet1')
    seedEngine.setCellValue('Sheet1', 'A1', 'Month')
    seedEngine.setCellFormula('Sheet1', 'A2', 'DATE(2026,12,1)')
    seedEngine.setCellValue('Sheet1', 'B1', 'Start Date')
    seedEngine.setCellValue('Sheet1', 'B2', 46023)
    let viewportReadCount = 0

    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory({
        state: {
          snapshot: seedEngine.exportSnapshot(),
          replica: seedEngine.exportReplicaSnapshot(),
          authoritativeRevision: 0,
          appliedPendingLocalSeq: 0,
        },
        onReadViewportProjection() {
          viewportReadCount += 1
        },
        projectionOverlay: {
          cells: [
            {
              sheetId: 1,
              sheetName: 'Sheet1',
              address: 'A2',
              rowNum: 1,
              colNum: 0,
              value: { tag: ValueTag.Number, value: 46357 },
              flags: 0,
              version: 2,
              input: '=DATE(2026,12,1)',
              formula: 'DATE(2026,12,1)',
              format: undefined,
              styleId: undefined,
              numberFormatId: undefined,
            },
            {
              sheetId: 1,
              sheetName: 'Sheet1',
              address: 'B2',
              rowNum: 1,
              colNum: 1,
              value: { tag: ValueTag.Number, value: 46023 },
              flags: 0,
              version: 2,
              input: 46023,
              formula: undefined,
              format: undefined,
              styleId: undefined,
              numberFormatId: undefined,
            },
          ],
          rowAxisEntries: [],
          columnAxisEntries: [],
          styles: [],
        },
      }),
    })

    await runtime.bootstrap({
      documentId: 'date-projection-doc',
      replicaId: 'browser:test',
      persistState: true,
    })

    const received = new Array<ReturnType<typeof decodeViewportPatch>>()
    runtime.subscribeViewportPatches(
      {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 1,
        colStart: 0,
        colEnd: 1,
      },
      (bytes) => {
        received.push(decodeViewportPatch(bytes))
      },
    )

    expect(viewportReadCount).toBe(1)
    expect(received[0]?.cells.find((cell) => cell.snapshot.address === 'A2')?.displayText).toBe('12/01/2026')
    expect(received[0]?.cells.find((cell) => cell.snapshot.address === 'B2')?.displayText).toBe('01/01/2026')
  })

  it('defers persisted snapshot parsing until the projection engine is actually needed', async () => {
    vi.useFakeTimers()
    const seedEngine = new SpreadsheetEngine({ workbookName: 'lazy-doc', replicaId: 'seed' })
    seedEngine.createSheet('Sheet1')
    seedEngine.setCellValue('Sheet1', 'A1', 42)

    let loadStateCount = 0
    let viewportReadCount = 0
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: {
        async open() {
          return {
            async loadBootstrapState() {
              return {
                workbookName: 'lazy-doc',
                sheetNames: ['Sheet1'],
                materializedCellCount: 250_000,
                authoritativeRevision: 0,
                appliedPendingLocalSeq: 0,
              }
            },
            async loadState() {
              loadStateCount += 1
              return {
                snapshot: seedEngine.exportSnapshot(),
                replica: seedEngine.exportReplicaSnapshot(),
                authoritativeRevision: 0,
                appliedPendingLocalSeq: 0,
              }
            },
            async persistProjectionState() {},
            async ingestAuthoritativeDelta() {},
            async listPendingMutations() {
              return []
            },
            async listMutationJournalEntries() {
              return []
            },
            async appendPendingMutation() {},
            async updatePendingMutation() {},
            async removePendingMutation() {},
            readViewportProjection() {
              viewportReadCount += 1
              return {
                sheetId: 1,
                sheetName: 'Sheet1',
                cells: [
                  {
                    row: 0,
                    col: 0,
                    snapshot: {
                      sheetName: 'Sheet1',
                      address: 'A1',
                      value: { tag: ValueTag.Number, value: 42 },
                      flags: 0,
                      version: 1,
                    },
                  },
                ],
                rowAxisEntries: [],
                columnAxisEntries: [],
                styles: [{ id: 'style-0' }],
              }
            },
            close() {},
          }
        },
      },
    })

    await runtime.bootstrap({
      documentId: 'lazy-doc',
      replicaId: 'browser:test',
      persistState: true,
    })

    expect(loadStateCount).toBe(0)
    expect(runtime.getCell('Sheet1', 'A1').value).toEqual({
      tag: ValueTag.Number,
      value: 42,
    })
    expect(viewportReadCount).toBe(1)
    expect(loadStateCount).toBe(0)

    await runtime.setCellValue('Sheet1', 'A1', 99)

    expect(loadStateCount).toBe(1)
    expect(runtime.getCell('Sheet1', 'A1').value).toEqual({
      tag: ValueTag.Number,
      value: 99,
    })
  })

  it('formats inferred dates directly from deferred local projection', async () => {
    vi.useFakeTimers()
    const seedEngine = new SpreadsheetEngine({ workbookName: 'lazy-date-doc', replicaId: 'seed' })
    seedEngine.createSheet('Sheet1')
    seedEngine.setCellValue('Sheet1', 'A1', 'Month')
    seedEngine.setCellFormula('Sheet1', 'A2', 'DATE(2026,12,1)')
    seedEngine.setCellValue('Sheet1', 'B1', 'Start Date')
    seedEngine.setCellValue('Sheet1', 'B2', 46023)

    let loadStateCount = 0
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: {
        async open() {
          return {
            async loadBootstrapState() {
              return {
                workbookName: 'lazy-date-doc',
                sheetNames: ['Sheet1'],
                materializedCellCount: 250_000,
                authoritativeRevision: 0,
                appliedPendingLocalSeq: 0,
              }
            },
            async loadState() {
              loadStateCount += 1
              return {
                snapshot: seedEngine.exportSnapshot(),
                replica: seedEngine.exportReplicaSnapshot(),
                authoritativeRevision: 0,
                appliedPendingLocalSeq: 0,
              }
            },
            async persistProjectionState() {},
            async ingestAuthoritativeDelta() {},
            async listPendingMutations() {
              return []
            },
            async listMutationJournalEntries() {
              return []
            },
            async appendPendingMutation() {},
            async updatePendingMutation() {},
            async removePendingMutation() {},
            readViewportProjection() {
              return {
                sheetId: 1,
                sheetName: 'Sheet1',
                cells: [
                  {
                    row: 0,
                    col: 0,
                    snapshot: {
                      sheetName: 'Sheet1',
                      address: 'A1',
                      value: { tag: ValueTag.String, value: 'Month' },
                      flags: 0,
                      version: 1,
                    },
                  },
                  {
                    row: 1,
                    col: 0,
                    snapshot: {
                      sheetName: 'Sheet1',
                      address: 'A2',
                      value: { tag: ValueTag.Number, value: 46357 },
                      flags: 0,
                      version: 1,
                      formula: 'DATE(2026,12,1)',
                    },
                  },
                  {
                    row: 0,
                    col: 1,
                    snapshot: {
                      sheetName: 'Sheet1',
                      address: 'B1',
                      value: { tag: ValueTag.String, value: 'Start Date' },
                      flags: 0,
                      version: 1,
                    },
                  },
                  {
                    row: 1,
                    col: 1,
                    snapshot: {
                      sheetName: 'Sheet1',
                      address: 'B2',
                      value: { tag: ValueTag.Number, value: 46023 },
                      flags: 0,
                      version: 1,
                    },
                  },
                ],
                rowAxisEntries: [],
                columnAxisEntries: [],
                styles: [{ id: 'style-0' }],
              }
            },
            close() {},
          }
        },
      },
    })

    await runtime.bootstrap({
      documentId: 'lazy-date-doc',
      replicaId: 'browser:test',
      persistState: true,
    })

    const received = new Array<ReturnType<typeof decodeViewportPatch>>()
    runtime.subscribeViewportPatches(
      {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 1,
        colStart: 0,
        colEnd: 1,
      },
      (bytes) => {
        received.push(decodeViewportPatch(bytes))
      },
    )

    expect(received).toHaveLength(1)
    expect(received[0]?.cells.find((cell) => cell.snapshot.address === 'A2')?.displayText).toBe('12/01/2026')
    expect(received[0]?.cells.find((cell) => cell.snapshot.address === 'B2')?.displayText).toBe('01/01/2026')
    expect(loadStateCount).toBe(0)

    await vi.runAllTimersAsync()

    expect(loadStateCount).toBe(1)
    expect(received).toHaveLength(1)
  })

  it('does not rewrite normalized sqlite state on a clean persisted restore', async () => {
    const seedEngine = new SpreadsheetEngine({
      workbookName: 'restored-no-rewrite-doc',
      replicaId: 'seed',
    })
    seedEngine.createSheet('Sheet1')
    seedEngine.setCellValue('Sheet1', 'A1', 7)

    const persistProjectionState = vi.fn(async () => {})
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory({
        state: {
          snapshot: seedEngine.exportSnapshot(),
          replica: seedEngine.exportReplicaSnapshot(),
          authoritativeRevision: 0,
          appliedPendingLocalSeq: 0,
        },
        onPersistProjectionState: persistProjectionState,
      }),
    })

    await runtime.bootstrap({
      documentId: 'restored-no-rewrite-doc',
      replicaId: 'browser:test',
      persistState: true,
    })

    expect(persistProjectionState).toHaveBeenCalledTimes(0)
    expect(runtime.getCell('Sheet1', 'A1').value).toEqual({
      tag: ValueTag.Number,
      value: 7,
    })
  })

  it('reads initial local full patches through 128x32 worker tiles instead of a single wide viewport query', async () => {
    const seedEngine = new SpreadsheetEngine({ workbookName: 'tile-doc', replicaId: 'seed' })
    seedEngine.createSheet('Sheet1')
    seedEngine.setCellValue('Sheet1', 'A1', 1)
    seedEngine.setCellValue('Sheet1', formatAddress(0, 128), 2)
    seedEngine.setCellValue('Sheet1', formatAddress(32, 0), 3)
    seedEngine.setCellValue('Sheet1', formatAddress(32, 128), 4)

    const viewportReads: Array<{
      rowStart: number
      rowEnd: number
      colStart: number
      colEnd: number
    }> = []
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory({
        state: {
          snapshot: seedEngine.exportSnapshot(),
          replica: seedEngine.exportReplicaSnapshot(),
          authoritativeRevision: 0,
          appliedPendingLocalSeq: 0,
        },
        authoritativeBase: buildWorkbookLocalAuthoritativeBase(seedEngine),
        onReadViewportProjection(_sheetName, viewport) {
          viewportReads.push({ ...viewport })
        },
      }),
    })

    await runtime.bootstrap({
      documentId: 'tile-doc',
      replicaId: 'browser:test',
      persistState: true,
    })

    const received = new Array<ReturnType<typeof decodeViewportPatch>>()
    runtime.subscribeViewportPatches(
      {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 40,
        colStart: 0,
        colEnd: 140,
      },
      (bytes) => {
        received.push(decodeViewportPatch(bytes))
      },
    )

    expect(viewportReads).toEqual([
      { rowStart: 0, rowEnd: 31, colStart: 0, colEnd: 127 },
      { rowStart: 0, rowEnd: 31, colStart: 128, colEnd: 255 },
      { rowStart: 32, rowEnd: 63, colStart: 0, colEnd: 127 },
      { rowStart: 32, rowEnd: 63, colStart: 128, colEnd: 255 },
    ])
    expect(received[0]?.cells.map((cell) => cell.displayText).toSorted()).toEqual(['1', '2', '3', '4'])
  })

  it('restores pending local projection overlays from persistence across bootstrap', async () => {
    const seedEngine = new SpreadsheetEngine({ workbookName: 'overlay-doc', replicaId: 'seed' })
    seedEngine.createSheet('Sheet1')
    seedEngine.setCellValue('Sheet1', 'A1', 5)
    let viewportReadCount = 0

    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory({
        state: {
          snapshot: seedEngine.exportSnapshot(),
          replica: seedEngine.exportReplicaSnapshot(),
          authoritativeRevision: 0,
          appliedPendingLocalSeq: 1,
        },
        pendingMutations: [
          {
            id: 'overlay-doc:pending:1',
            localSeq: 1,
            baseRevision: 0,
            method: 'setCellValue',
            args: ['Sheet1', 'A1', 17],
            enqueuedAtUnixMs: 1,
            submittedAtUnixMs: null,
            lastAttemptedAtUnixMs: null,
            ackedAtUnixMs: null,
            rebasedAtUnixMs: null,
            failedAtUnixMs: null,
            attemptCount: 0,
            failureMessage: null,
            status: 'local',
          },
        ],
        onReadViewportProjection() {
          viewportReadCount += 1
        },
        authoritativeBase: buildWorkbookLocalAuthoritativeBase(seedEngine),
        projectionOverlay: {
          cells: [
            {
              sheetId: 1,
              sheetName: 'Sheet1',
              address: 'A1',
              rowNum: 0,
              colNum: 0,
              value: { tag: ValueTag.Number, value: 17 },
              flags: 0,
              version: 2,
              input: 17,
              formula: undefined,
              format: undefined,
              styleId: undefined,
              numberFormatId: undefined,
            },
          ],
          rowAxisEntries: [],
          columnAxisEntries: [],
          styles: [],
        },
      }),
    })

    const bootstrap = await runtime.bootstrap({
      documentId: 'overlay-doc',
      replicaId: 'browser:test',
      persistState: true,
    })

    expect(bootstrap.restoredFromPersistence).toBe(true)
    expect(bootstrap.requiresAuthoritativeHydrate).toBe(false)
    expect(runtime.getCell('Sheet1', 'A1').value).toEqual({
      tag: ValueTag.Number,
      value: 17,
    })

    const received = new Array<ReturnType<typeof decodeViewportPatch>>()
    runtime.subscribeViewportPatches(
      {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 0,
        colStart: 0,
        colEnd: 0,
      },
      (bytes) => {
        received.push(decodeViewportPatch(bytes))
      },
    )

    expect(viewportReadCount).toBe(1)
    expect(received[0]?.cells[0]?.displayText).toBe('17')
  })

  it('patches only affected axis entries for column metadata edits', async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    })
    await runtime.bootstrap({
      documentId: 'axis-doc',
      replicaId: 'browser:test',
      persistState: false,
    })

    const received = new Array<ReturnType<typeof decodeViewportPatch>>()
    runtime.subscribeViewportPatches(
      {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 2,
        colStart: 0,
        colEnd: 3,
      },
      (bytes) => {
        received.push(decodeViewportPatch(bytes))
      },
    )

    await runtime.updateColumnWidth('Sheet1', 1, 160)

    const patch = received.at(-1)
    expect(patch?.full).toBe(false)
    expect(patch?.cells).toHaveLength(0)
    expect(patch?.rows).toHaveLength(0)
    expect(patch?.columns).toEqual([{ index: 1, size: 160, hidden: false }])
  })

  it('patches only affected axis entries for row metadata edits', async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    })
    await runtime.bootstrap({
      documentId: 'row-axis-doc',
      replicaId: 'browser:test',
      persistState: false,
    })

    const received = new Array<ReturnType<typeof decodeViewportPatch>>()
    runtime.subscribeViewportPatches(
      {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 3,
        colStart: 0,
        colEnd: 2,
      },
      (bytes) => {
        received.push(decodeViewportPatch(bytes))
      },
    )

    await runtime.updateRowMetadata('Sheet1', 2, 1, 42, true)

    const patch = received.at(-1)
    expect(patch?.full).toBe(false)
    expect(patch?.cells).toHaveLength(0)
    expect(patch?.columns).toHaveLength(0)
    expect(patch?.rows).toEqual([{ index: 2, size: 42, hidden: true }])
  })

  it('persists pending workbook mutations across bootstraps and removes them on ack', async () => {
    const localStoreFactory = createMemoryLocalStoreFactory()
    const runtime = new WorkbookWorkerRuntime({ localStoreFactory })
    await runtime.bootstrap({
      documentId: 'pending-doc',
      replicaId: 'browser:test',
      persistState: true,
    })

    const pending = await runtime.enqueuePendingMutation({
      method: 'setCellValue',
      args: ['Sheet1', 'A1', 17],
    })

    expect(runtime.listPendingMutations()).toEqual([pending])

    const reloaded = new WorkbookWorkerRuntime({ localStoreFactory })
    await reloaded.bootstrap({
      documentId: 'pending-doc',
      replicaId: 'browser:reloaded',
      persistState: true,
    })

    expect(reloaded.listPendingMutations()).toEqual([pending])

    await reloaded.ackPendingMutation(pending.id)
    expect(reloaded.listPendingMutations()).toEqual([])
    expect(reloaded.listMutationJournalEntries()).toEqual([
      {
        ...pending,
        args: [...pending.args],
        ackedAtUnixMs: expect.any(Number),
        status: 'acked',
      },
    ])

    const afterAck = new WorkbookWorkerRuntime({ localStoreFactory })
    await afterAck.bootstrap({
      documentId: 'pending-doc',
      replicaId: 'browser:after-ack',
      persistState: true,
    })

    expect(afterAck.listPendingMutations()).toEqual([])
  })

  it('absorbs submitted pending mutations when authoritative events arrive', async () => {
    const localStoreFactory = createMemoryLocalStoreFactory()
    const runtime = new WorkbookWorkerRuntime({ localStoreFactory })
    await runtime.bootstrap({
      documentId: 'authoritative-doc',
      replicaId: 'browser:test',
      persistState: true,
    })

    const pending = await runtime.enqueuePendingMutation({
      method: 'setCellValue',
      args: ['Sheet1', 'A1', 17],
    })
    await runtime.markPendingMutationSubmitted(pending.id)

    expect(runtime.listPendingMutations()).toEqual([
      {
        ...pending,
        args: [...pending.args],
        submittedAtUnixMs: expect.any(Number),
        status: 'submitted',
      },
    ])

    await runtime.applyAuthoritativeEvents(
      [
        {
          revision: 1,
          clientMutationId: pending.id,
          payload: {
            kind: 'setCellValue',
            sheetName: 'Sheet1',
            address: 'A1',
            value: 17,
          },
        },
      ],
      1,
    )

    expect(runtime.listPendingMutations()).toEqual([])
    expect(runtime.getCell('Sheet1', 'A1').value).toEqual({
      tag: ValueTag.Number,
      value: 17,
    })

    const reloaded = new WorkbookWorkerRuntime({ localStoreFactory })
    await reloaded.bootstrap({
      documentId: 'authoritative-doc',
      replicaId: 'browser:reloaded',
      persistState: true,
    })

    expect(reloaded.listPendingMutations()).toEqual([])
    expect(reloaded.getCell('Sheet1', 'A1').value).toEqual({
      tag: ValueTag.Number,
      value: 17,
    })
  })

  it('reopens submitted pending mutations from sqlite and absorbs them on authoritative ack', async () => {
    const localStoreFactory = createMemoryWorkbookLocalStoreFactory()
    const runtime = new WorkbookWorkerRuntime({ localStoreFactory })
    await runtime.bootstrap({
      documentId: 'submitted-reopen-doc',
      replicaId: 'browser:test',
      persistState: true,
    })

    const pending = await runtime.enqueuePendingMutation({
      method: 'setCellValue',
      args: ['Sheet1', 'A1', 17],
    })
    await runtime.markPendingMutationSubmitted(pending.id)

    const reloaded = new WorkbookWorkerRuntime({ localStoreFactory })
    await reloaded.bootstrap({
      documentId: 'submitted-reopen-doc',
      replicaId: 'browser:reloaded',
      persistState: true,
    })

    expect(reloaded.listPendingMutations()).toEqual([
      {
        ...pending,
        args: [...pending.args],
        submittedAtUnixMs: expect.any(Number),
        status: 'submitted',
      },
    ])

    await reloaded.applyAuthoritativeEvents(
      [
        {
          revision: 1,
          clientMutationId: pending.id,
          payload: {
            kind: 'setCellValue',
            sheetName: 'Sheet1',
            address: 'A1',
            value: 17,
          },
        },
      ],
      1,
    )

    expect(reloaded.listPendingMutations()).toEqual([])
    expect(reloaded.getCell('Sheet1', 'A1').value).toEqual({
      tag: ValueTag.Number,
      value: 17,
    })

    const afterAck = new WorkbookWorkerRuntime({ localStoreFactory })
    await afterAck.bootstrap({
      documentId: 'submitted-reopen-doc',
      replicaId: 'browser:after-ack',
      persistState: true,
    })

    expect(afterAck.listPendingMutations()).toEqual([])
    expect(afterAck.getCell('Sheet1', 'A1').value).toEqual({
      tag: ValueTag.Number,
      value: 17,
    })
  })

  it('marks unsent pending mutations as rebased when authoritative events replay over them', async () => {
    const localStoreFactory = createMemoryLocalStoreFactory()
    const runtime = new WorkbookWorkerRuntime({ localStoreFactory })
    await runtime.bootstrap({
      documentId: 'rebased-local-doc',
      replicaId: 'browser:test',
      persistState: true,
    })

    const pending = await runtime.enqueuePendingMutation({
      method: 'setCellValue',
      args: ['Sheet1', 'A1', 17],
    })

    await runtime.applyAuthoritativeEvents(
      [
        {
          revision: 1,
          clientMutationId: null,
          payload: {
            kind: 'setCellValue',
            sheetName: 'Sheet1',
            address: 'B1',
            value: 101,
          },
        },
      ],
      1,
    )

    expect(runtime.listPendingMutations()).toEqual([
      {
        ...pending,
        args: [...pending.args],
        rebasedAtUnixMs: expect.any(Number),
        status: 'rebased',
      },
    ])

    const reloaded = new WorkbookWorkerRuntime({ localStoreFactory })
    await reloaded.bootstrap({
      documentId: 'rebased-local-doc',
      replicaId: 'browser:reloaded',
      persistState: true,
    })

    expect(reloaded.listPendingMutations()).toEqual([
      {
        ...pending,
        args: [...pending.args],
        rebasedAtUnixMs: expect.any(Number),
        status: 'rebased',
      },
    ])
    expect(reloaded.getCell('Sheet1', 'A1').value).toEqual({
      tag: ValueTag.Number,
      value: 17,
    })
    expect(reloaded.getCell('Sheet1', 'B1').value).toEqual({
      tag: ValueTag.Number,
      value: 101,
    })
  })

  it('persists failed mutations until they are explicitly retried', async () => {
    const localStoreFactory = createMemoryLocalStoreFactory()
    const runtime = new WorkbookWorkerRuntime({ localStoreFactory })
    await runtime.bootstrap({
      documentId: 'failed-journal-doc',
      replicaId: 'browser:test',
      persistState: true,
    })

    const pending = await runtime.enqueuePendingMutation({
      method: 'setCellValue',
      args: ['Sheet1', 'A1', 17],
    })

    await runtime.markPendingMutationFailed(pending.id, 'mutation rejected by server')

    expect(runtime.listPendingMutations()).toEqual([
      {
        ...pending,
        args: [...pending.args],
        failedAtUnixMs: expect.any(Number),
        failureMessage: 'mutation rejected by server',
        status: 'failed',
      },
    ])

    const reloaded = new WorkbookWorkerRuntime({ localStoreFactory })
    await reloaded.bootstrap({
      documentId: 'failed-journal-doc',
      replicaId: 'browser:reloaded',
      persistState: true,
    })

    expect(reloaded.listPendingMutations()).toEqual([
      {
        ...pending,
        args: [...pending.args],
        failedAtUnixMs: expect.any(Number),
        failureMessage: 'mutation rejected by server',
        status: 'failed',
      },
    ])

    await reloaded.retryPendingMutation(pending.id)

    expect(reloaded.listPendingMutations()).toEqual([
      {
        ...pending,
        args: [...pending.args],
        status: 'local',
      },
    ])
  })

  it('ingests narrow authoritative event batches through delta persistence', async () => {
    const persistProjectionState = vi.fn(async () => {})
    const ingestAuthoritativeDelta = vi.fn(async () => {})
    const localStoreFactory = createMemoryLocalStoreFactory({
      onPersistProjectionState: persistProjectionState,
      onIngestAuthoritativeDelta: ingestAuthoritativeDelta,
    })
    const runtime = new WorkbookWorkerRuntime({ localStoreFactory })
    await runtime.bootstrap({
      documentId: 'authoritative-delta-doc',
      replicaId: 'browser:test',
      persistState: true,
    })

    expect(persistProjectionState).toHaveBeenCalledTimes(1)
    expect(ingestAuthoritativeDelta).toHaveBeenCalledTimes(0)

    await runtime.applyAuthoritativeEvents(
      [
        {
          revision: 1,
          clientMutationId: null,
          payload: {
            kind: 'setCellValue',
            sheetName: 'Sheet1',
            address: 'B2',
            value: 23,
          },
        },
      ],
      1,
    )

    expect(persistProjectionState).toHaveBeenCalledTimes(1)
    expect(ingestAuthoritativeDelta).toHaveBeenCalledTimes(1)

    const reloaded = new WorkbookWorkerRuntime({ localStoreFactory })
    await reloaded.bootstrap({
      documentId: 'authoritative-delta-doc',
      replicaId: 'browser:reloaded',
      persistState: true,
    })

    expect(reloaded.getCell('Sheet1', 'B2').value).toEqual({
      tag: ValueTag.Number,
      value: 23,
    })
  })

  it('applies restoreVersion authoritative events as full workbook replaces', async () => {
    const persistProjectionState = vi.fn(async () => {})
    const ingestAuthoritativeDelta = vi.fn(async () => {})
    const localStoreFactory = createMemoryLocalStoreFactory({
      onPersistProjectionState: persistProjectionState,
      onIngestAuthoritativeDelta: ingestAuthoritativeDelta,
    })
    const runtime = new WorkbookWorkerRuntime({ localStoreFactory })
    await runtime.bootstrap({
      documentId: 'authoritative-version-doc',
      replicaId: 'browser:test',
      persistState: true,
    })

    await runtime.applyAuthoritativeEvents(
      [
        {
          revision: 1,
          clientMutationId: null,
          payload: {
            kind: 'restoreVersion',
            versionId: 'version-1',
            versionName: 'Month close',
            sheetName: 'Sheet1',
            address: 'D5',
            snapshot: {
              version: 1,
              workbook: { name: 'authoritative-version-doc' },
              sheets: [
                {
                  id: 1,
                  name: 'Sheet1',
                  order: 0,
                  cells: [{ address: 'D5', value: 'restored' }],
                },
              ],
            },
          },
        },
      ],
      1,
    )

    expect(ingestAuthoritativeDelta).toHaveBeenCalledTimes(1)

    const reloaded = new WorkbookWorkerRuntime({ localStoreFactory })
    await reloaded.bootstrap({
      documentId: 'authoritative-version-doc',
      replicaId: 'browser:reloaded',
      persistState: true,
    })

    expect(reloaded.getCell('Sheet1', 'D5').value).toEqual({
      tag: ValueTag.String,
      value: 'restored',
      stringId: expect.any(Number),
    })
  })

  it('replays journaled mutations that were not yet captured in the persisted snapshot', async () => {
    const seedEngine = new SpreadsheetEngine({ workbookName: 'journal-doc', replicaId: 'seed' })
    seedEngine.createSheet('Sheet1')
    seedEngine.setCellValue('Sheet1', 'A1', 5)

    const localStoreFactory = createMemoryLocalStoreFactory({
      state: {
        snapshot: seedEngine.exportSnapshot(),
        replica: seedEngine.exportReplicaSnapshot(),
        authoritativeRevision: 0,
        appliedPendingLocalSeq: 0,
      },
      pendingMutations: [
        {
          id: 'journal-doc:pending:1',
          localSeq: 1,
          baseRevision: 0,
          method: 'setCellValue',
          args: ['Sheet1', 'A1', 17],
          enqueuedAtUnixMs: 1,
          submittedAtUnixMs: null,
          lastAttemptedAtUnixMs: null,
          ackedAtUnixMs: null,
          rebasedAtUnixMs: null,
          failedAtUnixMs: null,
          attemptCount: 0,
          failureMessage: null,
          status: 'local',
        },
      ],
    })

    const runtime = new WorkbookWorkerRuntime({ localStoreFactory })
    await runtime.bootstrap({
      documentId: 'journal-doc',
      replicaId: 'browser:test',
      persistState: true,
    })

    expect(runtime.getCell('Sheet1', 'A1').value).toEqual({
      tag: ValueTag.Number,
      value: 17,
    })
    expect(runtime.listPendingMutations()).toHaveLength(1)
  })

  it('installs authoritative reconcile snapshots by replaying pending local mutations', async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    })
    await runtime.bootstrap({
      documentId: 'rebase-doc',
      replicaId: 'browser:test',
      persistState: true,
    })

    await runtime.setCellValue('Sheet1', 'A1', 17)
    await runtime.enqueuePendingMutation({
      method: 'setCellValue',
      args: ['Sheet1', 'A1', 17],
    })

    await runtime.installAuthoritativeSnapshot({
      snapshot: {
        version: 1,
        workbook: { name: 'rebase-doc' },
        sheets: [
          {
            name: 'Sheet1',
            order: 0,
            cells: [{ address: 'A1', value: 5 }],
          },
        ],
      },
      authoritativeRevision: 3,
      mode: 'reconcile',
    })

    expect(runtime.getAuthoritativeRevision()).toBe(3)
    expect(runtime.getCell('Sheet1', 'A1').value).toEqual({
      tag: ValueTag.Number,
      value: 17,
    })
    expect(runtime.listPendingMutations()).toHaveLength(1)
  })

  it('skips unrelated viewport subscriptions when an edit is outside their sheet or region', async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    })
    await runtime.bootstrap({
      documentId: 'fanout-doc',
      replicaId: 'browser:test',
      persistState: false,
    })

    await runtime.renderCommit([{ kind: 'upsertSheet', name: 'Sheet2', order: 1 }])

    const primary = new Array<ReturnType<typeof decodeViewportPatch>>()
    const offsheet = new Array<ReturnType<typeof decodeViewportPatch>>()
    const offregion = new Array<ReturnType<typeof decodeViewportPatch>>()

    runtime.subscribeViewportPatches(
      {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 2,
        colStart: 0,
        colEnd: 2,
      },
      (bytes) => {
        primary.push(decodeViewportPatch(bytes))
      },
    )

    runtime.subscribeViewportPatches(
      {
        sheetName: 'Sheet2',
        rowStart: 0,
        rowEnd: 2,
        colStart: 0,
        colEnd: 2,
      },
      (bytes) => {
        offsheet.push(decodeViewportPatch(bytes))
      },
    )

    runtime.subscribeViewportPatches(
      {
        sheetName: 'Sheet1',
        rowStart: 10,
        rowEnd: 12,
        colStart: 10,
        colEnd: 12,
      },
      (bytes) => {
        offregion.push(decodeViewportPatch(bytes))
      },
    )

    expect(primary).toHaveLength(1)
    expect(offsheet).toHaveLength(1)
    expect(offregion).toHaveLength(1)

    await runtime.setCellValue('Sheet1', 'A1', 123)

    expect(primary).toHaveLength(2)
    expect(primary[1]?.cells[0]?.snapshot.address).toBe('A1')
    expect(offsheet).toHaveLength(1)
    expect(offregion).toHaveLength(1)
  })

  it('builds viewport patches only for subscriptions on impacted sheets', async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    })
    await runtime.bootstrap({
      documentId: 'sheet-index-doc',
      replicaId: 'browser:test',
      persistState: false,
    })

    await runtime.renderCommit([{ kind: 'upsertSheet', name: 'Sheet2', order: 1 }])

    runtime.subscribeViewportPatches(
      {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 2,
        colStart: 0,
        colEnd: 2,
      },
      () => {},
    )
    runtime.subscribeViewportPatches(
      {
        sheetName: 'Sheet2',
        rowStart: 0,
        rowEnd: 2,
        colStart: 0,
        colEnd: 2,
      },
      () => {},
    )
    runtime.subscribeViewportPatches(
      {
        sheetName: 'Sheet2',
        rowStart: 10,
        rowEnd: 12,
        colStart: 10,
        colEnd: 12,
      },
      () => {},
    )

    const originalBuildViewportPatch = runtime['buildViewportPatch']
    if (typeof originalBuildViewportPatch !== 'function') {
      throw new Error('Expected buildViewportPatch method')
    }

    let buildViewportPatchCalls = 0
    runtime['buildViewportPatch'] = (...args: unknown[]) => {
      buildViewportPatchCalls += 1
      return Reflect.apply(originalBuildViewportPatch, runtime, args)
    }

    await runtime.setCellValue('Sheet1', 'A1', 321)

    expect(buildViewportPatchCalls).toBe(1)
    runtime['buildViewportPatch'] = originalBuildViewportPatch
  })

  it('dedupes changed viewport cells against invalidated range expansion', async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    })
    await runtime.bootstrap({
      documentId: 'range-dedupe-doc',
      replicaId: 'browser:test',
      persistState: false,
    })

    const cells = collectViewportCells(
      {
        sheetName: 'Sheet1',
        rowStart: 0,
        rowEnd: 1,
        colStart: 0,
        colEnd: 1,
      },
      {
        addresses: new Set(['A1']),
        positions: [{ address: 'A1', row: 0, col: 0 }],
      },
      [{ rowStart: 0, rowEnd: 0, colStart: 0, colEnd: 1 }],
    )

    expect(cells).toEqual([
      { address: 'A1', row: 0, col: 0 },
      { address: 'B1', row: 0, col: 1 },
    ])
  })

  it('collects changed cells without qualified address string round-trips', async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    })
    await runtime.bootstrap({
      documentId: 'cell-store-impact-doc',
      replicaId: 'browser:test',
      persistState: false,
    })

    await runtime.setCellValue('Sheet1', 'A1', 7)

    const engine = runtime['engine']
    if (!engine || !engine.workbook) {
      throw new Error('Expected bootstrapped engine')
    }

    engine.workbook.getQualifiedAddress = () => {
      throw new Error('collectChangedCellsBySheet should not use getQualifiedAddress')
    }

    const impacts = collectChangedCellsBySheet(engine, [0])

    expect(impacts.get('Sheet1')?.positions).toEqual([{ address: 'A1', row: 0, col: 0 }])
  })

  it('does not rewrite authoritative persistence for projected-only edit bursts', async () => {
    const persistProjectionState = vi.fn(async () => {})
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory({
        onPersistProjectionState: persistProjectionState,
      }),
    })

    await runtime.bootstrap({
      documentId: 'perf-doc',
      replicaId: 'browser:test',
      persistState: true,
    })

    expect(persistProjectionState).toHaveBeenCalledTimes(1)

    await runtime.setCellValue('Sheet1', 'A1', 1)
    await runtime.setCellValue('Sheet1', 'A2', 2)
    await runtime.setCellValue('Sheet1', 'A3', 3)

    expect(persistProjectionState).toHaveBeenCalledTimes(1)
  })

  it('reuses exported snapshots until the workbook changes', async () => {
    const runtime = new WorkbookWorkerRuntime({
      localStoreFactory: createMemoryLocalStoreFactory(),
    })
    await runtime.bootstrap({
      documentId: 'snapshot-cache-doc',
      replicaId: 'browser:test',
      persistState: false,
    })

    const first = runtime.exportSnapshot()
    const second = runtime.exportSnapshot()

    expect(second).toBe(first)

    await runtime.setCellValue('Sheet1', 'A1', 42)

    const third = runtime.exportSnapshot()
    const fourth = runtime.exportSnapshot()

    expect(third).not.toBe(first)
    expect(third.sheets[0]?.cells).toContainEqual(expect.objectContaining({ address: 'A1', value: 42 }))
    expect(fourth).toBe(third)
  })
})
