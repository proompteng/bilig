import { Profiler, useCallback, useEffect, useMemo, useRef } from 'react'
import { WorkbookView } from '@bilig/grid'
import type { BiligRuntimeConfig } from '@bilig/zero-sync'
import { resolveRuntimeConfig } from './runtime-config.js'
import type { ZeroClient } from './runtime-session.js'
import { parseSelectionTarget, type ZeroConnectionState } from './worker-workbook-app-model.js'
import { getWorkbookScrollPerfCollector } from './perf/workbook-scroll-perf.js'
import { readSelectionFromUrl, subscribeSelectionUrlChanges } from './selection-persistence.js'
import { WorkbookToastRegion } from './WorkbookToastRegion.js'
import { useWorkbookImportPane } from './use-workbook-import-pane.js'
import { useWorkbookSheetKeyboardShortcuts } from './use-workbook-sheet-keyboard-shortcuts.js'
import { useWorkbookShortcutDialog } from './use-workbook-shortcut-dialog.js'
import { useWorkerWorkbookAppState } from './use-worker-workbook-app-state.js'

function formatFailedPendingMutationMessage(input: { failureMessage: string }): string {
  return `A local change could not be synced. ${input.failureMessage}`
}

const missingSheetActionClass =
  'inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] px-3 text-[12px] font-medium text-[var(--wb-text)] shadow-[var(--wb-shadow-sm)] transition-colors hover:border-[var(--wb-border-strong)] hover:bg-[var(--wb-muted)] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1'

function sameWorkbookSelection(
  left: { readonly sheetName: string; readonly address: string },
  right: { readonly sheetName: string; readonly address: string },
): boolean {
  return left.sheetName === right.sheetName && left.address === right.address
}

function isResolvingWorkbookSheet(syncState: unknown): boolean {
  return syncState === 'syncing' || syncState === 'reconnecting'
}

function WorkbookResolvingState(props: { readonly requestedSheetName: string }) {
  return (
    <div className="flex h-full min-h-0 items-center justify-center bg-[var(--wb-surface)] px-6" data-testid="workbook-resolving-state">
      <div className="max-w-lg rounded-[var(--wb-radius-panel)] border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] p-5 shadow-[var(--wb-shadow-sm)]">
        <div className="text-[13px] font-semibold text-[var(--wb-text)]">Loading workbook</div>
        <div className="mt-2 text-[12px] leading-5 text-[var(--wb-text-muted)]">
          Resolving <span className="font-medium text-[var(--wb-text)]">{props.requestedSheetName}</span>.
        </div>
      </div>
    </div>
  )
}

function MissingSheetState(props: {
  readonly availableSheets: readonly string[]
  readonly requestedSheetName: string
  readonly onSelectSheet: (sheetName: string) => void
}) {
  return (
    <div className="flex h-full min-h-0 items-center justify-center bg-[var(--wb-surface)] px-6" data-testid="missing-sheet-state">
      <div className="max-w-lg rounded-[var(--wb-radius-panel)] border border-[var(--wb-border)] bg-[var(--wb-surface-subtle)] p-5 shadow-[var(--wb-shadow-sm)]">
        <div className="text-[13px] font-semibold text-[var(--wb-text)]">Sheet not found</div>
        <div className="mt-2 text-[12px] leading-5 text-[var(--wb-text-muted)]">
          Requested sheet <span className="font-medium text-[var(--wb-text)]">{props.requestedSheetName}</span> is not in this workbook.
        </div>
        {props.availableSheets.length > 0 ? (
          <div className="mt-4 flex flex-wrap gap-2">
            {props.availableSheets.map((sheetName) => (
              <button className={missingSheetActionClass} key={sheetName} onClick={() => props.onSelectSheet(sheetName)} type="button">
                Open {sheetName}
              </button>
            ))}
          </div>
        ) : null}
      </div>
    </div>
  )
}

export function WorkerWorkbookApp(props: { config: BiligRuntimeConfig; connectionState: ZeroConnectionState; zero?: ZeroClient }) {
  const runtimeConfig = useMemo(() => resolveRuntimeConfig(props.config), [props.config])
  const runtimeKey = [
    runtimeConfig.documentId,
    runtimeConfig.persistState ? 'persist' : 'memory',
    runtimeConfig.serverUrl ?? 'same-origin',
  ].join('|')

  return (
    <WorkerWorkbookAppInner
      key={runtimeKey}
      runtimeConfig={runtimeConfig}
      connectionState={props.connectionState}
      {...(props.zero ? { zero: props.zero } : {})}
    />
  )
}

function WorkerWorkbookAppInner({
  runtimeConfig,
  connectionState,
  zero,
}: {
  runtimeConfig: ReturnType<typeof resolveRuntimeConfig>
  connectionState: ZeroConnectionState
  zero?: ZeroClient
}) {
  const { clearImportError, importError, importPanel, importToggle } = useWorkbookImportPane({
    currentDocumentId: runtimeConfig.documentId,
    enabled: true,
  })
  const shortcuts = useWorkbookShortcutDialog()
  const app = useWorkerWorkbookAppState({
    runtimeConfig,
    connectionState,
    toolbarControls: (
      <>
        {shortcuts.shortcutHelpButton}
        {importToggle}
      </>
    ),
    ...(zero ? { zero } : {}),
  })
  const {
    agentError,
    clearAgentError,
    clearRuntimeError,
    failedPendingMutation,
    reportRuntimeError,
    retryFailedPendingMutation,
    runtimeError,
  } = app
  const benchmarkCorpus = useMemo(() => new URLSearchParams(window.location.search).get('benchmarkCorpus'), [])
  const installedBenchmarkCorpusRef = useRef<string | null>(null)
  const installingBenchmarkCorpusRef = useRef<string | null>(null)
  const { installBenchmarkCorpus, runtimeReady } = app

  useEffect(() => {
    getWorkbookScrollPerfCollector()
  }, [])

  useEffect(() => {
    if (!benchmarkCorpus || !runtimeReady) {
      return
    }
    if (installedBenchmarkCorpusRef.current === benchmarkCorpus) {
      return
    }
    if (installingBenchmarkCorpusRef.current === benchmarkCorpus) {
      return
    }
    const collector = getWorkbookScrollPerfCollector()
    if (collector?.getBenchmarkState().state === 'ready' && collector.getBenchmarkState().fixture?.id === benchmarkCorpus) {
      installedBenchmarkCorpusRef.current = benchmarkCorpus
      return
    }
    installingBenchmarkCorpusRef.current = benchmarkCorpus
    void (async () => {
      try {
        await installBenchmarkCorpus(benchmarkCorpus)
        installedBenchmarkCorpusRef.current = benchmarkCorpus
      } catch (error) {
        getWorkbookScrollPerfCollector()?.setBenchmarkState('error', error instanceof Error ? error.message : String(error))
        reportRuntimeError(error)
      } finally {
        if (installingBenchmarkCorpusRef.current === benchmarkCorpus) {
          installingBenchmarkCorpusRef.current = null
        }
      }
    })()
  }, [benchmarkCorpus, installBenchmarkCorpus, reportRuntimeError, runtimeReady])
  const reportAsyncError = useCallback(
    (task: Promise<unknown>): void => {
      void (async () => {
        try {
          await task
        } catch (error) {
          reportRuntimeError(error)
        }
      })()
    },
    [reportRuntimeError],
  )
  const toasts = useMemo(
    () =>
      [
        runtimeError
          ? {
              id: 'runtime-error',
              tone: 'error' as const,
              message: runtimeError,
              onDismiss: clearRuntimeError,
            }
          : null,
        failedPendingMutation
          ? {
              id: `pending-mutation-${failedPendingMutation.id}`,
              tone: 'error' as const,
              message: formatFailedPendingMutationMessage(failedPendingMutation),
              action: {
                label: 'Retry',
                onAction: () => {
                  reportAsyncError(Promise.resolve(retryFailedPendingMutation()))
                },
              },
            }
          : null,
        agentError
          ? {
              id: 'agent-error',
              tone: 'error' as const,
              message: agentError,
              onDismiss: clearAgentError,
            }
          : null,
        importError
          ? {
              id: 'import-error',
              tone: 'error' as const,
              message: importError,
              onDismiss: clearImportError,
            }
          : null,
      ].flatMap((toast) => (toast ? [toast] : [])),
    [
      agentError,
      clearAgentError,
      clearRuntimeError,
      failedPendingMutation,
      clearImportError,
      importError,
      reportAsyncError,
      retryFailedPendingMutation,
      runtimeError,
    ],
  )
  const visibleSelection = app.visibleSelection ?? app.selection ?? { sheetName: 'Sheet1', address: 'A1' }
  const visibleSelectedCell = app.visibleSelectedCell ?? app.selectedCell
  const visibleSelectionRef = useRef(visibleSelection)
  const selectAddressRef = useRef(app.selectAddress)
  visibleSelectionRef.current = visibleSelection
  selectAddressRef.current = app.selectAddress
  const appSelectAddress = app.selectAddress
  const appSheetNames = app.sheetNames
  const missingVisibleSheetFromWorkbook =
    app.workbookReady && app.workerHandle && appSheetNames.length > 0 && !appSheetNames.includes(visibleSelection.sheetName)
  const resolvingVisibleSheet = missingVisibleSheetFromWorkbook ? !runtimeError && isResolvingWorkbookSheet(app.runtimeSyncState) : false
  const missingVisibleSheet = missingVisibleSheetFromWorkbook && !runtimeError && !resolvingVisibleSheet

  useEffect(() => {
    if (!missingVisibleSheetFromWorkbook || !runtimeError) {
      return
    }
    const fallbackSheetName = appSheetNames[0]
    if (!fallbackSheetName) {
      return
    }
    appSelectAddress(fallbackSheetName, visibleSelection.address)
  }, [appSelectAddress, appSheetNames, missingVisibleSheetFromWorkbook, runtimeError, visibleSelection.address])

  const syncSelectionFromUrl = useCallback(() => {
    const nextSelection = readSelectionFromUrl()
    if (!nextSelection || sameWorkbookSelection(nextSelection, visibleSelectionRef.current)) {
      return
    }
    selectAddressRef.current(nextSelection.sheetName, nextSelection.address)
  }, [])

  useEffect(() => {
    syncSelectionFromUrl()
    return subscribeSelectionUrlChanges(syncSelectionFromUrl)
  }, [syncSelectionFromUrl])

  useWorkbookSheetKeyboardShortcuts({
    address: visibleSelection.address,
    enabled: app.workbookReady && Boolean(app.workerHandle),
    onSelectSheet: app.selectAddress,
    sheetName: visibleSelection.sheetName,
    sheetNames: app.sheetNames,
  })

  return (
    <div className="flex h-dvh min-h-0 flex-col overflow-hidden bg-[var(--wb-app-bg)] text-[var(--wb-text)]">
      {app.editorConflictBanner}
      <div className="relative flex min-h-0 flex-1">
        <WorkbookToastRegion toasts={toasts} />
        <div className="min-h-0 min-w-0 flex-1">
          {missingVisibleSheet ? (
            <MissingSheetState
              availableSheets={app.sheetNames}
              requestedSheetName={visibleSelection.sheetName}
              onSelectSheet={(sheetName) => app.selectAddress(sheetName, 'A1')}
            />
          ) : resolvingVisibleSheet ? (
            <WorkbookResolvingState requestedSheetName={visibleSelection.sheetName} />
          ) : app.workbookReady && app.workerHandle ? (
            <Profiler
              id="workbook-shell"
              onRender={() => {
                getWorkbookScrollPerfCollector()?.noteSurfaceCommit('workbookShell')
              }}
            >
              <WorkbookView
                ribbon={app.ribbon}
                editorValue={app.visibleEditorValue}
                editorTargetSelection={app.editorTargetSelection}
                editorSelectionBehavior={app.editorSelectionBehavior}
                engine={app.workerHandle.viewportStore}
                definedNames={app.definedNames}
                isEditing={Boolean(app.writesAllowed && app.isEditing)}
                isEditingCell={Boolean(app.writesAllowed && app.isEditingCell)}
                getCellEditorSeed={app.getCellEditorSeed}
                onAddressCommit={(input) => {
                  const nextTarget = parseSelectionTarget(input, visibleSelection.sheetName, app.definedNames)
                  if (!nextTarget) {
                    return false
                  }
                  app.selectSelectionSnapshot(nextTarget)
                  return true
                }}
                onAutofitColumn={(columnIndex: number, fallbackWidth: number) => {
                  return (async () => {
                    try {
                      await app.autofitColumn(visibleSelection.sheetName, columnIndex, fallbackWidth)
                    } catch (error) {
                      app.reportRuntimeError(error)
                    }
                  })()
                }}
                onBeginEdit={(seed, selectionBehavior, targetSelection) =>
                  app.beginEditing(seed, selectionBehavior, 'cell', targetSelection)
                }
                onBeginFormulaEdit={(seed?: string) => app.beginEditing(seed, 'select-all', 'formula')}
                onCancelEdit={app.cancelEditor}
                onClearCell={app.clearSelectedCell}
                onColumnWidthChange={(columnIndex: number, newSize: number) => {
                  reportAsyncError(
                    app.invokeColumnWidthMutation(visibleSelection.sheetName, columnIndex, newSize, {
                      deferLocalApplication: true,
                      deferPersistence: true,
                    }),
                  )
                }}
                onRowHeightChange={(rowIndex: number, newSize: number) => {
                  reportAsyncError(
                    app.invokeRowHeightMutation(visibleSelection.sheetName, rowIndex, newSize, {
                      deferLocalApplication: true,
                      deferPersistence: true,
                    }),
                  )
                }}
                onSetColumnHidden={(columnIndex: number, hidden: boolean) => {
                  reportAsyncError(app.invokeColumnVisibilityMutation(visibleSelection.sheetName, columnIndex, hidden))
                }}
                onInsertColumns={(startCol: number, count: number) => {
                  reportAsyncError(app.invokeInsertColumnsMutation(visibleSelection.sheetName, startCol, count))
                }}
                onDeleteColumns={(startCol: number, count: number) => {
                  const task = app.invokeDeleteColumnsMutation(visibleSelection.sheetName, startCol, count)
                  reportAsyncError(task)
                  return task
                }}
                onSetRowHidden={(rowIndex: number, hidden: boolean) => {
                  reportAsyncError(app.invokeRowVisibilityMutation(visibleSelection.sheetName, rowIndex, hidden))
                }}
                onInsertRows={(startRow: number, count: number) => {
                  reportAsyncError(app.invokeInsertRowsMutation(visibleSelection.sheetName, startRow, count))
                }}
                onDeleteRows={(startRow: number, count: number) => {
                  const task = app.invokeDeleteRowsMutation(visibleSelection.sheetName, startRow, count)
                  reportAsyncError(task)
                  return task
                }}
                onSetFreezePane={(rows: number, cols: number) => {
                  reportAsyncError(app.invokeSetFreezePaneMutation(visibleSelection.sheetName, rows, cols))
                }}
                onVisibleViewportChange={app.handleVisibleViewportChange}
                onCommitEdit={app.commitEditor}
                onCopyRange={app.copySelectionRange}
                onCreateSheet={app.writesAllowed ? app.createSheet : undefined}
                onDeleteSheet={app.writesAllowed ? app.deleteSheet : undefined}
                onEditorChange={app.handleEditorChange}
                onFillRange={app.fillSelectionRange}
                onMoveRange={app.moveSelectionRange}
                onPaste={app.pasteIntoSelection}
                previewRanges={app.previewRanges}
                onToggleBooleanCell={app.toggleBooleanCell}
                onRenameSheet={app.writesAllowed ? app.renameSheet : undefined}
                onSelectionChange={app.handleSelectionChange}
                onExternalSelectionSync={app.acknowledgeExternalSelectionSync}
                onSelectSheet={(sheetName) => app.selectAddress(sheetName, 'A1')}
                resolvedValue={app.resolvedValue}
                selectedAddr={visibleSelection.address}
                selectedCellSnapshot={visibleSelectedCell}
                selectionSnapshot={app.selectionSnapshot}
                sheetId={app.sheetIdsByName?.[visibleSelection.sheetName]}
                sheetOrdinal={app.sheetOrdinalsByName?.[visibleSelection.sheetName]}
                sheetName={visibleSelection.sheetName}
                sheetNames={app.sheetNames}
                renderTileSource={app.workerHandle?.viewportStore}
                columnWidths={app.columnWidths}
                hiddenColumns={app.hiddenColumns}
                hiddenRows={app.hiddenRows}
                rowHeights={app.rowHeights}
                freezeRows={app.freezeRows}
                freezeCols={app.freezeCols}
                onSidePanelWidthChange={app.setSidePanelWidth}
                sidePanelId={app.sidePanelId}
                sidePanel={app.sidePanel}
                sidePanelWidth={app.sidePanelWidth}
              />
            </Profiler>
          ) : null}
        </div>
        {importPanel}
        {shortcuts.shortcutDialog}
      </div>
    </div>
  )
}
