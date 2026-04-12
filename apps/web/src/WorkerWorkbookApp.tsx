import { useMemo } from "react";
import { WorkbookView } from "@bilig/grid";
import type { BiligRuntimeConfig } from "@bilig/zero-sync";
import { resolveRuntimeConfig } from "./runtime-config.js";
import type { ZeroClient } from "./runtime-session.js";
import { parseSelectionTarget, type ZeroConnectionState } from "./worker-workbook-app-model.js";
import { WorkbookToastRegion } from "./WorkbookToastRegion.js";
import { useWorkbookImportPane } from "./use-workbook-import-pane.js";
import { useWorkbookShortcutDialog } from "./use-workbook-shortcut-dialog.js";
import { useWorkerWorkbookAppState } from "./use-worker-workbook-app-state.js";

function formatFailedPendingMutationMessage(input: { failureMessage: string }): string {
  return `A local change could not be synced. ${input.failureMessage}`;
}

const persistenceBannerButtonClass =
  "inline-flex h-8 items-center rounded-[var(--wb-radius-control)] border border-[var(--wb-border-strong)] bg-[var(--wb-surface)] px-3 text-[12px] font-medium text-[var(--wb-text)] transition-colors hover:bg-[var(--wb-surface)]/80 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)] focus-visible:ring-offset-1";

export function WorkerWorkbookApp(props: {
  config: BiligRuntimeConfig;
  connectionState: ZeroConnectionState;
  zero?: ZeroClient;
}) {
  const runtimeConfig = useMemo(() => resolveRuntimeConfig(props.config), [props.config]);
  const runtimeKey = [
    runtimeConfig.documentId,
    runtimeConfig.persistState ? "persist" : "memory",
  ].join("|");

  return (
    <WorkerWorkbookAppInner
      key={runtimeKey}
      runtimeConfig={runtimeConfig}
      connectionState={props.connectionState}
      {...(props.zero ? { zero: props.zero } : {})}
    />
  );
}

function WorkerWorkbookAppInner({
  runtimeConfig,
  connectionState,
  zero,
}: {
  runtimeConfig: ReturnType<typeof resolveRuntimeConfig>;
  connectionState: ZeroConnectionState;
  zero?: ZeroClient;
}) {
  const app = useWorkerWorkbookAppState({
    runtimeConfig,
    connectionState,
    ...(zero ? { zero } : {}),
  });
  const { clearImportError, importError, importPanel, importToggle } = useWorkbookImportPane({
    currentDocumentId: runtimeConfig.documentId,
    enabled: true,
  });
  const shortcuts = useWorkbookShortcutDialog();
  const {
    agentError,
    clearAgentError,
    clearRuntimeError,
    failedPendingMutation,
    reportRuntimeError,
    retryFailedPendingMutation,
    runtimeError,
  } = app;
  const showFollowerPersistenceBanner =
    app.localPersistenceMode === "follower" && (app.transferRequested || !app.remoteSyncAvailable);
  const toasts = useMemo(
    () =>
      [
        runtimeError
          ? {
              id: "runtime-error",
              tone: "error" as const,
              message: runtimeError,
              onDismiss: clearRuntimeError,
            }
          : null,
        failedPendingMutation
          ? {
              id: `pending-mutation-${failedPendingMutation.id}`,
              tone: "error" as const,
              message: formatFailedPendingMutationMessage(failedPendingMutation),
              action: {
                label: "Retry",
                onAction: () => {
                  void Promise.resolve(retryFailedPendingMutation()).catch(reportRuntimeError);
                },
              },
            }
          : null,
        agentError
          ? {
              id: "agent-error",
              tone: "error" as const,
              message: agentError,
              onDismiss: clearAgentError,
            }
          : null,
        importError
          ? {
              id: "import-error",
              tone: "error" as const,
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
      reportRuntimeError,
      retryFailedPendingMutation,
      runtimeError,
    ],
  );

  return (
    <div className="flex h-screen flex-col overflow-hidden bg-[var(--wb-app-bg)] text-[var(--wb-text)]">
      {app.runtimeReady && app.zeroConfigured && !app.remoteSyncAvailable ? (
        <div className="border-b border-[var(--wb-accent-ring)] bg-[var(--wb-accent-soft)] px-3 py-2 text-sm text-[var(--wb-accent)]">
          Zero is {app.statusModeLabel.toLowerCase()}. Local edits remain available while sync is
          degraded.
        </div>
      ) : null}
      {showFollowerPersistenceBanner ? (
        <div className="border-b border-[var(--wb-border)] bg-[var(--wb-surface-muted)] px-3 py-2 text-sm text-[var(--wb-text-subtle)]">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div className="max-w-[72ch]">
              Another tab is the local writer.
              {app.transferRequested ? (
                <div className="mt-1 text-[12px] text-[var(--wb-text-muted)]">
                  Writer handoff requested. This tab will retry as soon as the writer releases the
                  local store.
                </div>
              ) : null}
            </div>
            <button
              className={persistenceBannerButtonClass}
              onClick={() => {
                app.requestPersistenceTransfer();
              }}
              type="button"
            >
              {app.transferRequested ? "Retry writer" : "Become writer"}
            </button>
          </div>
        </div>
      ) : null}
      {app.localPersistenceMode === "persistent" && app.pendingTransferRequest ? (
        <div className="border-b border-[var(--wb-border)] bg-[var(--wb-surface-muted)] px-3 py-2 text-sm text-[var(--wb-text-subtle)]">
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div className="max-w-[72ch]">
              Another tab wants writer ownership for local storage. If you transfer it, this tab
              stays live but loses offline persistence until it becomes the writer again.
            </div>
            <div className="flex flex-wrap items-center gap-2">
              <button
                className={persistenceBannerButtonClass}
                onClick={() => {
                  app.approvePersistenceTransfer();
                }}
                type="button"
              >
                Transfer writer
              </button>
              <button
                className={persistenceBannerButtonClass}
                onClick={() => {
                  app.dismissPersistenceTransferRequest();
                }}
                type="button"
              >
                Stay writer
              </button>
            </div>
          </div>
        </div>
      ) : null}
      {app.editorConflictBanner}
      <div className="relative flex min-h-0 flex-1">
        <WorkbookToastRegion toasts={toasts} />
        <div className="min-h-0 min-w-0 flex-1">
          {app.workbookReady && app.workerHandle ? (
            <WorkbookView
              ribbon={app.ribbon}
              editorValue={app.visibleEditorValue}
              editorSelectionBehavior={app.editorSelectionBehavior}
              engine={app.workerHandle.viewportStore}
              definedNames={app.definedNames}
              isEditing={Boolean(app.writesAllowed && app.isEditing)}
              isEditingCell={Boolean(app.writesAllowed && app.isEditingCell)}
              onAddressCommit={(input) => {
                const nextTarget = parseSelectionTarget(
                  input,
                  app.selection.sheetName,
                  app.definedNames,
                );
                if (nextTarget) {
                  app.selectAddress(nextTarget.sheetName, nextTarget.address);
                }
              }}
              onAutofitColumn={(columnIndex: number, fallbackWidth: number) => {
                return app
                  .invokeColumnWidthMutation(app.selection.sheetName, columnIndex, fallbackWidth, {
                    flush: true,
                  })
                  .then(() => undefined)
                  .catch(app.reportRuntimeError);
              }}
              onBeginEdit={app.beginEditing}
              onBeginFormulaEdit={(seed?: string) =>
                app.beginEditing(seed, "select-all", "formula")
              }
              onCancelEdit={app.cancelEditor}
              onClearCell={app.clearSelectedCell}
              onColumnWidthChange={(columnIndex: number, newSize: number) => {
                void app
                  .invokeColumnWidthMutation(app.selection.sheetName, columnIndex, newSize)
                  .catch(app.reportRuntimeError);
              }}
              onRowHeightChange={(rowIndex: number, newSize: number) => {
                void app
                  .invokeRowHeightMutation(app.selection.sheetName, rowIndex, newSize)
                  .catch(app.reportRuntimeError);
              }}
              onSetColumnHidden={(columnIndex: number, hidden: boolean) => {
                void app
                  .invokeColumnVisibilityMutation(app.selection.sheetName, columnIndex, hidden)
                  .catch(app.reportRuntimeError);
              }}
              onInsertColumns={(startCol: number, count: number) => {
                void app
                  .invokeInsertColumnsMutation(app.selection.sheetName, startCol, count)
                  .catch(app.reportRuntimeError);
              }}
              onDeleteColumns={(startCol: number, count: number) => {
                void app
                  .invokeDeleteColumnsMutation(app.selection.sheetName, startCol, count)
                  .catch(app.reportRuntimeError);
              }}
              onSetRowHidden={(rowIndex: number, hidden: boolean) => {
                void app
                  .invokeRowVisibilityMutation(app.selection.sheetName, rowIndex, hidden)
                  .catch(app.reportRuntimeError);
              }}
              onInsertRows={(startRow: number, count: number) => {
                void app
                  .invokeInsertRowsMutation(app.selection.sheetName, startRow, count)
                  .catch(app.reportRuntimeError);
              }}
              onDeleteRows={(startRow: number, count: number) => {
                void app
                  .invokeDeleteRowsMutation(app.selection.sheetName, startRow, count)
                  .catch(app.reportRuntimeError);
              }}
              onSetFreezePane={(rows: number, cols: number) => {
                void app
                  .invokeSetFreezePaneMutation(app.selection.sheetName, rows, cols)
                  .catch(app.reportRuntimeError);
              }}
              onVisibleViewportChange={app.handleVisibleViewportChange}
              onCommitEdit={app.commitEditor}
              onCopyRange={app.copySelectionRange}
              onCreateSheet={app.writesAllowed ? app.createSheet : undefined}
              onEditorChange={app.handleEditorChange}
              onFillRange={app.fillSelectionRange}
              onMoveRange={app.moveSelectionRange}
              onPaste={app.pasteIntoSelection}
              previewRanges={app.previewRanges}
              onToggleBooleanCell={app.toggleBooleanCell}
              onRenameSheet={app.writesAllowed ? app.renameSheet : undefined}
              onSelectionLabelChange={app.setSelectionLabel}
              onSelectionRangeChange={app.handleAgentSelectionRangeChange}
              onSelect={(addr) => app.selectAddress(app.selection.sheetName, addr)}
              onSelectSheet={(sheetName) => app.selectAddress(sheetName, "A1")}
              resolvedValue={app.resolvedValue}
              selectedAddr={app.selection.address}
              selectedCellSnapshot={app.selectedCell}
              selectionStatus={app.selectionStatus}
              sheetName={app.selection.sheetName}
              sheetNames={app.sheetNames}
              headerStatus={
                <div className="flex flex-wrap items-center justify-end gap-1.5">
                  {shortcuts.shortcutHelpButton}
                  {importToggle}
                  {app.headerStatus}
                </div>
              }
              subscribeViewport={app.subscribeViewport}
              columnWidths={app.columnWidths}
              hiddenColumns={app.hiddenColumns}
              hiddenRows={app.hiddenRows}
              rowHeights={app.rowHeights}
              freezeRows={app.freezeRows}
              freezeCols={app.freezeCols}
              onSideRailWidthChange={app.setSideRailWidth}
              sideRailId={app.sideRailId}
              sideRail={app.sideRail}
              sideRailWidth={app.sideRailWidth}
            />
          ) : null}
        </div>
        {importPanel}
        {shortcuts.shortcutDialog}
      </div>
    </div>
  );
}
