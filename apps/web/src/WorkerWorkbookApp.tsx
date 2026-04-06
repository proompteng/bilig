import { useMemo } from "react";
import { WorkbookView } from "@bilig/grid";
import type { BiligRuntimeConfig } from "@bilig/zero-sync";
import { resolveRuntimeConfig } from "./runtime-config.js";
import type { ZeroClient } from "./runtime-session.js";
import { parseSelectionTarget, type ZeroConnectionState } from "./worker-workbook-app-model.js";
import { useWorkbookImportPane } from "./use-workbook-import-pane.js";
import { useWorkerWorkbookAppState } from "./use-worker-workbook-app-state.js";

export function WorkerWorkbookApp(props: {
  config: BiligRuntimeConfig;
  connectionState: ZeroConnectionState;
  zero: ZeroClient;
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
      zero={props.zero}
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
  zero: ZeroClient;
}) {
  const app = useWorkerWorkbookAppState({ runtimeConfig, connectionState, zero });
  const { importPanel, importToggle } = useWorkbookImportPane({
    currentDocumentId: runtimeConfig.documentId,
    enabled: true,
  });

  return (
    <div className="flex h-screen flex-col overflow-hidden bg-[var(--wb-app-bg)] text-[var(--wb-text)]">
      {app.runtimeError ? (
        <div
          className="border-b border-[#f1b5b5] bg-[#fff7f7] px-3 py-2 text-sm text-[#991b1b]"
          data-testid="worker-error"
        >
          {app.runtimeError}
        </div>
      ) : null}
      {app.runtimeReady && !app.remoteSyncAvailable ? (
        <div className="border-b border-[var(--wb-accent-ring)] bg-[var(--wb-accent-soft)] px-3 py-2 text-sm text-[var(--wb-accent)]">
          Zero is {app.statusModeLabel.toLowerCase()}. Local edits remain available while sync is
          degraded.
        </div>
      ) : null}
      {app.editorConflictBanner}
      <div className="relative flex min-h-0 flex-1">
        <div className="min-h-0 min-w-0 flex-1">
          {app.workbookReady && app.workerHandle ? (
            <WorkbookView
              ribbon={app.ribbon}
              editorValue={app.visibleEditorValue}
              editorSelectionBehavior={app.editorSelectionBehavior}
              engine={app.workerHandle.viewportStore}
              isEditing={Boolean(app.writesAllowed && app.isEditing)}
              isEditingCell={Boolean(app.writesAllowed && app.isEditingCell)}
              onAddressCommit={(input) => {
                const nextTarget = parseSelectionTarget(input, app.selection.sheetName);
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
              onSelect={(addr) => app.selectAddress(app.selection.sheetName, addr)}
              onSelectSheet={(sheetName) => app.selectAddress(sheetName, "A1")}
              resolvedValue={app.resolvedValue}
              selectedAddr={app.selection.address}
              selectedCellSnapshot={app.selectedCell}
              selectionStatus={app.selectionStatus}
              sheetName={app.selection.sheetName}
              sheetNames={app.sheetNames}
              headerStatus={
                <>
                  {importToggle}
                  {app.headerStatus}
                </>
              }
              restoreViewportTarget={app.restoreViewportTarget}
              subscribeViewport={app.subscribeViewport}
              columnWidths={app.columnWidths}
              sideRail={app.agentPanel}
            />
          ) : null}
        </div>
        {app.versionsPanel}
        {app.viewsPanel}
        {app.changesPanel}
        {importPanel}
      </div>
    </div>
  );
}
