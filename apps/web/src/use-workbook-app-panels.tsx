import { useMemo, type ReactNode } from "react";
import type { WorkbookAgentCommandBundle } from "@bilig/agent-api";
import type { WorkerRuntimeSelection } from "./runtime-session.js";
import {
  WorkbookHeaderActionButton,
  workbookHeaderCountClass,
} from "./workbook-header-controls.js";
import { WorkbookPresenceBar } from "./WorkbookPresenceBar.js";
import { WorkbookSideRailTabs } from "./WorkbookSideRailTabs.js";
import { useWorkbookAgentPane } from "./use-workbook-agent-pane.js";
import { useWorkbookPresence } from "./use-workbook-presence.js";
import { useWorkbookShellLayout } from "./use-workbook-shell-layout.js";

type WorkbookPanelsZeroSource = Parameters<typeof useWorkbookPresence>[0]["zero"];

type WorkbookAgentContextGetter = Parameters<typeof useWorkbookAgentPane>[0]["getContext"];
type WorkbookAgentPreviewBundle = (
  bundle: WorkbookAgentCommandBundle,
) => ReturnType<Parameters<typeof useWorkbookAgentPane>[0]["previewBundle"]>;

export function useWorkbookAppPanels(input: {
  documentId: string;
  currentUserId: string;
  replicaId: string;
  selection: WorkerRuntimeSelection;
  sheetNames: readonly string[];
  zero: WorkbookPanelsZeroSource;
  runtimeReady: boolean;
  zeroConfigured: boolean;
  remoteSyncAvailable: boolean;
  changeCount: number;
  changesPanel: ReactNode;
  toolbarHeaderStatus: ReactNode;
  selectAddress: (sheetName: string, address: string) => void;
  getAgentContext: WorkbookAgentContextGetter;
  previewAgentBundle: WorkbookAgentPreviewBundle;
}) {
  const {
    changeCount,
    changesPanel,
    currentUserId,
    documentId,
    getAgentContext,
    previewAgentBundle,
    remoteSyncAvailable,
    replicaId,
    runtimeReady,
    selection,
    selectAddress,
    sheetNames,
    toolbarHeaderStatus,
    zero,
    zeroConfigured,
  } = input;

  const collaborators = useWorkbookPresence({
    documentId,
    sessionId: `${documentId}:${replicaId}`,
    selection,
    sheetNames,
    zero,
    enabled: runtimeReady && zeroConfigured && remoteSyncAvailable,
  });

  const { agentPanel, agentError, clearAgentError, pendingCommandCount, previewRanges } =
    useWorkbookAgentPane({
      currentUserId,
      documentId,
      enabled: runtimeReady,
      getContext: getAgentContext,
      previewBundle: previewAgentBundle,
      zero,
      zeroEnabled: runtimeReady && zeroConfigured && remoteSyncAvailable,
    });

  const sideRailTabs = useMemo(
    () => [
      {
        value: "assistant",
        label: "Assistant",
        count: pendingCommandCount > 0 ? pendingCommandCount : undefined,
        panel: agentPanel,
      },
      {
        value: "changes",
        label: "Changes",
        count: changeCount > 0 ? changeCount : undefined,
        panel: changesPanel,
      },
    ],
    [agentPanel, changeCount, changesPanel, pendingCommandCount],
  );
  const {
    activeSideRailTab,
    isSideRailOpen,
    setActiveSideRailTab,
    setSideRailWidth,
    sideRailWidth,
    toggleSideRail,
  } = useWorkbookShellLayout({
    documentId,
    persistenceKey: `${documentId}:${currentUserId}`,
    availableTabs: sideRailTabs.map((tab) => tab.value),
    defaultTab: null,
  });
  const sideRailId = `workbook-side-rail-${documentId}`;

  const sideRailToggleControls = useMemo(
    () => (
      <div
        className="inline-flex items-center gap-1 rounded-md border border-[var(--color-mauve-200)] bg-[var(--color-mauve-50)] p-1 shadow-[0_1px_2px_rgba(15,23,42,0.04)]"
        data-testid="workbook-side-rail-toggle-group"
      >
        {sideRailTabs.map((tab) => {
          const active = isSideRailOpen && activeSideRailTab === tab.value;
          return (
            <WorkbookHeaderActionButton
              aria-controls={sideRailId}
              aria-expanded={active}
              aria-pressed={active}
              data-testid={`workbook-side-rail-toggle-${tab.value}`}
              isActive={active}
              isGrouped
              key={tab.value}
              onClick={() => {
                toggleSideRail(tab.value);
              }}
            >
              <span>{tab.label}</span>
              {typeof tab.count === "number" ? (
                <span className={workbookHeaderCountClass}>{String(Math.min(tab.count, 99))}</span>
              ) : null}
            </WorkbookHeaderActionButton>
          );
        })}
      </div>
    ),
    [activeSideRailTab, isSideRailOpen, sideRailId, sideRailTabs, toggleSideRail],
  );

  const headerStatus = useMemo(
    () => (
      <div className="flex flex-wrap items-center justify-end gap-1.5">
        {toolbarHeaderStatus}
        {!isSideRailOpen ? sideRailToggleControls : null}
        {collaborators.length > 0 ? (
          <WorkbookPresenceBar
            collaborators={collaborators}
            onJump={(sheetName, address) => {
              selectAddress(sheetName, address);
            }}
          />
        ) : null}
      </div>
    ),
    [collaborators, isSideRailOpen, selectAddress, sideRailToggleControls, toolbarHeaderStatus],
  );

  const sideRail = useMemo(
    () =>
      isSideRailOpen && activeSideRailTab ? (
        <WorkbookSideRailTabs
          tabs={sideRailTabs}
          value={activeSideRailTab}
          onValueChange={setActiveSideRailTab}
        />
      ) : null,
    [activeSideRailTab, isSideRailOpen, setActiveSideRailTab, sideRailTabs],
  );

  return {
    agentError,
    agentPanel,
    clearAgentError,
    headerStatus,
    pendingCommandCount,
    previewRanges,
    sideRailId,
    sideRail,
    setSideRailWidth,
    sideRailWidth,
  };
}
