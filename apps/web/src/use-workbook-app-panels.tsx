import { useMemo, type ReactNode } from "react";
import type { WorkbookAgentCommandBundle } from "@bilig/agent-api";
import type { WorkerRuntimeSelection } from "./runtime-session.js";
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
      documentId,
      enabled: runtimeReady,
      getContext: getAgentContext,
      previewBundle: previewAgentBundle,
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
    availableTabs: sideRailTabs.map((tab) => tab.value),
    defaultTab: "assistant",
  });

  const sideRailToggleControls = useMemo(
    () => (
      <div
        className="inline-flex items-center gap-1 rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface)] p-1 shadow-[var(--wb-shadow-sm)]"
        data-testid="workbook-side-rail-toggle-group"
      >
        {sideRailTabs.map((tab) => {
          const active = isSideRailOpen && activeSideRailTab === tab.value;
          return (
            <button
              aria-pressed={active}
              className={[
                "inline-flex h-7 items-center gap-1.5 rounded-[calc(var(--wb-radius-control)-1px)] px-2.5 text-[12px] font-medium transition-colors",
                active
                  ? "bg-[var(--wb-surface-muted)] text-[var(--wb-text)]"
                  : "text-[var(--wb-text-subtle)] hover:bg-[var(--wb-hover)] hover:text-[var(--wb-text)]",
              ].join(" ")}
              data-testid={`workbook-side-rail-toggle-${tab.value}`}
              key={tab.value}
              type="button"
              onClick={() => {
                toggleSideRail(tab.value);
              }}
            >
              <span>{tab.label}</span>
              {typeof tab.count === "number" ? (
                <span className="inline-flex min-w-4 items-center justify-center rounded-full bg-[var(--wb-surface-subtle)] px-1.5 text-[10px] font-semibold leading-none text-[var(--wb-text-subtle)]">
                  {String(Math.min(tab.count, 99))}
                </span>
              ) : null}
            </button>
          );
        })}
      </div>
    ),
    [activeSideRailTab, isSideRailOpen, sideRailTabs, toggleSideRail],
  );

  const headerStatus = useMemo(
    () => (
      <div className="flex flex-wrap items-center justify-end gap-1.5">
        {toolbarHeaderStatus}
        {sideRailToggleControls}
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
    [collaborators, selectAddress, sideRailToggleControls, toolbarHeaderStatus],
  );

  const sideRail = useMemo(
    () =>
      isSideRailOpen && activeSideRailTab ? (
        <WorkbookSideRailTabs
          defaultValue="assistant"
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
    sideRail,
    setSideRailWidth,
    sideRailWidth,
  };
}
