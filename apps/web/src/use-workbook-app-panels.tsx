import { useEffect, useMemo, useRef, type ReactNode } from "react";
import { Button } from "@base-ui/react/button";
import { Tabs } from "@base-ui/react/tabs";
import type { WorkbookAgentCommandBundle } from "@bilig/agent-api";
import type { WorkerRuntimeSelection } from "./runtime-session.js";
import {
  workbookHeaderActionButtonClass,
  workbookHeaderCountClass,
} from "./workbook-header-controls.js";
import { WorkbookPresenceBar } from "./WorkbookPresenceBar.js";
import {
  railCountClass,
  railIndicatorClass,
  railListClass,
  railPanelClass,
  railRootClass,
  railTabClass,
  type WorkbookSideRailTabDefinition,
} from "./WorkbookSideRailTabs.js";
import { cn } from "./cn.js";
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
  presenceClientId: string;
  replicaId: string;
  selection: WorkerRuntimeSelection;
  sheetNames: readonly string[];
  zero: WorkbookPanelsZeroSource;
  runtimeReady: boolean;
  zeroConfigured: boolean;
  remoteSyncAvailable: boolean;
  changeCount: number;
  changesPanel: ReactNode;
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
    presenceClientId,
    previewAgentBundle,
    remoteSyncAvailable,
    replicaId,
    runtimeReady,
    selection,
    selectAddress,
    sheetNames,
    zero,
    zeroConfigured,
  } = input;

  const collaborators = useWorkbookPresence({
    documentId,
    currentUserId,
    currentPresenceClientId: presenceClientId,
    sessionId: `${documentId}:${replicaId}`,
    selection,
    sheetNames,
    zero,
    enabled: runtimeReady && zeroConfigured && remoteSyncAvailable,
  });

  const {
    agentPanel,
    agentError,
    clearAgentError,
    pendingCommandCount,
    previewRanges,
    startNewThread,
  } = useWorkbookAgentPane({
    currentUserId,
    documentId,
    enabled: runtimeReady,
    getContext: getAgentContext,
    previewBundle: previewAgentBundle,
    zero,
    zeroEnabled: runtimeReady && zeroConfigured && remoteSyncAvailable,
  });

  const sideRailTabs = useMemo<readonly WorkbookSideRailTabDefinition[]>(
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
  const visibleSideRailTabs = useMemo(
    () => sideRailTabs.filter((tab) => tab.panel != null),
    [sideRailTabs],
  );
  const {
    activeSideRailTab,
    isSideRailOpen,
    openSideRail,
    setActiveSideRailTab,
    setSideRailWidth,
    sideRailWidth,
    toggleSideRail,
  } = useWorkbookShellLayout({
    documentId,
    persistenceKey: `${documentId}:${currentUserId}`,
    availableTabs: visibleSideRailTabs.map((tab) => tab.value),
    defaultTab: null,
  });
  const sideRailId = `workbook-side-rail-${documentId}`;
  const previousPendingCommandCountRef = useRef(pendingCommandCount);

  useEffect(() => {
    const hadPendingCommands = previousPendingCommandCountRef.current > 0;
    const hasPendingCommands = pendingCommandCount > 0;
    previousPendingCommandCountRef.current = pendingCommandCount;
    if (!hasPendingCommands || hadPendingCommands) {
      return;
    }
    if (!visibleSideRailTabs.some((tab) => tab.value === "assistant")) {
      return;
    }
    openSideRail("assistant");
  }, [openSideRail, pendingCommandCount, visibleSideRailTabs]);

  const sideRailToggleControls = useMemo(
    () => (
      <div
        className="inline-flex items-center gap-1 rounded-md border border-[var(--color-mauve-200)] bg-[var(--color-mauve-50)] p-1 shadow-[0_1px_2px_rgba(15,23,42,0.04)]"
        data-testid="workbook-side-rail-toggle-group"
      >
        {visibleSideRailTabs.map((tab) => {
          const active = isSideRailOpen && activeSideRailTab === tab.value;
          return (
            <Button
              aria-controls={sideRailId}
              aria-expanded={active}
              aria-pressed={active}
              className={workbookHeaderActionButtonClass({
                active,
                grouped: true,
              })}
              data-testid={`workbook-side-rail-toggle-${tab.value}`}
              key={tab.value}
              type="button"
              onClick={() => {
                toggleSideRail(tab.value);
              }}
            >
              <span>{tab.label}</span>
              {typeof tab.count === "number" ? (
                <span className={workbookHeaderCountClass}>{String(Math.min(tab.count, 99))}</span>
              ) : null}
            </Button>
          );
        })}
      </div>
    ),
    [activeSideRailTab, isSideRailOpen, sideRailId, toggleSideRail, visibleSideRailTabs],
  );

  const toolbarTrailingContent = useMemo(() => {
    if (visibleSideRailTabs.length === 0 && collaborators.length === 0) {
      return null;
    }
    return (
      <>
        {visibleSideRailTabs.length > 0 ? sideRailToggleControls : null}
        {collaborators.length > 0 ? (
          <WorkbookPresenceBar
            collaborators={collaborators}
            onJump={(sheetName, address) => {
              selectAddress(sheetName, address);
            }}
          />
        ) : null}
      </>
    );
  }, [collaborators, selectAddress, sideRailToggleControls, visibleSideRailTabs.length]);

  const sideRail = useMemo(
    () =>
      isSideRailOpen &&
      activeSideRailTab &&
      visibleSideRailTabs.some((tab) => tab.value === activeSideRailTab) ? (
        <Tabs.Root
          className={railRootClass()}
          value={activeSideRailTab}
          onValueChange={(nextValue) => {
            setActiveSideRailTab(String(nextValue));
          }}
        >
          <Tabs.List aria-label="Workbook panels" className={railListClass()}>
            <div className="flex min-w-0 flex-1 items-end gap-1">
              {visibleSideRailTabs.map((tab) => (
                <Tabs.Tab
                  className={(state) => railTabClass({ active: state.active })}
                  data-testid={`workbook-side-rail-tab-${tab.value}`}
                  key={tab.value}
                  value={tab.value}
                >
                  <span>{tab.label}</span>
                  {typeof tab.count === "number" ? (
                    <span
                      className={cn(
                        railCountClass({
                          active: activeSideRailTab === tab.value,
                        }),
                      )}
                    >
                      {String(Math.min(tab.count, 99))}
                    </span>
                  ) : null}
                </Tabs.Tab>
              ))}
            </div>
            <Button
              className={cn(
                workbookHeaderActionButtonClass({ active: false }),
                "mb-2 ml-auto shrink-0",
              )}
              data-testid="workbook-agent-new-thread"
              type="button"
              onClick={startNewThread}
            >
              New thread
            </Button>
            <Tabs.Indicator className={railIndicatorClass()} renderBeforeHydration />
          </Tabs.List>
          {visibleSideRailTabs.map((tab) => (
            <Tabs.Panel
              className={railPanelClass()}
              data-testid={`workbook-side-rail-panel-${tab.value}`}
              keepMounted
              key={tab.value}
              value={tab.value}
            >
              {tab.panel}
            </Tabs.Panel>
          ))}
        </Tabs.Root>
      ) : null,
    [activeSideRailTab, isSideRailOpen, setActiveSideRailTab, startNewThread, visibleSideRailTabs],
  );

  return {
    agentError,
    agentPanel,
    clearAgentError,
    pendingCommandCount,
    previewRanges,
    sideRailId,
    sideRail,
    setSideRailWidth,
    sideRailWidth,
    toolbarTrailingContent,
  };
}
