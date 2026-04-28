import { useEffect, useMemo, useRef, type ReactNode } from 'react'
import { Button } from '@base-ui/react/button'
import { Popover } from '@base-ui/react/popover'
import { Tabs } from '@base-ui/react/tabs'
import { History, PanelRightClose, PanelRightOpen, Plus } from 'lucide-react'
import type { WorkbookAgentCommandBundle } from '@bilig/agent-api'
import type { WorkbookAgentThreadSummary } from '@bilig/contracts'
import type { WorkerRuntimeSelection } from './runtime-session.js'
import { workbookHeaderActionButtonClass } from './workbook-header-controls.js'
import { WorkbookPresenceBar } from './WorkbookPresenceBar.js'
import {
  formatPanelCount,
  panelCountClass,
  panelIndicatorClass,
  panelListClass,
  panelContentClass,
  panelRootClass,
  panelTabClass,
  type WorkbookSidePanelTabDefinition,
} from './WorkbookSidePanelTabs.js'
import { cn } from './cn.js'
import { useWorkbookAgentPane } from './use-workbook-agent-pane.js'
import { useWorkbookPresence } from './use-workbook-presence.js'
import { formatWorkbookCollaboratorLabel } from './workbook-presence-model.js'
import { useWorkbookShellLayout } from './use-workbook-shell-layout.js'

type WorkbookPanelsZeroSource = Parameters<typeof useWorkbookPresence>[0]['zero']

type WorkbookAgentContextGetter = Parameters<typeof useWorkbookAgentPane>[0]['getContext']
type WorkbookAgentPreviewCommandBundle = (
  bundle: WorkbookAgentCommandBundle,
) => ReturnType<Parameters<typeof useWorkbookAgentPane>[0]['previewCommandBundle']>

function summarizeThreadActivity(text: string | null): string | null {
  if (!text) {
    return null
  }
  const normalized = text.trim().replaceAll(/\s+/g, ' ')
  if (normalized.length === 0) {
    return null
  }
  return normalized.length <= 72 ? normalized : `${normalized.slice(0, 69)}...`
}

function formatThreadEntryCount(entryCount: number): string {
  return `${entryCount} ${entryCount === 1 ? 'item' : 'items'}`
}

function WorkbookAgentHistoryMenu(props: {
  readonly activeThreadId: string | null
  readonly threadSummaries: readonly WorkbookAgentThreadSummary[]
  readonly onSelectThread: (threadId: string) => void
}) {
  const previousThreadSummaries = props.threadSummaries.filter((summary) => summary.threadId !== props.activeThreadId)
  if (previousThreadSummaries.length === 0) {
    return null
  }

  return (
    <Popover.Root modal={false}>
      <Popover.Trigger
        aria-label="Previous conversations"
        className={cn(
          workbookHeaderActionButtonClass({ active: false, iconOnly: true }),
          'shrink-0 border-transparent bg-transparent shadow-none hover:bg-[var(--color-mauve-100)] hover:text-[var(--color-mauve-900)]',
        )}
        data-testid="workbook-agent-history-trigger"
        title="Previous conversations"
        type="button"
      >
        <History aria-hidden="true" className="size-4" strokeWidth={1.9} />
      </Popover.Trigger>
      <Popover.Portal>
        <Popover.Positioner align="end" className="z-[1100]" side="bottom" sideOffset={8}>
          <Popover.Popup
            aria-label="Previous conversations"
            className="w-[280px] rounded-xl border border-[var(--color-mauve-200)] bg-white p-1.5 shadow-[0_12px_32px_rgba(15,23,42,0.14)] outline-none"
            data-testid="workbook-agent-history-menu"
          >
            <div className="px-2 py-1 text-[11px] font-semibold uppercase tracking-[0.08em] text-[var(--color-mauve-500)]">
              Previous conversations
            </div>
            <div className="mt-1 flex max-h-[320px] flex-col gap-1 overflow-y-auto">
              {previousThreadSummaries.map((threadSummary) => {
                const latestActivity = summarizeThreadActivity(threadSummary.latestEntryText)
                return (
                  <Button
                    key={threadSummary.threadId}
                    aria-label={`Open previous conversation ${threadSummary.threadId}`}
                    className="flex w-full items-start gap-2 rounded-lg border border-transparent px-2.5 py-2 text-left transition-colors hover:border-[var(--color-mauve-200)] hover:bg-[var(--color-mauve-50)]"
                    data-testid={`workbook-agent-history-thread-${threadSummary.threadId}`}
                    type="button"
                    onClick={() => {
                      props.onSelectThread(threadSummary.threadId)
                    }}
                  >
                    <div className="min-w-0 flex-1">
                      <div className="flex flex-wrap items-center gap-x-2 gap-y-1">
                        <span className="text-[12px] font-semibold text-[var(--color-mauve-900)]">
                          {threadSummary.scope === 'shared' ? 'Shared' : 'Private'}
                        </span>
                        <span className="text-[11px] text-[var(--color-mauve-600)]">
                          {threadSummary.scope === 'shared' ? formatWorkbookCollaboratorLabel(threadSummary.ownerUserId) : 'Just you'}
                        </span>
                        <span className="text-[11px] text-[var(--color-mauve-500)]">
                          {formatThreadEntryCount(threadSummary.entryCount)}
                        </span>
                      </div>
                      {latestActivity ? (
                        <div className="mt-1 truncate text-[11px] text-[var(--color-mauve-600)]">{latestActivity}</div>
                      ) : null}
                    </div>
                  </Button>
                )
              })}
            </div>
          </Popover.Popup>
        </Popover.Positioner>
      </Popover.Portal>
    </Popover.Root>
  )
}

export function useWorkbookAppPanels(input: {
  documentId: string
  currentUserId: string
  presenceClientId: string
  replicaId: string
  selection: WorkerRuntimeSelection
  sheetNames: readonly string[]
  zero: WorkbookPanelsZeroSource
  runtimeReady: boolean
  zeroConfigured: boolean
  remoteSyncAvailable: boolean
  changeCount: number
  changesPanel: ReactNode
  selectAddress: (sheetName: string, address: string) => void
  getAgentContext: WorkbookAgentContextGetter
  applyAgentContext?: (context: ReturnType<WorkbookAgentContextGetter>) => void
  previewAgentCommandBundle: WorkbookAgentPreviewCommandBundle
}) {
  const {
    changeCount,
    changesPanel,
    currentUserId,
    documentId,
    applyAgentContext,
    getAgentContext,
    presenceClientId,
    previewAgentCommandBundle,
    remoteSyncAvailable,
    replicaId,
    runtimeReady,
    selection,
    selectAddress,
    sheetNames,
    zero,
    zeroConfigured,
  } = input

  const collaborators = useWorkbookPresence({
    documentId,
    currentUserId,
    currentPresenceClientId: presenceClientId,
    sessionId: `${documentId}:${replicaId}`,
    selection,
    sheetNames,
    zero,
    enabled: runtimeReady && zeroConfigured && remoteSyncAvailable,
  })

  const {
    activeThreadId,
    agentPanel,
    agentError,
    clearAgentError,
    pendingCommandCount,
    previewRanges,
    selectThread,
    startNewThread,
    threadSummaries,
  } = useWorkbookAgentPane({
    currentUserId,
    documentId,
    enabled: runtimeReady,
    getContext: getAgentContext,
    ...(applyAgentContext ? { applyContext: applyAgentContext } : {}),
    previewCommandBundle: previewAgentCommandBundle,
    zero,
    zeroEnabled: runtimeReady && zeroConfigured && remoteSyncAvailable,
  })

  const sidePanelTabs = useMemo<readonly WorkbookSidePanelTabDefinition[]>(
    () => [
      {
        value: 'assistant',
        label: 'Assistant',
        count: pendingCommandCount > 0 ? pendingCommandCount : undefined,
        panel: agentPanel,
      },
      {
        value: 'changes',
        label: 'Changes',
        count: changeCount > 0 ? changeCount : undefined,
        panel: changesPanel,
      },
    ],
    [agentPanel, changeCount, changesPanel, pendingCommandCount],
  )
  const visibleSidePanelTabs = useMemo(() => sidePanelTabs.filter((tab) => tab.panel != null), [sidePanelTabs])
  const { activeSidePanelTab, closeSidePanel, isSidePanelOpen, openSidePanel, setActiveSidePanelTab, setSidePanelWidth, sidePanelWidth } =
    useWorkbookShellLayout({
      documentId,
      persistenceKey: `${documentId}:${currentUserId}`,
      availableTabs: visibleSidePanelTabs.map((tab) => tab.value),
      defaultOpen: true,
      defaultTab: 'assistant',
    })
  const sidePanelId = `workbook-side-panel-${documentId}`
  const previousPendingCommandCountRef = useRef(pendingCommandCount)

  useEffect(() => {
    const hadPendingCommands = previousPendingCommandCountRef.current > 0
    const hasPendingCommands = pendingCommandCount > 0
    previousPendingCommandCountRef.current = pendingCommandCount
    if (!hasPendingCommands || hadPendingCommands) {
      return
    }
    if (!visibleSidePanelTabs.some((tab) => tab.value === 'assistant')) {
      return
    }
    setActiveSidePanelTab('assistant')
  }, [pendingCommandCount, setActiveSidePanelTab, visibleSidePanelTabs])

  const toolbarTrailingContent = useMemo(() => {
    const sidePanelTabToOpen =
      activeSidePanelTab && visibleSidePanelTabs.some((tab) => tab.value === activeSidePanelTab)
        ? activeSidePanelTab
        : visibleSidePanelTabs[0]?.value
    const sidePanelOpenButton =
      !isSidePanelOpen && sidePanelTabToOpen ? (
        <Button
          aria-label="Open workbook side panel"
          className={cn(
            workbookHeaderActionButtonClass({ active: false, iconOnly: true }),
            'shrink-0 border-transparent bg-transparent shadow-none hover:bg-[var(--color-mauve-100)] hover:text-[var(--color-mauve-900)]',
          )}
          data-testid="workbook-side-panel-open"
          title="Open side panel"
          type="button"
          onClick={() => openSidePanel(sidePanelTabToOpen)}
        >
          <PanelRightOpen aria-hidden="true" className="size-4" strokeWidth={1.9} />
        </Button>
      ) : null

    if (collaborators.length === 0 && !sidePanelOpenButton) {
      return null
    }
    return (
      <>
        {collaborators.length > 0 ? (
          <WorkbookPresenceBar
            collaborators={collaborators}
            onJump={(sheetName, address) => {
              selectAddress(sheetName, address)
            }}
          />
        ) : null}
        {sidePanelOpenButton}
      </>
    )
  }, [activeSidePanelTab, collaborators, isSidePanelOpen, openSidePanel, selectAddress, visibleSidePanelTabs])

  const sidePanel = useMemo(
    () =>
      isSidePanelOpen && activeSidePanelTab && visibleSidePanelTabs.some((tab) => tab.value === activeSidePanelTab) ? (
        <Tabs.Root
          className={panelRootClass()}
          value={activeSidePanelTab}
          onValueChange={(nextValue) => {
            setActiveSidePanelTab(String(nextValue))
          }}
        >
          <Tabs.List aria-label="Workbook panels" className={panelListClass()}>
            <div className="flex min-w-0 flex-1 items-end gap-1">
              {visibleSidePanelTabs.map((tab) => (
                <Tabs.Tab
                  className={(state) => panelTabClass({ active: state.active })}
                  data-testid={`workbook-side-panel-tab-${tab.value}`}
                  key={tab.value}
                  value={tab.value}
                >
                  <span>{tab.label}</span>
                  {typeof tab.count === 'number' ? (
                    <span
                      className={cn(
                        panelCountClass({
                          active: activeSidePanelTab === tab.value,
                        }),
                      )}
                    >
                      {formatPanelCount(tab.count)}
                    </span>
                  ) : null}
                </Tabs.Tab>
              ))}
            </div>
            {activeSidePanelTab === 'assistant' ? (
              <WorkbookAgentHistoryMenu activeThreadId={activeThreadId} threadSummaries={threadSummaries} onSelectThread={selectThread} />
            ) : null}
            <Button
              aria-label="New thread"
              className={cn(
                workbookHeaderActionButtonClass({ active: false, iconOnly: true }),
                'ml-auto shrink-0 self-center border-transparent bg-[var(--color-mauve-50)] shadow-none hover:bg-[var(--color-mauve-100)] hover:text-[var(--color-mauve-900)]',
              )}
              data-testid="workbook-agent-new-thread"
              title="New thread"
              type="button"
              disabled={activeSidePanelTab !== 'assistant'}
              onClick={startNewThread}
            >
              <Plus aria-hidden="true" className="size-4" strokeWidth={1.9} />
            </Button>
            <Button
              aria-label="Close workbook side panel"
              className={cn(
                workbookHeaderActionButtonClass({ active: false, iconOnly: true }),
                'shrink-0 border-transparent bg-transparent shadow-none hover:bg-[var(--color-mauve-100)] hover:text-[var(--color-mauve-900)]',
              )}
              data-testid="workbook-side-panel-close"
              title="Close side panel"
              type="button"
              onClick={closeSidePanel}
            >
              <PanelRightClose aria-hidden="true" className="size-4" strokeWidth={1.9} />
            </Button>
            <Tabs.Indicator className={panelIndicatorClass()} renderBeforeHydration />
          </Tabs.List>
          {visibleSidePanelTabs.map((tab) => (
            <Tabs.Panel
              className={panelContentClass()}
              data-testid={`workbook-side-panel-panel-${tab.value}`}
              keepMounted
              key={tab.value}
              value={tab.value}
            >
              {tab.panel}
            </Tabs.Panel>
          ))}
        </Tabs.Root>
      ) : null,
    [
      activeSidePanelTab,
      activeThreadId,
      closeSidePanel,
      isSidePanelOpen,
      selectThread,
      setActiveSidePanelTab,
      startNewThread,
      threadSummaries,
      visibleSidePanelTabs,
    ],
  )

  return {
    agentError,
    agentPanel,
    clearAgentError,
    pendingCommandCount,
    previewRanges,
    sidePanelId,
    sidePanel,
    setSidePanelWidth,
    sidePanelWidth,
    toolbarTrailingContent,
  }
}
