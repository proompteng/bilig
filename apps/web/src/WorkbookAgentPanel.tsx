import { Fragment, type CSSProperties, useEffect, useLayoutEffect, useRef } from 'react'
import { Button } from '@base-ui/react/button'
import { ScrollArea } from '@base-ui/react/scroll-area'
import { ArrowUp, Square } from 'lucide-react'
import { describeWorkbookAgentCommand } from '@bilig/agent-api'
import { cva } from 'class-variance-authority'
import type {
  WorkbookAgentCommandBundle,
  WorkbookAgentPreviewChangeKind,
  WorkbookAgentPreviewSummary,
  WorkbookAgentSharedReviewRecommendation,
} from '@bilig/agent-api'
import type {
  WorkbookAgentThreadSnapshot,
  WorkbookAgentTimelineCitation,
  WorkbookAgentThreadSummary,
  WorkbookAgentTimelineEntry,
  WorkbookAgentWorkflowRun,
} from '@bilig/contracts'
import { cn } from './cn.js'
import {
  workbookAlertClass,
  workbookButtonClass,
  workbookInsetClass,
  workbookPillClass,
  workbookSurfaceClass,
} from './workbook-shell-chrome.js'
import { WorkbookAgentDisclosureRow } from './workbook-agent-panel-disclosure-row.js'
import {
  agentPanelBodyMutedTextClass,
  agentPanelBodyTextClass,
  agentPanelComposerFrameClass,
  agentPanelComposerScrollContentClass,
  agentPanelComposerScrollRootClass,
  agentPanelComposerScrollViewportClass,
  agentPanelComposerSendButtonClass,
  agentPanelComposerTextareaClass,
  agentPanelEyebrowTextClass,
  agentPanelFooterClass,
  agentPanelLabelTextClass,
  agentPanelMetaTextClass,
  agentPanelScrollAreaContentClass,
  agentPanelScrollAreaRootClass,
  agentPanelScrollAreaScrollbarClass,
  agentPanelScrollAreaThumbClass,
  agentPanelScrollAreaViewportClass,
  agentPanelTimelineListClass,
  agentPanelThreadButtonClass,
  agentPanelThreadListClass,
} from './workbook-agent-panel-primitives.js'
import { renderToolDisplayName, safeParseToolOutput, StructuredToolOutput, summarizeToolEntry } from './workbook-agent-tool-output.js'
import { WorkbookAgentMarkdown } from './workbook-agent-markdown.js'
import { formatWorkbookCollaboratorLabel } from './workbook-presence-model.js'
import { AssistantProgressRow, PreviewRangeList, WorkflowRunRow } from './workbook-agent-panel-history.js'

const toolStatusPillClass = cva(
  'inline-flex h-5 items-center rounded-full border px-2 text-[10px] leading-none font-medium uppercase tracking-[0.05em]',
  {
    variants: {
      status: {
        completed: workbookPillClass({ tone: 'neutral', weight: 'strong' }),
        failed: workbookPillClass({ tone: 'danger', weight: 'strong' }),
        running: workbookPillClass({ tone: 'accent', weight: 'strong' }),
      },
    },
  },
)

const agentPanelThemeStyle: CSSProperties & Record<`--${string}`, string> = {
  '--wb-app-bg': 'var(--color-mauve-50)',
  '--wb-surface': 'white',
  '--wb-surface-subtle': 'var(--color-mauve-50)',
  '--wb-surface-muted': 'var(--color-mauve-100)',
  '--wb-border': 'var(--color-mauve-200)',
  '--wb-border-strong': 'var(--color-mauve-300)',
  '--wb-grid-border': 'var(--color-mauve-100)',
  '--wb-text': 'var(--color-mauve-900)',
  '--wb-text-muted': 'var(--color-mauve-700)',
  '--wb-text-subtle': 'var(--color-mauve-600)',
  '--wb-accent': 'var(--color-mauve-900)',
  '--wb-accent-soft': 'var(--color-mauve-100)',
  '--wb-accent-ring': 'var(--color-mauve-400)',
  '--wb-hover': 'var(--color-mauve-100)',
  '--wb-shadow-sm': '0 1px 2px rgba(15, 23, 42, 0.04)',
}

const AGENT_COMPOSER_MIN_HEIGHT = 112
const AGENT_COMPOSER_MAX_HEIGHT = 224

function formatThreadEntryCount(entryCount: number): string {
  return `${entryCount} ${entryCount === 1 ? 'item' : 'items'}`
}

function summarizeThreadActivity(text: string | null): string | null {
  if (!text) {
    return null
  }
  const normalized = text.trim().replaceAll(/\s+/g, ' ')
  if (normalized.length === 0) {
    return null
  }
  return normalized.length <= 64 ? normalized : `${normalized.slice(0, 61)}...`
}

function ThreadSummaryStrip(props: {
  readonly activeThreadId: string | null
  readonly threadSummaries: readonly WorkbookAgentThreadSummary[]
  readonly onSelectThread: (threadId: string) => void
}) {
  const visibleThreadSummaries = props.threadSummaries.filter((threadSummary) => threadSummary.threadId !== props.activeThreadId)
  if (visibleThreadSummaries.length === 0) {
    return null
  }

  return (
    <div className={agentPanelThreadListClass()}>
      {visibleThreadSummaries.map((threadSummary) => {
        const latestActivity = summarizeThreadActivity(threadSummary.latestEntryText)
        return (
          <Button
            key={threadSummary.threadId}
            aria-label={`Open ${threadSummary.scope} thread ${threadSummary.threadId}`}
            aria-pressed={false}
            className={agentPanelThreadButtonClass({ active: false })}
            data-testid={`workbook-agent-thread-${threadSummary.threadId}`}
            type="button"
            onClick={() => {
              props.onSelectThread(threadSummary.threadId)
            }}
          >
            <div className="min-w-0 flex-1">
              <div className="flex items-center gap-2">
                <span className={cn(agentPanelLabelTextClass(), 'font-semibold')}>
                  {threadSummary.scope === 'shared' ? 'Shared' : 'Private'}
                </span>
                <span className={agentPanelMetaTextClass()}>
                  {threadSummary.scope === 'shared' ? formatWorkbookCollaboratorLabel(threadSummary.ownerUserId) : 'Just you'}
                </span>
                <span className={agentPanelMetaTextClass()}>{formatThreadEntryCount(threadSummary.entryCount)}</span>
              </div>
              {latestActivity ? <div className={cn(agentPanelMetaTextClass(), 'mt-0.5 truncate')}>{latestActivity}</div> : null}
            </div>
            {threadSummary.scope === 'shared' && threadSummary.reviewQueueItemCount > 0 ? (
              <span className={workbookPillClass({ tone: 'accent', weight: 'strong' })}>Review</span>
            ) : null}
          </Button>
        )
      })}
    </div>
  )
}

function ToolStatusPill(props: { readonly status: WorkbookAgentTimelineEntry['toolStatus'] }) {
  if (props.status === 'completed') {
    return null
  }
  const label = props.status === 'failed' ? 'Failed' : 'Running'
  return (
    <span
      className={toolStatusPillClass({
        status: props.status === 'failed' ? 'failed' : 'running',
      })}
    >
      {label}
    </span>
  )
}

function renderPreviewChangeKind(kind: WorkbookAgentPreviewChangeKind): string {
  switch (kind) {
    case 'input':
      return 'value'
    case 'formula':
      return 'formula'
    case 'style':
      return 'style'
    case 'numberFormat':
      return 'number format'
  }
}

function renderMarkdownPlainText(markdown: string): string {
  return markdown
    .replaceAll(/```[\s\S]*?```/g, ' ')
    .replaceAll(/`([^`]+)`/g, '$1')
    .replaceAll(/\[([^\]]+)\]\(([^)]+)\)/g, '$1')
    .replaceAll(/!\[([^\]]*)\]\(([^)]+)\)/g, '$1')
    .replaceAll(/^#{1,6}\s+/gm, '')
    .replaceAll(/^>\s?/gm, '')
    .replaceAll(/[*_~]+/g, ' ')
    .replaceAll(/\s+/g, ' ')
}

function summarizeDisclosureText(text: string | null): string | null {
  if (!text) {
    return null
  }
  const normalized = renderMarkdownPlainText(text).trim().replaceAll(/\s+/g, ' ')
  if (normalized.length === 0) {
    return null
  }
  return normalized.length <= 88 ? normalized : `${normalized.slice(0, 85)}...`
}

function isAppliedExecutionSystemEntry(entry: WorkbookAgentTimelineEntry): boolean {
  return (
    entry.kind === 'system' &&
    (entry.text?.startsWith('Applied workbook change set at revision r') === true ||
      entry.text?.startsWith('Applied automatically workbook change set at revision r') === true ||
      entry.text?.startsWith('Applied automatically selected workbook change set at revision r') === true ||
      entry.text?.startsWith('Applied selected workbook change set at revision r') === true)
  )
}

function TextDisclosureEntryRow(props: { readonly entry: WorkbookAgentTimelineEntry; readonly label: 'Thought' | 'Plan' }) {
  const bodyText = props.entry.kind === 'reasoning' || props.entry.kind === 'plan' ? props.entry.text : null
  if (!bodyText?.trim().length) {
    return null
  }
  const summary = summarizeDisclosureText(bodyText)
  const disclosureKey = props.entry.kind

  return (
    <WorkbookAgentDisclosureRow
      id={props.entry.id}
      label={props.label}
      panelTestId={`workbook-agent-${disclosureKey}-panel-${props.entry.id}`}
      summary={summary ?? `No ${props.label.toLowerCase()} details available.`}
      triggerLabel={{
        expanded: `Collapse ${props.label.toLowerCase()}`,
        collapsed: `Expand ${props.label.toLowerCase()}`,
      }}
      triggerTestId={`workbook-agent-${disclosureKey}-toggle-${props.entry.id}`}
    >
      <div className={agentPanelBodyMutedTextClass()}>
        <WorkbookAgentMarkdown markdown={bodyText} />
      </div>
      <TimelineCitationList citations={props.entry.citations} />
    </WorkbookAgentDisclosureRow>
  )
}

function WorkbookAgentEntryRow(props: { readonly entry: WorkbookAgentTimelineEntry }) {
  const { entry } = props
  if (entry.kind === 'user') {
    return (
      <div className="flex justify-end px-3 py-3">
        <div
          className={cn(
            agentPanelBodyTextClass(),
            'max-w-[82%] break-words rounded-[var(--wb-radius-control)] border border-[var(--wb-border)] bg-[var(--wb-surface-muted)] px-3 py-2 shadow-[0_1px_2px_rgba(15,23,42,0.04)]',
          )}
        >
          <WorkbookAgentMarkdown markdown={entry.text ?? ''} />
          <TimelineCitationList citations={entry.citations} />
        </div>
      </div>
    )
  }

  if (entry.kind === 'assistant') {
    if (entry.phase === 'progress') {
      return null
    }
    if (!entry.text?.trim().length) {
      return null
    }
    return (
      <div className={cn(agentPanelBodyTextClass(), 'px-3 py-3')}>
        <WorkbookAgentMarkdown markdown={entry.text} />
        <TimelineCitationList citations={entry.citations} />
      </div>
    )
  }

  if (entry.kind === 'reasoning') {
    return <TextDisclosureEntryRow entry={entry} label="Thought" />
  }

  if (entry.kind === 'plan') {
    return <TextDisclosureEntryRow entry={entry} label="Plan" />
  }

  if (entry.kind === 'tool') {
    const displayName = renderToolDisplayName(entry.toolName)
    const summary = summarizeToolEntry(entry)
    const parsedOutput = safeParseToolOutput(entry.outputText)
    const hasDetails = (entry.argumentsText?.trim().length ?? 0) > 0 || (entry.outputText?.trim().length ?? 0) > 0
    if (!hasDetails) {
      return (
        <div className="px-3 py-2.5">
          <div className="flex items-center justify-between gap-3">
            <div className="min-w-0 flex flex-1 items-center gap-2.5">
              <div className={cn(agentPanelLabelTextClass(), 'min-w-0')}>{displayName}</div>
              {summary ? (
                <div className={cn(agentPanelMetaTextClass(), 'min-w-0 flex-1 whitespace-normal break-words')}>{summary}</div>
              ) : null}
            </div>
            <ToolStatusPill status={entry.toolStatus} />
          </div>
          <TimelineCitationList citations={entry.citations} />
        </div>
      )
    }
    return (
      <WorkbookAgentDisclosureRow
        badge={<ToolStatusPill status={entry.toolStatus} />}
        id={entry.id}
        label={displayName}
        panelTestId={`workbook-agent-tool-panel-${entry.id}`}
        summary={summary}
        triggerLabel={{
          expanded: `Collapse ${displayName}`,
          collapsed: `Expand ${displayName}`,
        }}
        triggerTestId={`workbook-agent-tool-toggle-${entry.id}`}
      >
        {entry.argumentsText?.trim().length ? (
          <div>
            <div className={agentPanelEyebrowTextClass()}>Arguments</div>
            <pre
              className={cn(
                agentPanelMetaTextClass(),
                'mt-1 box-border w-full min-w-0 max-w-full overflow-x-hidden whitespace-pre-wrap break-all rounded-[var(--wb-radius-control)] bg-[var(--wb-surface-subtle)] px-2 py-2',
              )}
            >
              {entry.argumentsText}
            </pre>
          </div>
        ) : null}
        {entry.outputText?.trim().length ? (
          <div className={entry.argumentsText?.trim().length ? 'mt-2' : undefined}>
            <div className={agentPanelEyebrowTextClass()}>Output</div>
            {parsedOutput !== null ? (
              <StructuredToolOutput toolName={entry.toolName} outputText={entry.outputText} />
            ) : (
              <pre
                className={cn(
                  agentPanelMetaTextClass(),
                  'mt-1 box-border w-full min-w-0 max-w-full overflow-x-hidden whitespace-pre-wrap break-all rounded-[var(--wb-radius-control)] bg-[var(--wb-surface-subtle)] px-2 py-2',
                )}
              >
                {entry.outputText}
              </pre>
            )}
          </div>
        ) : null}
        <TimelineCitationList citations={entry.citations} />
      </WorkbookAgentDisclosureRow>
    )
  }

  if (isAppliedExecutionSystemEntry(entry)) {
    return null
  }

  return (
    <div className="px-3 py-2.5">
      <div className={agentPanelMetaTextClass()}>{entry.text}</div>
      <TimelineCitationList citations={entry.citations} />
    </div>
  )
}

function TimelineCitationList(props: { readonly citations: readonly WorkbookAgentTimelineCitation[] }) {
  const segments = summarizeTimelineCitations(props.citations)
  if (segments.length === 0) {
    return null
  }
  return (
    <div className={cn(agentPanelMetaTextClass(), 'mt-1 break-words')}>
      {segments.map((segment, index) => (
        <Fragment key={segment}>
          {index > 0 ? <span aria-hidden="true"> · </span> : null}
          <span>{segment}</span>
        </Fragment>
      ))}
    </div>
  )
}

function summarizeTimelineCitations(citations: readonly WorkbookAgentTimelineCitation[]): readonly string[] {
  const seen = new Set<string>()
  const segments: string[] = []
  for (const citation of citations) {
    if (citation.kind === 'revision') {
      continue
    }
    const address =
      citation.startAddress === citation.endAddress
        ? `${citation.sheetName}!${citation.startAddress}`
        : `${citation.sheetName}!${citation.startAddress}:${citation.endAddress}`
    const segment = `${citation.role === 'target' ? 'Target' : 'Source'} ${address}`
    if (seen.has(segment)) {
      continue
    }
    seen.add(segment)
    segments.push(segment)
  }
  return segments
}

function ReviewItemCard(props: {
  readonly activeReviewBundle: WorkbookAgentCommandBundle
  readonly preview: WorkbookAgentPreviewSummary | null
  readonly sharedApprovalOwnerUserId: string | null
  readonly sharedReviewOwnerUserId: string | null
  readonly sharedReviewStatus: 'pending' | 'approved' | 'rejected' | null
  readonly sharedReviewDecidedByUserId: string | null
  readonly sharedReviewRecommendations: readonly WorkbookAgentSharedReviewRecommendation[]
  readonly currentUserSharedRecommendation: 'approved' | 'rejected' | null
  readonly canFinalizeSharedBundle: boolean
  readonly canRecommendSharedBundle: boolean
  readonly selectedCommandIndexes: readonly number[]
  readonly isApplyingReviewItem: boolean
  readonly onApply: () => void
  readonly onDismiss: () => void
  readonly onReview: (decision: 'approved' | 'rejected') => void
  readonly onSelectAll: () => void
  readonly onToggleCommand: (commandIndex: number) => void
}) {
  const selectedCount = props.selectedCommandIndexes.length
  const hasFullSelection = selectedCount === props.activeReviewBundle.commands.length
  const sharedApprovalOwnerLabel = props.sharedApprovalOwnerUserId ? formatWorkbookCollaboratorLabel(props.sharedApprovalOwnerUserId) : null
  const sharedReviewOwnerLabel = props.sharedReviewOwnerUserId ? formatWorkbookCollaboratorLabel(props.sharedReviewOwnerUserId) : null
  const sharedReviewDecisionLabel = props.sharedReviewDecidedByUserId
    ? formatWorkbookCollaboratorLabel(props.sharedReviewDecidedByUserId)
    : null
  const approvalRecommendationCount = props.sharedReviewRecommendations.filter(
    (recommendation) => recommendation.decision === 'approved',
  ).length
  const rejectionRecommendationCount = props.sharedReviewRecommendations.filter(
    (recommendation) => recommendation.decision === 'rejected',
  ).length
  const recommendationSummary =
    props.sharedReviewRecommendations.length === 0
      ? null
      : `${String(approvalRecommendationCount)} approval ${approvalRecommendationCount === 1 ? 'recommendation' : 'recommendations'} · ${String(rejectionRecommendationCount)} rejection ${rejectionRecommendationCount === 1 ? 'recommendation' : 'recommendations'}`
  const canApply =
    props.preview !== null &&
    !props.isApplyingReviewItem &&
    selectedCount > 0 &&
    sharedApprovalOwnerLabel === null &&
    (props.sharedReviewStatus === null || props.sharedReviewStatus === 'approved')
  const applyLabel =
    selectedCount > 0 && !hasFullSelection
      ? props.isApplyingReviewItem
        ? 'Applying…'
        : 'Apply'
      : props.sharedReviewStatus === 'pending'
        ? 'Owner review'
        : props.sharedReviewStatus === 'rejected'
          ? 'Returned'
          : props.isApplyingReviewItem
            ? 'Applying…'
            : 'Apply'
  return (
    <div className={cn(workbookSurfaceClass({ emphasis: 'raised' }), 'border-[var(--wb-border-strong)] px-3 py-3')}>
      <div className={cn(agentPanelLabelTextClass(), 'font-semibold')}>{props.activeReviewBundle.summary}</div>
      <div className={cn(workbookInsetClass(), 'mt-3 px-2 py-2')}>
        <div className="flex items-center justify-between gap-3">
          <div className={agentPanelEyebrowTextClass()}>
            {String(selectedCount)}/{String(props.activeReviewBundle.commands.length)}
          </div>
          {!hasFullSelection ? (
            <Button
              className="text-[12px] leading-5 font-medium text-[var(--wb-accent)] transition-colors hover:brightness-[0.95] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[var(--wb-accent-ring)]"
              type="button"
              onClick={props.onSelectAll}
            >
              All
            </Button>
          ) : null}
        </div>
        <div className="mt-2 flex flex-col gap-2">
          {props.activeReviewBundle.commands.map((command, index) => {
            const checked = props.selectedCommandIndexes.includes(index)
            const commandLabel = describeWorkbookAgentCommand(command)
            return (
              <div
                key={`${props.activeReviewBundle.id}:${JSON.stringify(command)}`}
                className={cn(
                  'flex items-start gap-3 rounded-[var(--wb-radius-control)] border px-3 py-2 transition-colors',
                  checked ? 'border-[var(--wb-accent-ring)] bg-[var(--wb-surface)]' : 'border-[var(--wb-border)] bg-[var(--wb-surface)]',
                )}
              >
                <input
                  aria-label={`Toggle workbook review item change ${String(index + 1)}: ${commandLabel}`}
                  checked={checked}
                  className="mt-0.5 h-4 w-4 rounded border-[var(--wb-border)] text-[var(--wb-accent)] focus:ring-[var(--wb-accent-ring)]"
                  data-testid={`workbook-agent-review-command-toggle-${String(index)}`}
                  type="checkbox"
                  onChange={() => {
                    props.onToggleCommand(index)
                  }}
                />
                <div className="min-w-0">
                  <div className={agentPanelLabelTextClass()}>{commandLabel}</div>
                </div>
              </div>
            )
          })}
        </div>
      </div>
      <PreviewRangeList ranges={props.preview?.ranges ?? props.activeReviewBundle.affectedRanges} />
      {props.preview?.structuralChanges?.length ? (
        <div className={cn(workbookInsetClass(), agentPanelMetaTextClass(), 'mt-2 border-transparent px-2 py-2')}>
          {props.preview.structuralChanges.join(' · ')}
        </div>
      ) : null}
      {sharedApprovalOwnerLabel ? (
        <div className={cn(workbookAlertClass({ tone: 'warning' }), agentPanelMetaTextClass(), 'mt-2 border-[var(--wb-border-strong)]')}>
          Owner review routes medium/high-risk changes to {sharedApprovalOwnerLabel} on this shared thread.
        </div>
      ) : null}
      {recommendationSummary ? (
        <div className={cn(workbookInsetClass(), agentPanelMetaTextClass(), 'mt-2 border-transparent px-2 py-2')}>
          {recommendationSummary}
          {props.currentUserSharedRecommendation
            ? ` You recommended ${props.currentUserSharedRecommendation === 'approved' ? 'approval' : 'rejection'}.`
            : ''}
        </div>
      ) : null}
      {sharedReviewOwnerLabel && props.sharedReviewStatus ? (
        <div
          className={cn(
            workbookAlertClass({
              tone: props.sharedReviewStatus === 'rejected' ? 'danger' : 'warning',
            }),
            agentPanelMetaTextClass(),
            'mt-2 border-[var(--wb-border-strong)]',
          )}
        >
          {props.sharedReviewStatus === 'pending'
            ? `Owner review is in progress with ${sharedReviewOwnerLabel}.`
            : props.sharedReviewStatus === 'approved'
              ? `Approved by ${sharedReviewDecisionLabel ?? sharedReviewOwnerLabel}.`
              : `Returned by ${sharedReviewDecisionLabel ?? sharedReviewOwnerLabel}.`}
        </div>
      ) : null}
      {props.canFinalizeSharedBundle && props.sharedReviewStatus !== null ? (
        <div className="mt-2 flex items-center justify-end gap-2">
          <Button
            className={workbookButtonClass({ tone: 'neutral' })}
            data-testid="workbook-agent-review-item-reject"
            type="button"
            onClick={() => {
              props.onReview('rejected')
            }}
          >
            Reject
          </Button>
          <Button
            className={workbookButtonClass({ tone: 'accent' })}
            data-testid="workbook-agent-review-item-approve"
            type="button"
            onClick={() => {
              props.onReview('approved')
            }}
          >
            Approve
          </Button>
        </div>
      ) : null}
      {props.canRecommendSharedBundle && props.sharedReviewStatus === 'pending' ? (
        <div className="mt-2 flex items-center justify-end gap-2">
          <Button
            className={workbookButtonClass({ tone: 'neutral' })}
            data-testid="workbook-agent-review-item-reject"
            type="button"
            onClick={() => {
              props.onReview('rejected')
            }}
          >
            Recommend reject
          </Button>
          <Button
            className={workbookButtonClass({ tone: 'accent' })}
            data-testid="workbook-agent-review-item-approve"
            type="button"
            onClick={() => {
              props.onReview('approved')
            }}
          >
            Recommend approve
          </Button>
        </div>
      ) : null}
      {props.preview?.cellDiffs?.length ? (
        <div className="mt-2 overflow-hidden rounded-[var(--wb-radius-control)] border border-[var(--wb-border)]">
          <div className="max-h-44 overflow-y-auto">
            {props.preview.cellDiffs.map((diff) => (
              <div
                key={`${diff.sheetName}:${diff.address}`}
                className="grid grid-cols-[minmax(0,1fr)_minmax(0,1fr)] gap-x-2 border-t border-[var(--wb-border)] px-2 py-2 first:border-t-0"
              >
                <div className={cn(agentPanelLabelTextClass(), 'col-span-2')}>
                  {diff.sheetName}!{diff.address}
                </div>
                <div className="col-span-2 mt-1 flex flex-wrap gap-1">
                  {diff.changeKinds.map((kind) => (
                    <span key={kind} className={workbookPillClass({ tone: 'neutral' })}>
                      {renderPreviewChangeKind(kind)}
                    </span>
                  ))}
                </div>
                <div className={agentPanelMetaTextClass()}>{(diff.beforeFormula ?? String(diff.beforeInput ?? '')) || '(empty)'}</div>
                <div className={agentPanelLabelTextClass()}>{(diff.afterFormula ?? String(diff.afterInput ?? '')) || '(empty)'}</div>
              </div>
            ))}
          </div>
        </div>
      ) : null}
      <div className="mt-3 flex items-center justify-end gap-2">
        <Button className={workbookButtonClass({ tone: 'neutral' })} type="button" onClick={props.onDismiss}>
          Clear
        </Button>
        <Button
          className={workbookButtonClass({ tone: 'accent', weight: 'strong' })}
          data-testid="workbook-agent-apply-review-item"
          disabled={!canApply}
          type="button"
          onClick={props.onApply}
        >
          {applyLabel}
        </Button>
      </div>
    </div>
  )
}

export function WorkbookAgentPanel(props: {
  readonly activeThreadId: string | null
  readonly optimisticEntries?: readonly WorkbookAgentTimelineEntry[]
  readonly snapshot: WorkbookAgentThreadSnapshot | null
  readonly activeResponseTurnId: string | null
  readonly showAssistantProgress: boolean
  readonly activeReviewBundle: WorkbookAgentCommandBundle | null
  readonly preview: WorkbookAgentPreviewSummary | null
  readonly sharedApprovalOwnerUserId: string | null
  readonly sharedReviewOwnerUserId: string | null
  readonly sharedReviewStatus: 'pending' | 'approved' | 'rejected' | null
  readonly sharedReviewDecidedByUserId: string | null
  readonly sharedReviewRecommendations: readonly WorkbookAgentSharedReviewRecommendation[]
  readonly currentUserSharedRecommendation: 'approved' | 'rejected' | null
  readonly canFinalizeSharedBundle: boolean
  readonly canRecommendSharedBundle: boolean
  readonly selectedCommandIndexes: readonly number[]
  readonly workflowRuns: readonly WorkbookAgentWorkflowRun[]
  readonly cancellingWorkflowRunId: string | null
  readonly threadSummaries: readonly WorkbookAgentThreadSummary[]
  readonly draft: string
  readonly isLoading: boolean
  readonly isApplyingReviewItem: boolean
  readonly onApplyReviewItem: () => void
  readonly onDraftChange: (value: string) => void
  readonly onDismissReviewItem: () => void
  readonly onReviewReviewItem: (decision: 'approved' | 'rejected') => void
  readonly onInterrupt: () => void
  readonly onSelectAllReviewCommands: () => void
  readonly onSelectThread: (threadId: string) => void
  readonly onToggleReviewCommand: (commandIndex: number) => void
  readonly onCancelWorkflowRun: (runId: string) => void
  readonly onSubmit: () => void
}) {
  const scrollRef = useRef<HTMLDivElement | null>(null)
  const composerViewportRef = useRef<HTMLDivElement | null>(null)
  const composerTextareaRef = useRef<HTMLTextAreaElement | null>(null)

  const optimisticEntries = props.optimisticEntries ?? []

  useEffect(() => {
    const node = scrollRef.current
    if (!node) {
      return
    }
    node.scrollTop = node.scrollHeight
  }, [optimisticEntries.length, props.snapshot?.entries.length, props.snapshot?.status])

  useLayoutEffect(() => {
    const viewport = composerViewportRef.current
    const textarea = composerTextareaRef.current
    if (!viewport || !textarea) {
      return
    }

    textarea.style.height = '0px'
    const measuredHeight = Math.max(textarea.scrollHeight, AGENT_COMPOSER_MIN_HEIGHT)
    textarea.style.height = `${measuredHeight}px`
    viewport.style.height = `${Math.min(measuredHeight, AGENT_COMPOSER_MAX_HEIGHT)}px`
    viewport.scrollTop = viewport.scrollHeight
  }, [props.draft])

  const isRunning = props.snapshot?.status === 'inProgress'
  const visibleEntries = [...optimisticEntries, ...(props.snapshot?.entries ?? [])]
  const progressAnchorIndex =
    props.showAssistantProgress && props.activeResponseTurnId
      ? visibleEntries.findLastIndex((entry) => entry.turnId === props.activeResponseTurnId && !isAppliedExecutionSystemEntry(entry))
      : -1

  return (
    <div
      className="flex h-full min-h-0 w-full flex-col bg-[var(--wb-app-bg)]"
      data-testid="workbook-agent-panel"
      id="workbook-agent-panel"
      style={agentPanelThemeStyle}
    >
      <ScrollArea.Root className={agentPanelScrollAreaRootClass()}>
        <ScrollArea.Viewport
          ref={scrollRef}
          className={agentPanelScrollAreaViewportClass()}
          data-testid="workbook-agent-panel-scroll-viewport"
        >
          <ScrollArea.Content className={agentPanelScrollAreaContentClass()}>
            <div className="bg-[var(--wb-app-bg)] px-2.5 py-2.5">
              <ThreadSummaryStrip
                activeThreadId={props.activeThreadId}
                threadSummaries={props.threadSummaries}
                onSelectThread={props.onSelectThread}
              />
              {props.isLoading ? null : visibleEntries.length > 0 ? (
                <div className="flex flex-col gap-2">
                  <div className={agentPanelTimelineListClass()}>
                    {visibleEntries.map((entry, index) => (
                      <Fragment key={entry.id}>
                        <div className="min-w-0 w-full max-w-full overflow-hidden">
                          <WorkbookAgentEntryRow entry={entry} />
                        </div>
                        {props.showAssistantProgress && progressAnchorIndex === index ? <AssistantProgressRow /> : null}
                      </Fragment>
                    ))}
                    {props.showAssistantProgress && progressAnchorIndex < 0 ? <AssistantProgressRow /> : null}
                  </div>
                  {props.workflowRuns.length > 0 ? (
                    <div className="pt-1">
                      <div className={cn(agentPanelEyebrowTextClass(), 'mb-2')}>Workflows</div>
                      <div className="flex flex-col gap-2">
                        {props.workflowRuns.slice(0, 5).map((run) => (
                          <WorkflowRunRow
                            key={run.runId}
                            isCancelling={props.cancellingWorkflowRunId === run.runId}
                            run={run}
                            onCancel={() => {
                              props.onCancelWorkflowRun(run.runId)
                            }}
                          />
                        ))}
                      </div>
                    </div>
                  ) : null}
                </div>
              ) : null}
            </div>
          </ScrollArea.Content>
        </ScrollArea.Viewport>
        <ScrollArea.Scrollbar className={agentPanelScrollAreaScrollbarClass()} keepMounted orientation="vertical">
          <ScrollArea.Thumb className={agentPanelScrollAreaThumbClass()} />
        </ScrollArea.Scrollbar>
      </ScrollArea.Root>
      <div className={agentPanelFooterClass()}>
        {props.activeReviewBundle ? (
          <div className="mb-3">
            <ReviewItemCard
              activeReviewBundle={props.activeReviewBundle}
              preview={props.preview}
              sharedApprovalOwnerUserId={props.sharedApprovalOwnerUserId}
              sharedReviewOwnerUserId={props.sharedReviewOwnerUserId}
              sharedReviewStatus={props.sharedReviewStatus}
              sharedReviewDecidedByUserId={props.sharedReviewDecidedByUserId}
              sharedReviewRecommendations={props.sharedReviewRecommendations}
              currentUserSharedRecommendation={props.currentUserSharedRecommendation}
              canFinalizeSharedBundle={props.canFinalizeSharedBundle}
              canRecommendSharedBundle={props.canRecommendSharedBundle}
              selectedCommandIndexes={props.selectedCommandIndexes}
              isApplyingReviewItem={props.isApplyingReviewItem}
              onApply={props.onApplyReviewItem}
              onDismiss={props.onDismissReviewItem}
              onReview={props.onReviewReviewItem}
              onSelectAll={props.onSelectAllReviewCommands}
              onToggleCommand={props.onToggleReviewCommand}
            />
          </div>
        ) : null}
        <form
          onSubmit={(event) => {
            event.preventDefault()
            props.onSubmit()
          }}
        >
          <label className="sr-only" htmlFor="workbook-agent-input">
            Ask the workbook assistant
          </label>
          <div className={agentPanelComposerFrameClass()}>
            <ScrollArea.Root className={agentPanelComposerScrollRootClass()}>
              <ScrollArea.Viewport
                ref={composerViewportRef}
                className={agentPanelComposerScrollViewportClass()}
                data-testid="workbook-agent-input-viewport"
              >
                <ScrollArea.Content className={agentPanelComposerScrollContentClass()}>
                  <textarea
                    ref={composerTextareaRef}
                    id="workbook-agent-input"
                    className={agentPanelComposerTextareaClass()}
                    data-testid="workbook-agent-input"
                    placeholder="Ask the workbook assistant"
                    value={props.draft}
                    onChange={(event) => {
                      props.onDraftChange(event.target.value)
                    }}
                    onKeyDown={(event) => {
                      if (isRunning) {
                        return
                      }
                      if (event.key !== 'Enter' || event.shiftKey || event.nativeEvent.isComposing) {
                        return
                      }
                      event.preventDefault()
                      props.onSubmit()
                    }}
                  />
                </ScrollArea.Content>
              </ScrollArea.Viewport>
              <ScrollArea.Scrollbar className={agentPanelScrollAreaScrollbarClass()} orientation="vertical">
                <ScrollArea.Thumb className={agentPanelScrollAreaThumbClass()} />
              </ScrollArea.Scrollbar>
            </ScrollArea.Root>
            <Button
              aria-label={isRunning ? 'Stop' : 'Send message'}
              className={agentPanelComposerSendButtonClass()}
              data-testid="workbook-agent-send"
              disabled={!isRunning && (props.draft.trim().length === 0 || props.isLoading)}
              type="button"
              onClick={() => {
                if (isRunning) {
                  props.onInterrupt()
                  return
                }
                props.onSubmit()
              }}
            >
              {isRunning ? <StopIcon /> : <SendArrowIcon />}
            </Button>
          </div>
        </form>
      </div>
    </div>
  )
}

function SendArrowIcon() {
  return <ArrowUp aria-hidden="true" className="size-5" strokeWidth={1.9} />
}

function StopIcon() {
  return <Square aria-hidden="true" className="size-4 fill-current" strokeWidth={0} />
}
