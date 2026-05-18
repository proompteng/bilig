import { formatAddress, parseCellAddress } from '@bilig/formula'
import type { CellRangeRef } from '@bilig/protocol'
import { WORKBOOK_AGENT_TOOL_NAMES, normalizeWorkbookAgentToolName } from '@bilig/agent-api'
import type {
  CodexDynamicToolCallRequest,
  CodexDynamicToolCallResult,
  WorkbookAgentCommand,
  WorkbookAgentCommandBundle,
} from '@bilig/agent-api'
import type { WorkbookAgentUiContext, WorkbookAgentWorkflowRun } from '@bilig/contracts'
import type { SessionIdentity } from '../http/session.js'
import type { ZeroSyncService } from '../zero/service.js'
import {
  findWorkbookFormulaIssues,
  searchWorkbook,
  summarizeWorkbookStructure,
  traceWorkbookDependencies,
} from './workbook-agent-comprehension.js'
import {
  inspectWorkbookCell,
  inspectWorkbookContext,
  inspectWorkbookRange,
  normalizeWorkbookAgentUiContext,
} from './workbook-agent-inspection.js'
import { handleWorkbookAgentAnnotationToolCall } from './workbook-agent-annotation-tools.js'
import { handleWorkbookAgentAuditToolCall } from './workbook-agent-audit-tools.js'
import { handleWorkbookAgentConditionalFormatToolCall } from './workbook-agent-conditional-format-tools.js'
import { handleWorkbookAgentObjectToolCall } from './workbook-agent-object-tools.js'
import { handleWorkbookAgentMediaToolCall } from './workbook-agent-media-tools.js'
import { handleWorkbookAgentProtectionToolCall } from './workbook-agent-protection-tools.js'
import { handleWorkbookAgentSheetReadToolCall } from './workbook-agent-sheet-read-tools.js'
import { handleWorkbookAgentValidationToolCall } from './workbook-agent-validation-tools.js'
import {
  rangeOrSelectorSchema,
  readRangeToolArgsSchema,
  resolveFormulaRangeRequest,
  resolveRangeOrSelectorRequest,
  resolveReadRangeRequest,
  resolveTransferRangeRequest,
  resolveWriteRangeRequest,
  setFormulaToolArgsSchema,
  transferRangeToolArgsSchema,
  writeRangeToolArgsSchema,
} from './workbook-agent-selector-tooling.js'
import { normalizeWorkbookAgentStylePatch, workbookAgentStylePatchHasChanges } from './workbook-agent-style-patches.js'
import { stringifyJson, textToolResult, type WorkbookAgentStageCommandResult } from './workbook-agent-tool-shared.js'
import {
  normalizeWorkbookAgentToolNumberFormatInput,
  normalizeWorkbookAgentWriteCellInput,
} from './workbook-agent-tool-input-normalization.js'
import {
  resolveWorkbookAgentInspectionTarget,
  resolveWorkbookAgentSelectionRange,
  resolveWorkbookAgentVisibleRange,
  workbookAgentViewportAroundAddress,
} from './workbook-agent-context-geometry.js'
import { listWorkbookNamedRanges, listWorkbookTables, type ResolvedWorkbookSelector } from './workbook-selector-resolver.js'
import { parseWorkbookAgentStructuralToolCommand, sortToolArgsSchema } from './workbook-agent-structural-tools.js'
import { buildWorkbookAgentVerificationReport, stageWorkbookAgentCommandResult } from './workbook-agent-mutation-receipt.js'
import { selectWorkbookRenderedReadback } from './workbook-agent-rendered-readback.js'
import {
  countWorkbookAgentRangesCells,
  createWorkbookAgentRangeChunkPlan,
  ensureWorkbookAgentRangeCellLimit,
  normalizeWorkbookAgentRange,
  toWorkbookAgentRangeRef,
} from './workbook-agent-range-chunks.js'
import { summarizeWorkbookAgentVerificationStatus } from './workbook-agent-verification-status.js'
import {
  MAX_MUTATION_RANGE_CELLS,
  MAX_READ_RANGE_CELLS,
  applyAndVerifyToolArgsSchema,
  clearRangeToolArgsSchema,
  formatRangeToolArgsSchema,
  formulaIssueToolArgsSchema,
  inspectCellToolArgsSchema,
  readRecentChangesToolArgsSchema,
  readRenderedRangeToolArgsSchema,
  searchWorkbookToolArgsSchema,
  setActiveSheetToolArgsSchema,
  setSelectionToolArgsSchema,
  startWorkflowToolArgsSchema,
  traceDependenciesToolArgsSchema,
  undoWorkbookMutationToolArgsSchema,
  type WorkbookAgentStartWorkflowRequest,
} from './workbook-agent-tool-schemas.js'
export type { WorkbookAgentStartWorkflowRequest } from './workbook-agent-tool-schemas.js'
export { workbookAgentDynamicToolSpecs } from './workbook-agent-tool-specs.js'

function serializeSelectorResolution(resolution: ResolvedWorkbookSelector | null) {
  if (!resolution) {
    return null
  }
  return {
    objectType: resolution.objectType,
    displayLabel: resolution.displayLabel,
    resolvedRevision: resolution.resolvedRevision,
    derivedA1Ranges: resolution.derivedA1Ranges,
    table: resolution.table,
    namedRange: resolution.namedRange,
  }
}

function summarizeWorkbookChangeRecord(record: Awaited<ReturnType<ZeroSyncService['listWorkbookChanges']>>[number]) {
  return {
    revision: record.revision,
    actorUserId: record.actorUserId,
    eventKind: record.eventKind,
    summary: record.summary,
    sheetName: record.sheetName,
    anchorAddress: record.anchorAddress,
    range: record.range,
    createdAtUnixMs: record.createdAtUnixMs,
    revertedByRevision: record.revertedByRevision,
    revertsRevision: record.revertsRevision,
  }
}

async function buildVerificationReport(input: {
  readonly context: WorkbookAgentToolContext
  readonly revision: number | null
  readonly ranges: readonly CellRangeRef[]
  readonly includeFormulaIssues?: boolean
  readonly includeInvariants?: boolean
}) {
  return await buildWorkbookAgentVerificationReport({
    context: input.context,
    revision: input.revision,
    ranges: input.ranges,
    ...(input.includeFormulaIssues !== undefined ? { includeFormulaIssues: input.includeFormulaIssues } : {}),
    ...(input.includeInvariants !== undefined ? { includeInvariants: input.includeInvariants } : {}),
  })
}

export interface WorkbookAgentToolContext {
  readonly documentId: string
  readonly session: SessionIdentity
  readonly uiContext: WorkbookAgentUiContext | null
  readonly zeroSyncService: ZeroSyncService
  readonly stageCommand: (command: WorkbookAgentCommand) => Promise<WorkbookAgentCommandBundle | WorkbookAgentStageCommandResult>
  readonly updateUiContext?: (context: WorkbookAgentUiContext | null) => Promise<void>
  readonly awaitRenderedRevision?: (revision: number) => Promise<void>
  readonly startWorkflow?: (input: WorkbookAgentStartWorkflowRequest) => Promise<WorkbookAgentWorkflowRun>
}

async function stageCommandResult(context: WorkbookAgentToolContext, command: WorkbookAgentCommand): Promise<CodexDynamicToolCallResult> {
  return await stageWorkbookAgentCommandResult(context, command, command.kind)
}

function workflowToolResult(run: WorkbookAgentWorkflowRun): CodexDynamicToolCallResult {
  return textToolResult(
    stringifyJson({
      workflowRun: {
        runId: run.runId,
        workflowTemplate: run.workflowTemplate,
        title: run.title,
        summary: run.summary,
        status: run.status,
        completedAtUnixMs: run.completedAtUnixMs,
        errorMessage: run.errorMessage,
      },
      artifact: run.artifact,
    }),
  )
}

export async function handleWorkbookAgentToolCall(
  context: WorkbookAgentToolContext,
  request: CodexDynamicToolCallRequest,
): Promise<CodexDynamicToolCallResult> {
  try {
    const normalizedTool = normalizeWorkbookAgentToolName(request.tool)
    const sheetReadToolResult = await handleWorkbookAgentSheetReadToolCall(context, request)
    if (sheetReadToolResult) {
      return sheetReadToolResult
    }
    const auditToolResult = await handleWorkbookAgentAuditToolCall(context, request)
    if (auditToolResult) {
      return auditToolResult
    }
    const objectToolResult = await handleWorkbookAgentObjectToolCall(context, request)
    if (objectToolResult) {
      return objectToolResult
    }
    const mediaToolResult = await handleWorkbookAgentMediaToolCall(context, request)
    if (mediaToolResult) {
      return mediaToolResult
    }
    const protectionToolResult = await handleWorkbookAgentProtectionToolCall(context, request)
    if (protectionToolResult) {
      return protectionToolResult
    }
    const annotationToolResult = await handleWorkbookAgentAnnotationToolCall(context, request)
    if (annotationToolResult) {
      return annotationToolResult
    }
    const conditionalFormatToolResult = await handleWorkbookAgentConditionalFormatToolCall(context, request)
    if (conditionalFormatToolResult) {
      return conditionalFormatToolResult
    }
    const validationToolResult = await handleWorkbookAgentValidationToolCall(context, request)
    if (validationToolResult) {
      return validationToolResult
    }
    const structuralCommand = parseWorkbookAgentStructuralToolCommand(request)
    if (structuralCommand) {
      return await stageCommandResult(context, structuralCommand)
    }
    switch (normalizedTool) {
      case WORKBOOK_AGENT_TOOL_NAMES.getContext: {
        const text = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          inspectWorkbookContext(runtime, context.uiContext),
        )
        return textToolResult(text)
      }
      case WORKBOOK_AGENT_TOOL_NAMES.readWorkbook: {
        const summary = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => ({
          documentId: context.documentId,
          context: normalizeWorkbookAgentUiContext(runtime, context.uiContext),
          ...summarizeWorkbookStructure(runtime),
        }))
        return textToolResult(stringifyJson(summary))
      }
      case WORKBOOK_AGENT_TOOL_NAMES.setActiveSheet: {
        const args = setActiveSheetToolArgsSchema.parse(request.arguments)
        const nextContext = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
          const sheets = runtime.engine.exportSnapshot().sheets.map((sheet) => sheet.name)
          if (!sheets.includes(args.sheetName)) {
            throw new Error(`Sheet ${args.sheetName} does not exist`)
          }
          const currentContext = normalizeWorkbookAgentUiContext(runtime, context.uiContext)
          const address = args.address ?? currentContext?.selection.address ?? 'A1'
          return normalizeWorkbookAgentUiContext(runtime, {
            selection: {
              sheetName: args.sheetName,
              address,
              range: {
                startAddress: address,
                endAddress: address,
              },
            },
            viewport: workbookAgentViewportAroundAddress(args.sheetName, address, currentContext?.viewport),
          })
        })
        if (!context.updateUiContext) {
          throw new Error('Active sheet control is not available in this workbook assistant session')
        }
        await context.updateUiContext(nextContext)
        return textToolResult(
          stringifyJson({
            updated: true,
            verificationComplete: false,
            modelConfirmation: {
              matched: nextContext?.selection.sheetName === args.sheetName,
              sheetName: nextContext?.selection.sheetName ?? null,
              address: nextContext?.selection.address ?? null,
            },
            selectionConfidence: {
              level: 'model_only',
              browserConfirmed: false,
              reason:
                'The active sheet was updated in the workbook assistant model, but browser-rendered confirmation is not available synchronously from this tool call.',
            },
            browserConfirmation: {
              status: 'not_proven',
              reason:
                'The server emitted the new workbook context, but this tool call has no synchronous browser acknowledgement channel. Use read_rendered_selection or read_rendered_range after the browser refreshes to prove visible state.',
            },
            context: nextContext,
          }),
        )
      }
      case WORKBOOK_AGENT_TOOL_NAMES.setSelection: {
        const args = setSelectionToolArgsSchema.parse(request.arguments)
        const nextContext = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
          const currentContext = normalizeWorkbookAgentUiContext(runtime, context.uiContext)
          const sheetName = args.sheetName ?? currentContext?.selection.sheetName
          if (!sheetName) {
            throw new Error('sheetName is required when no browser workbook context exists')
          }
          const sheets = runtime.engine.exportSnapshot().sheets.map((sheet) => sheet.name)
          if (!sheets.includes(sheetName)) {
            throw new Error(`Sheet ${sheetName} does not exist`)
          }
          const start = parseCellAddress(args.address, sheetName)
          const end = parseCellAddress(args.endAddress ?? args.address, sheetName)
          const startAddress = formatAddress(Math.min(start.row, end.row), Math.min(start.col, end.col))
          const endAddress = formatAddress(Math.max(start.row, end.row), Math.max(start.col, end.col))
          return normalizeWorkbookAgentUiContext(runtime, {
            selection: {
              sheetName,
              address: args.address,
              range: {
                startAddress,
                endAddress,
              },
            },
            viewport: workbookAgentViewportAroundAddress(sheetName, args.address, currentContext?.viewport),
          })
        })
        if (!context.updateUiContext) {
          throw new Error('Selection control is not available in this workbook assistant session')
        }
        await context.updateUiContext(nextContext)
        return textToolResult(
          stringifyJson({
            updated: true,
            verificationComplete: false,
            modelConfirmation: {
              matched:
                nextContext?.selection.address === args.address &&
                (args.sheetName === undefined || nextContext?.selection.sheetName === args.sheetName),
              sheetName: nextContext?.selection.sheetName ?? null,
              address: nextContext?.selection.address ?? null,
              range: nextContext?.selection.range ?? null,
            },
            selectionConfidence: {
              level: 'model_only',
              browserConfirmed: false,
              reason:
                'The selection was updated in the workbook assistant model, but browser-rendered confirmation is not available synchronously from this tool call.',
            },
            browserConfirmation: {
              status: 'not_proven',
              reason:
                'The server emitted the new selection context, but this tool call has no synchronous browser acknowledgement channel. Use read_rendered_selection after the browser refreshes to prove visible selection state.',
            },
            context: nextContext,
          }),
        )
      }
      case WORKBOOK_AGENT_TOOL_NAMES.readRenderedSelection: {
        const result = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
          const uiContext = normalizeWorkbookAgentUiContext(runtime, context.uiContext)
          const range = resolveWorkbookAgentSelectionRange(uiContext)
          ensureWorkbookAgentRangeCellLimit(range, MAX_READ_RANGE_CELLS)
          const authoritativeReadback = inspectWorkbookRange(runtime, range)
          const authoritativeRows = authoritativeReadback.rows.filter(Array.isArray) as readonly (readonly unknown[])[]
          return {
            authoritativeReadback,
            renderedReadback: selectWorkbookRenderedReadback({
              renderedContext: uiContext?.rendered,
              requestedRange: range,
              authoritativeRows,
              minRevision: runtime.headRevision,
            }),
          }
        })
        return textToolResult(stringifyJson(result))
      }
      case WORKBOOK_AGENT_TOOL_NAMES.readRenderedRange: {
        const args = readRenderedRangeToolArgsSchema.parse(request.arguments)
        const range = normalizeWorkbookAgentRange({
          sheetName: args.sheetName,
          startAddress: args.startAddress,
          endAddress: args.endAddress,
        })
        ensureWorkbookAgentRangeCellLimit(range, MAX_READ_RANGE_CELLS)
        const result = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
          const normalizedContext = normalizeWorkbookAgentUiContext(runtime, context.uiContext)
          const authoritativeReadback = inspectWorkbookRange(runtime, range)
          const authoritativeRows = authoritativeReadback.rows.filter(Array.isArray) as readonly (readonly unknown[])[]
          return {
            authoritativeReadback,
            renderedReadback: selectWorkbookRenderedReadback({
              renderedContext: normalizedContext?.rendered,
              requestedRange: range,
              authoritativeRows,
              minRevision: runtime.headRevision,
            }),
          }
        })
        return textToolResult(stringifyJson(result))
      }
      case WORKBOOK_AGENT_TOOL_NAMES.applyAndVerify: {
        const args = applyAndVerifyToolArgsSchema.parse(request.arguments)
        const revision = await context.zeroSyncService.getWorkbookHeadRevision(context.documentId)
        await context.awaitRenderedRevision?.(revision)
        const ranges = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
          if (args.range) {
            return [toWorkbookAgentRangeRef(args.range)]
          }
          const uiContext = normalizeWorkbookAgentUiContext(runtime, context.uiContext)
          return uiContext ? [resolveWorkbookAgentSelectionRange(uiContext)] : []
        })
        const report = await buildVerificationReport({
          context,
          revision,
          ranges,
          ...(args.includeFormulaIssues !== undefined ? { includeFormulaIssues: args.includeFormulaIssues } : {}),
          ...(args.includeInvariants !== undefined ? { includeInvariants: args.includeInvariants } : {}),
        })
        const { verificationComplete } = summarizeWorkbookAgentVerificationStatus({
          renderedReadback: report.renderedReadback,
          formulaIssues: report.formulaIssues,
          invariants: report.invariants,
          requireTargetRange: true,
          targetRangeCount: ranges.length,
        })
        return textToolResult(
          stringifyJson({
            status: verificationComplete ? 'verified' : 'verification_incomplete',
            verificationComplete,
            ...report,
          }),
        )
      }
      case WORKBOOK_AGENT_TOOL_NAMES.undoWorkbookMutation: {
        const args = undoWorkbookMutationToolArgsSchema.parse(request.arguments)
        const beforeRevision = await context.zeroSyncService.getWorkbookHeadRevision(context.documentId)
        const recentChanges = await context.zeroSyncService.listWorkbookChanges(context.documentId, 25)
        const targetChange =
          args.revision !== undefined
            ? (recentChanges.find((change) => change.revision === args.revision) ?? null)
            : (recentChanges.find(
                (change) =>
                  change.undoBundle !== null &&
                  change.revertedByRevision === null &&
                  change.eventKind !== 'revertChange' &&
                  change.revertsRevision === null,
              ) ?? null)
        if (args.revision !== undefined) {
          await context.zeroSyncService.applyServerMutator(
            'workbook.revertChange',
            {
              documentId: context.documentId,
              revision: args.revision,
            },
            context.session,
          )
        } else {
          await context.zeroSyncService.applyServerMutator(
            'workbook.undoLatestChange',
            {
              documentId: context.documentId,
            },
            context.session,
          )
        }
        const afterRevision = await context.zeroSyncService.getWorkbookHeadRevision(context.documentId)
        await context.awaitRenderedRevision?.(afterRevision)
        const verificationRange = targetChange?.range
          ? [
              {
                sheetName: targetChange.range.sheetName,
                startAddress: targetChange.range.startAddress,
                endAddress: targetChange.range.endAddress,
              },
            ]
          : []
        const verification = await buildVerificationReport({
          context,
          revision: afterRevision,
          ranges: verificationRange,
        })
        const { verificationComplete } = summarizeWorkbookAgentVerificationStatus({
          renderedReadback: verification.renderedReadback,
          formulaIssues: verification.formulaIssues,
          invariants: verification.invariants,
          requireTargetRange: true,
          targetRangeCount: verificationRange.length,
        })
        return textToolResult(
          stringifyJson({
            undone: true,
            applied: true,
            staged: false,
            queuedForTurnApply: false,
            status: verificationComplete ? 'applied' : 'verification_incomplete',
            verificationComplete,
            revision: {
              before: beforeRevision,
              after: afterRevision,
              reverted: args.revision ?? targetChange?.revision ?? null,
            },
            targetChange: targetChange
              ? {
                  revision: targetChange.revision,
                  summary: targetChange.summary,
                  sheetName: targetChange.sheetName,
                  anchorAddress: targetChange.anchorAddress,
                  range: targetChange.range,
                }
              : null,
            verification,
          }),
        )
      }
      case WORKBOOK_AGENT_TOOL_NAMES.listNamedRanges: {
        const namedRanges = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => ({
          documentId: context.documentId,
          namedRangeCount: runtime.engine.getDefinedNames().length,
          namedRanges: listWorkbookNamedRanges(runtime),
        }))
        return textToolResult(stringifyJson(namedRanges))
      }
      case WORKBOOK_AGENT_TOOL_NAMES.listTables: {
        const tables = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => ({
          documentId: context.documentId,
          tableCount: runtime.engine.getTables().length,
          tables: listWorkbookTables(runtime),
        }))
        return textToolResult(stringifyJson(tables))
      }
      case WORKBOOK_AGENT_TOOL_NAMES.readRange: {
        const args = readRangeToolArgsSchema.parse(request.arguments)
        const result = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
          const uiContext = normalizeWorkbookAgentUiContext(runtime, context.uiContext)
          const resolved = resolveReadRangeRequest({
            runtime,
            args,
            uiContext,
          })
          const totalCells = countWorkbookAgentRangesCells(resolved.ranges)
          if (totalCells > MAX_READ_RANGE_CELLS) {
            const firstRange = resolved.ranges[0]
            if (!firstRange) {
              throw new Error('Resolved selector did not produce a readable range')
            }
            const firstPlan = createWorkbookAgentRangeChunkPlan(firstRange, MAX_READ_RANGE_CELLS)
            const firstChunk = firstPlan.chunks[0]
            if (!firstChunk) {
              throw new Error('Resolved selector did not produce a readable chunk')
            }
            const currentRange = {
              sheetName: firstChunk.sheetName,
              startAddress: firstChunk.startAddress,
              endAddress: firstChunk.endAddress,
            }
            return {
              resolvedSelector: serializeSelectorResolution(resolved.resolution),
              chunked: true,
              truncated: true,
              totalCells,
              cellLimit: MAX_READ_RANGE_CELLS,
              rangeCount: resolved.ranges.length,
              currentChunk: firstChunk,
              nextChunk: firstPlan.chunks[1] ?? null,
              chunkPlan:
                resolved.ranges.length === 1
                  ? firstPlan
                  : {
                      rangeCount: resolved.ranges.length,
                      plans: resolved.ranges.map((range) => createWorkbookAgentRangeChunkPlan(range, MAX_READ_RANGE_CELLS)),
                    },
              readback: inspectWorkbookRange(runtime, currentRange),
            }
          }
          const inspectedRanges = resolved.ranges.map((range) => inspectWorkbookRange(runtime, range))
          if (inspectedRanges.length === 1) {
            return {
              resolvedSelector: serializeSelectorResolution(resolved.resolution),
              ...inspectedRanges[0],
            }
          }
          return {
            resolvedSelector: serializeSelectorResolution(resolved.resolution),
            rangeCount: inspectedRanges.length,
            ranges: inspectedRanges,
          }
        })
        return textToolResult(stringifyJson(result))
      }
      case WORKBOOK_AGENT_TOOL_NAMES.readSelection: {
        const result = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
          const range = resolveWorkbookAgentSelectionRange(normalizeWorkbookAgentUiContext(runtime, context.uiContext))
          ensureWorkbookAgentRangeCellLimit(range, MAX_READ_RANGE_CELLS)
          return inspectWorkbookRange(runtime, range)
        })
        return textToolResult(stringifyJson(result))
      }
      case WORKBOOK_AGENT_TOOL_NAMES.readVisibleRange: {
        const result = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
          const range = resolveWorkbookAgentVisibleRange(normalizeWorkbookAgentUiContext(runtime, context.uiContext))
          ensureWorkbookAgentRangeCellLimit(range, MAX_READ_RANGE_CELLS)
          return inspectWorkbookRange(runtime, range)
        })
        return textToolResult(stringifyJson(result))
      }
      case WORKBOOK_AGENT_TOOL_NAMES.readRecentChanges: {
        const args = readRecentChangesToolArgsSchema.parse(request.arguments)
        const changes = await context.zeroSyncService.listWorkbookChanges(context.documentId, args.limit)
        return textToolResult(
          stringifyJson({
            documentId: context.documentId,
            changeCount: changes.length,
            changes: changes.map((record) => summarizeWorkbookChangeRecord(record)),
          }),
        )
      }
      case WORKBOOK_AGENT_TOOL_NAMES.startWorkflow: {
        const args = startWorkflowToolArgsSchema.parse(request.arguments)
        if (!context.startWorkflow) {
          throw new Error('Built-in workflow execution is not available in this session')
        }
        return workflowToolResult(await context.startWorkflow(args))
      }
      case WORKBOOK_AGENT_TOOL_NAMES.inspectCell: {
        const args = inspectCellToolArgsSchema.parse(request.arguments)
        const result = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          inspectWorkbookCell(
            runtime,
            resolveWorkbookAgentInspectionTarget(normalizeWorkbookAgentUiContext(runtime, context.uiContext), args),
          ),
        )
        return textToolResult(stringifyJson(result))
      }
      case WORKBOOK_AGENT_TOOL_NAMES.findFormulaIssues: {
        const args = formulaIssueToolArgsSchema.parse(request.arguments)
        const report = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          findWorkbookFormulaIssues(runtime, {
            ...(args.sheetName ? { sheetName: args.sheetName } : {}),
            ...(args.limit !== undefined ? { limit: args.limit } : {}),
          }),
        )
        return textToolResult(stringifyJson(report))
      }
      case WORKBOOK_AGENT_TOOL_NAMES.searchWorkbook: {
        const args = searchWorkbookToolArgsSchema.parse(request.arguments)
        const report = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          searchWorkbook(runtime, {
            query: args.query,
            ...(args.sheetName ? { sheetName: args.sheetName } : {}),
            ...(args.limit !== undefined ? { limit: args.limit } : {}),
          }),
        )
        return textToolResult(stringifyJson(report))
      }
      case WORKBOOK_AGENT_TOOL_NAMES.traceDependencies: {
        const args = traceDependenciesToolArgsSchema.parse(request.arguments)
        const report = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) => {
          const target = resolveWorkbookAgentInspectionTarget(normalizeWorkbookAgentUiContext(runtime, context.uiContext), args)
          return traceWorkbookDependencies(runtime, {
            sheetName: target.sheetName,
            address: target.address,
            ...(args.direction ? { direction: args.direction } : {}),
            ...(args.depth !== undefined ? { depth: args.depth } : {}),
          })
        })
        return textToolResult(stringifyJson(report))
      }
      case WORKBOOK_AGENT_TOOL_NAMES.writeRange: {
        const args = writeRangeToolArgsSchema.parse(request.arguments)
        const resolved = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          resolveWriteRangeRequest({
            runtime,
            args,
            uiContext: normalizeWorkbookAgentUiContext(runtime, context.uiContext),
          }),
        )
        const values = args.values
        const start = parseCellAddress(resolved.startAddress, resolved.sheetName)
        const maxWidth = values.reduce((width, rowValues) => Math.max(width, rowValues.length), 0)
        const endAddress = formatAddress(start.row + values.length - 1, start.col + maxWidth - 1)
        ensureWorkbookAgentRangeCellLimit(
          {
            sheetName: resolved.sheetName,
            startAddress: resolved.startAddress,
            endAddress,
          },
          MAX_MUTATION_RANGE_CELLS,
        )
        return await stageCommandResult(context, {
          kind: 'writeRange',
          sheetName: resolved.sheetName,
          startAddress: resolved.startAddress,
          values: values.map((rowValues) => rowValues.map((cellInput) => normalizeWorkbookAgentWriteCellInput(cellInput))),
        })
      }
      case WORKBOOK_AGENT_TOOL_NAMES.setFormula: {
        const args = setFormulaToolArgsSchema.parse(request.arguments)
        const resolved = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          resolveFormulaRangeRequest({
            runtime,
            args,
            uiContext: normalizeWorkbookAgentUiContext(runtime, context.uiContext),
          }),
        )
        ensureWorkbookAgentRangeCellLimit(resolved.range, MAX_MUTATION_RANGE_CELLS)
        return await stageCommandResult(context, {
          kind: 'setRangeFormulas',
          range: resolved.range,
          formulas: args.formulas,
        })
      }
      case WORKBOOK_AGENT_TOOL_NAMES.clearRange: {
        const args = clearRangeToolArgsSchema.parse(request.arguments)
        const resolved = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          resolveRangeOrSelectorRequest({
            runtime,
            args,
            uiContext: normalizeWorkbookAgentUiContext(runtime, context.uiContext),
          }),
        )
        ensureWorkbookAgentRangeCellLimit(resolved.range, MAX_MUTATION_RANGE_CELLS)
        return await stageCommandResult(context, {
          kind: 'clearRange',
          range: resolved.range,
        })
      }
      case WORKBOOK_AGENT_TOOL_NAMES.formatRange: {
        const args = formatRangeToolArgsSchema.parse(request.arguments)
        const resolved = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          resolveRangeOrSelectorRequest({
            runtime,
            args: {
              ...(args.range ? { range: args.range } : {}),
              ...(args.selector ? { selector: args.selector } : {}),
            },
            uiContext: normalizeWorkbookAgentUiContext(runtime, context.uiContext),
          }),
        )
        ensureWorkbookAgentRangeCellLimit(resolved.range, MAX_MUTATION_RANGE_CELLS)
        const formatCommand: Extract<WorkbookAgentCommand, { kind: 'formatRange' }> = {
          kind: 'formatRange',
          range: resolved.range,
        }
        if (args.patch !== undefined) {
          const normalizedPatch = normalizeWorkbookAgentStylePatch(args.patch)
          if (!workbookAgentStylePatchHasChanges(normalizedPatch)) {
            throw new Error('format_range patch did not include any supported style fields')
          }
          formatCommand.patch = normalizedPatch
        }
        if (args.numberFormat !== undefined) {
          formatCommand.numberFormat = normalizeWorkbookAgentToolNumberFormatInput(args.numberFormat)
        }
        return await stageCommandResult(context, formatCommand)
      }
      case WORKBOOK_AGENT_TOOL_NAMES.fillRange: {
        const args = transferRangeToolArgsSchema.parse(request.arguments)
        const resolved = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          resolveTransferRangeRequest({
            runtime,
            args,
            uiContext: normalizeWorkbookAgentUiContext(runtime, context.uiContext),
          }),
        )
        ensureWorkbookAgentRangeCellLimit(resolved.source, MAX_MUTATION_RANGE_CELLS)
        ensureWorkbookAgentRangeCellLimit(resolved.target, MAX_MUTATION_RANGE_CELLS)
        return await stageCommandResult(context, {
          kind: 'fillRange',
          source: resolved.source,
          target: resolved.target,
        })
      }
      case WORKBOOK_AGENT_TOOL_NAMES.copyRange: {
        const args = transferRangeToolArgsSchema.parse(request.arguments)
        const resolved = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          resolveTransferRangeRequest({
            runtime,
            args,
            uiContext: normalizeWorkbookAgentUiContext(runtime, context.uiContext),
          }),
        )
        ensureWorkbookAgentRangeCellLimit(resolved.source, MAX_MUTATION_RANGE_CELLS)
        ensureWorkbookAgentRangeCellLimit(resolved.target, MAX_MUTATION_RANGE_CELLS)
        return await stageCommandResult(context, {
          kind: 'copyRange',
          source: resolved.source,
          target: resolved.target,
        })
      }
      case WORKBOOK_AGENT_TOOL_NAMES.moveRange: {
        const args = transferRangeToolArgsSchema.parse(request.arguments)
        const resolved = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          resolveTransferRangeRequest({
            runtime,
            args,
            uiContext: normalizeWorkbookAgentUiContext(runtime, context.uiContext),
          }),
        )
        ensureWorkbookAgentRangeCellLimit(resolved.source, MAX_MUTATION_RANGE_CELLS)
        ensureWorkbookAgentRangeCellLimit(resolved.target, MAX_MUTATION_RANGE_CELLS)
        return await stageCommandResult(context, {
          kind: 'moveRange',
          source: resolved.source,
          target: resolved.target,
        })
      }
      case WORKBOOK_AGENT_TOOL_NAMES.setFilter: {
        const args = rangeOrSelectorSchema.parse(request.arguments)
        const resolved = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          resolveRangeOrSelectorRequest({
            runtime,
            args,
            uiContext: normalizeWorkbookAgentUiContext(runtime, context.uiContext),
          }),
        )
        ensureWorkbookAgentRangeCellLimit(resolved.range, MAX_MUTATION_RANGE_CELLS)
        return await stageCommandResult(context, {
          kind: 'setFilter',
          range: resolved.range,
        })
      }
      case WORKBOOK_AGENT_TOOL_NAMES.clearFilter: {
        const args = rangeOrSelectorSchema.parse(request.arguments)
        const resolved = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          resolveRangeOrSelectorRequest({
            runtime,
            args,
            uiContext: normalizeWorkbookAgentUiContext(runtime, context.uiContext),
          }),
        )
        ensureWorkbookAgentRangeCellLimit(resolved.range, MAX_MUTATION_RANGE_CELLS)
        return await stageCommandResult(context, {
          kind: 'clearFilter',
          range: resolved.range,
        })
      }
      case WORKBOOK_AGENT_TOOL_NAMES.setSort: {
        const args = sortToolArgsSchema.parse(request.arguments)
        const resolved = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          resolveRangeOrSelectorRequest({
            runtime,
            args: {
              ...(args.range ? { range: args.range } : {}),
              ...(args.selector ? { selector: args.selector } : {}),
            },
            uiContext: normalizeWorkbookAgentUiContext(runtime, context.uiContext),
          }),
        )
        ensureWorkbookAgentRangeCellLimit(resolved.range, MAX_MUTATION_RANGE_CELLS)
        return await stageCommandResult(context, {
          kind: 'setSort',
          range: resolved.range,
          keys: args.keys.map((key) => ({
            keyAddress: key.keyAddress,
            direction: key.direction,
          })),
        })
      }
      case WORKBOOK_AGENT_TOOL_NAMES.clearSort: {
        const args = rangeOrSelectorSchema.parse(request.arguments)
        const resolved = await context.zeroSyncService.inspectWorkbook(context.documentId, (runtime) =>
          resolveRangeOrSelectorRequest({
            runtime,
            args,
            uiContext: normalizeWorkbookAgentUiContext(runtime, context.uiContext),
          }),
        )
        ensureWorkbookAgentRangeCellLimit(resolved.range, MAX_MUTATION_RANGE_CELLS)
        return await stageCommandResult(context, {
          kind: 'clearSort',
          range: resolved.range,
        })
      }
      default:
        return textToolResult(`Unknown bilig tool: ${request.tool}`, false)
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error)
    return textToolResult(`Tool ${request.tool} failed: ${message}`, false)
  }
}
