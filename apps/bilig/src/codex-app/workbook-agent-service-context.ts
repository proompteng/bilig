import type { WorkbookAgentCommand } from '@bilig/agent-api'
import type { WorkbookAgentUiContext } from '@bilig/contracts'
import { cloneUiContext, type WorkbookAgentThreadState } from './workbook-agent-service-shared.js'

export function stripRenderedWorkbookAgentContext(context: WorkbookAgentUiContext): WorkbookAgentUiContext {
  return {
    selection: {
      sheetName: context.selection.sheetName,
      address: context.selection.address,
      ...(context.selection.range
        ? {
            range: {
              startAddress: context.selection.range.startAddress,
              endAddress: context.selection.range.endAddress,
            },
          }
        : {}),
    },
    viewport: { ...context.viewport },
  }
}

export function applyWorkbookAgentStructuralContextHints(
  context: WorkbookAgentUiContext | null,
  commands: readonly WorkbookAgentCommand[],
): WorkbookAgentUiContext | null {
  let nextContext = cloneUiContext(context)
  let structuralContextChanged = false
  for (const command of commands) {
    if (command.kind === 'createSheet') {
      nextContext = {
        selection: {
          sheetName: command.name,
          address: 'A1',
          range: {
            startAddress: 'A1',
            endAddress: 'A1',
          },
        },
        viewport: {
          rowStart: 0,
          rowEnd: 20,
          colStart: 0,
          colEnd: 10,
        },
      }
      structuralContextChanged = true
      continue
    }
    if (!nextContext) {
      continue
    }
    if (command.kind === 'renameSheet' && nextContext.selection.sheetName === command.currentName) {
      nextContext = {
        ...stripRenderedWorkbookAgentContext(nextContext),
        selection: {
          ...nextContext.selection,
          sheetName: command.nextName,
        },
      }
      structuralContextChanged = true
      continue
    }
    if (command.kind === 'deleteSheet' && nextContext.selection.sheetName === command.name) {
      nextContext = stripRenderedWorkbookAgentContext(nextContext)
      structuralContextChanged = true
    }
  }
  return structuralContextChanged && nextContext ? stripRenderedWorkbookAgentContext(nextContext) : nextContext
}

export function updateWorkbookAgentDurableUiContextFromUser(input: {
  readonly sessionState: WorkbookAgentThreadState
  readonly context: WorkbookAgentUiContext
  readonly userId: string
}): void {
  input.sessionState.durable.context = cloneUiContext(input.context)
  const activeTurnId = input.sessionState.live.activeTurnId
  if (!activeTurnId) {
    return
  }
  const activeTurnActorUserId = input.sessionState.live.turnActorUserIdByTurn.get(activeTurnId)
  if (activeTurnActorUserId === undefined || activeTurnActorUserId === input.userId) {
    input.sessionState.live.turnContextByTurn.set(activeTurnId, cloneUiContext(input.context))
  }
}
