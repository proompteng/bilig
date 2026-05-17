import type { TransformQueryFunction } from '@rocicorp/zero/server'
import type { z } from 'zod'
import {
  queries,
  workbookCellArgsSchema,
  workbookColumnTileArgsSchema,
  workbookQueryArgsSchema,
  workbookRowTileArgsSchema,
  workbookThreadArgsSchema,
  workbookTileArgsSchema,
} from './queries.js'

type QueryTransformHandler = {
  readonly execute: (args: Parameters<TransformQueryFunction>[1], userID: string) => ReturnType<TransformQueryFunction>
}

function createQueryTransformHandler<TArgs>(
  argsSchema: z.ZodType<TArgs>,
  execute: (options: { readonly args: TArgs; readonly ctx: { readonly userID: string } }) => ReturnType<TransformQueryFunction>,
): QueryTransformHandler {
  return {
    execute: (args, userID) => execute({ args: argsSchema.parse(args), ctx: { userID } }),
  }
}

const zeroQueryTransforms: Record<string, QueryTransformHandler> = {
  'workbook.get': createQueryTransformHandler(workbookQueryArgsSchema, queries.workbook.get.fn),
  'workbooks.get': createQueryTransformHandler(workbookQueryArgsSchema, queries.workbooks.get.fn),
  'sheet.byWorkbook': createQueryTransformHandler(workbookQueryArgsSchema, queries.sheet.byWorkbook.fn),
  'sheets.byWorkbook': createQueryTransformHandler(workbookQueryArgsSchema, queries.sheets.byWorkbook.fn),
  'cellInput.one': createQueryTransformHandler(workbookCellArgsSchema, queries.cellInput.one.fn),
  'cellInput.tile': createQueryTransformHandler(workbookTileArgsSchema, queries.cellInput.tile.fn),
  'cells.one': createQueryTransformHandler(workbookCellArgsSchema, queries.cells.one.fn),
  'cells.tile': createQueryTransformHandler(workbookTileArgsSchema, queries.cells.tile.fn),
  'cellEval.one': createQueryTransformHandler(workbookCellArgsSchema, queries.cellEval.one.fn),
  'cellEval.tile': createQueryTransformHandler(workbookTileArgsSchema, queries.cellEval.tile.fn),
  'cellRender.one': createQueryTransformHandler(workbookCellArgsSchema, queries.cellRender.one.fn),
  'cellRender.tile': createQueryTransformHandler(workbookTileArgsSchema, queries.cellRender.tile.fn),
  'sheetRow.tile': createQueryTransformHandler(workbookRowTileArgsSchema, queries.sheetRow.tile.fn),
  'rowMetadata.tile': createQueryTransformHandler(workbookRowTileArgsSchema, queries.rowMetadata.tile.fn),
  'sheetCol.tile': createQueryTransformHandler(workbookColumnTileArgsSchema, queries.sheetCol.tile.fn),
  'columnMetadata.tile': createQueryTransformHandler(workbookColumnTileArgsSchema, queries.columnMetadata.tile.fn),
  'cellStyle.byWorkbook': createQueryTransformHandler(workbookQueryArgsSchema, queries.cellStyle.byWorkbook.fn),
  'numberFormat.byWorkbook': createQueryTransformHandler(workbookQueryArgsSchema, queries.numberFormat.byWorkbook.fn),
  'presenceCoarse.byWorkbook': createQueryTransformHandler(workbookQueryArgsSchema, queries.presenceCoarse.byWorkbook.fn),
  'presence.byWorkbook': createQueryTransformHandler(workbookQueryArgsSchema, queries.presence.byWorkbook.fn),
  'workbookChange.byWorkbook': createQueryTransformHandler(workbookQueryArgsSchema, queries.workbookChange.byWorkbook.fn),
  'workbookChanges.byWorkbook': createQueryTransformHandler(workbookQueryArgsSchema, queries.workbookChanges.byWorkbook.fn),
  'workbookChatThread.byWorkbook': createQueryTransformHandler(workbookQueryArgsSchema, queries.workbookChatThread.byWorkbook.fn),
  'workbookChatThread.visibleByWorkbook': createQueryTransformHandler(
    workbookQueryArgsSchema,
    queries.workbookChatThread.visibleByWorkbook.fn,
  ),
  'workbookChatItem.byThread': createQueryTransformHandler(workbookThreadArgsSchema, queries.workbookChatItem.byThread.fn),
  'workbookChatToolCall.byThread': createQueryTransformHandler(workbookThreadArgsSchema, queries.workbookChatToolCall.byThread.fn),
  'workbookReviewQueueItem.byThread': createQueryTransformHandler(workbookThreadArgsSchema, queries.workbookReviewQueueItem.byThread.fn),
  'workbookAgentThread.byWorkbook': createQueryTransformHandler(workbookQueryArgsSchema, queries.workbookAgentThread.byWorkbook.fn),
  'workbookAgentThread.visibleByWorkbook': createQueryTransformHandler(
    workbookQueryArgsSchema,
    queries.workbookAgentThread.visibleByWorkbook.fn,
  ),
  'workbookAgentRun.byWorkbook': createQueryTransformHandler(workbookQueryArgsSchema, queries.workbookAgentRun.byWorkbook.fn),
  'workbookAgentRun.byThread': createQueryTransformHandler(workbookThreadArgsSchema, queries.workbookAgentRun.byThread.fn),
  'workbookAgentRun.visibleByThread': createQueryTransformHandler(workbookThreadArgsSchema, queries.workbookAgentRun.visibleByThread.fn),
  'workbookWorkflowRun.byThread': createQueryTransformHandler(workbookThreadArgsSchema, queries.workbookWorkflowRun.byThread.fn),
  'workbookWorkflowRun.visibleByThread': createQueryTransformHandler(
    workbookThreadArgsSchema,
    queries.workbookWorkflowRun.visibleByThread.fn,
  ),
  'workbookWorkflowStep.byThread': createQueryTransformHandler(workbookThreadArgsSchema, queries.workbookWorkflowStep.byThread.fn),
  'workbookWorkflowArtifact.byThread': createQueryTransformHandler(workbookThreadArgsSchema, queries.workbookWorkflowArtifact.byThread.fn),
  'workbookAgentWorkflowRun.byThread': createQueryTransformHandler(workbookThreadArgsSchema, queries.workbookAgentWorkflowRun.byThread.fn),
  'workbookAgentWorkflowRun.visibleByThread': createQueryTransformHandler(
    workbookThreadArgsSchema,
    queries.workbookAgentWorkflowRun.visibleByThread.fn,
  ),
}

export const zeroQueryTransformNames = Object.freeze(Object.keys(zeroQueryTransforms).toSorted())

export function executeZeroQueryTransform(
  name: string,
  args: Parameters<TransformQueryFunction>[1],
  userID: string,
): ReturnType<TransformQueryFunction> {
  const query = zeroQueryTransforms[name]
  if (!query) {
    throw new Error(`Unknown Zero query: ${name}`)
  }
  return query.execute(args, userID)
}
