import { boolean, createSchema, json, number, relationships, string, table } from '@rocicorp/zero'

const workbooks = table('workbooks')
  .columns({
    id: string(),
    name: string(),
    ownerUserId: string().from('owner_user_id'),
    headRevision: number().from('head_revision'),
    calculatedRevision: number().from('calculated_revision'),
    calcMode: string<'automatic' | 'manual'>().from('calc_mode'),
    compatibilityMode: string<'excel-modern' | 'odf-1.4'>().from('compatibility_mode'),
    recalcEpoch: number().from('recalc_epoch'),
    createdAt: number().from('created_at'),
    updatedAt: number().from('updated_at'),
  })
  .primaryKey('id')

const sheets = table('sheets')
  .columns({
    workbookId: string().from('workbook_id'),
    sheetId: number().from('sheet_id'),
    name: string(),
    sortOrder: number().from('sort_order'),
    freezeRows: number().from('freeze_rows'),
    freezeCols: number().from('freeze_cols'),
    createdAt: number().from('created_at'),
    updatedAt: number().from('updated_at'),
  })
  .primaryKey('workbookId', 'name')

const cellStyles = table('cell_styles')
  .columns({
    workbookId: string().from('workbook_id'),
    styleId: string().from('style_id'),
    styleJson: json().from('record_json'),
    hash: string(),
    createdAt: number().from('created_at'),
  })
  .primaryKey('workbookId', 'styleId')

const numberFormats = table('cell_number_formats')
  .columns({
    workbookId: string().from('workbook_id'),
    formatId: string().from('format_id'),
    kind: string(),
    code: string(),
    createdAt: number().from('created_at'),
  })
  .primaryKey('workbookId', 'formatId')

const cells = table('cells')
  .columns({
    workbookId: string().from('workbook_id'),
    sheetName: string().from('sheet_name'),
    rowNum: number().from('row_num'),
    colNum: number().from('col_num'),
    address: string(),
    inputValue: json().from('input_value').optional(),
    formula: string().optional(),
    format: string().optional(),
    styleId: string().from('style_id').optional(),
    explicitFormatId: string().from('explicit_format_id').optional(),
    sourceRevision: number().from('source_revision'),
    updatedBy: string().from('updated_by'),
    updatedAt: number().from('updated_at'),
  })
  .primaryKey('workbookId', 'sheetName', 'address')

const rowMetadata = table('row_metadata')
  .columns({
    workbookId: string().from('workbook_id'),
    sheetName: string().from('sheet_name'),
    startIndex: number().from('start_index'),
    count: number(),
    size: number().optional(),
    hidden: boolean().optional(),
    sourceRevision: number().from('source_revision'),
    updatedAt: number().from('updated_at'),
  })
  .primaryKey('workbookId', 'sheetName', 'startIndex')

const columnMetadata = table('column_metadata')
  .columns({
    workbookId: string().from('workbook_id'),
    sheetName: string().from('sheet_name'),
    startIndex: number().from('start_index'),
    count: number(),
    size: number().optional(),
    hidden: boolean().optional(),
    sourceRevision: number().from('source_revision'),
    updatedAt: number().from('updated_at'),
  })
  .primaryKey('workbookId', 'sheetName', 'startIndex')

const cellEval = table('cell_eval')
  .columns({
    workbookId: string().from('workbook_id'),
    sheetName: string().from('sheet_name'),
    rowNum: number().from('row_num'),
    colNum: number().from('col_num'),
    address: string(),
    value: json(),
    styleId: string().from('style_id').optional(),
    formatId: string().from('format_id').optional(),
    styleJson: json().from('style_json').optional(),
    formatCode: string().from('format_code').optional(),
    flags: number(),
    version: number(),
    calcRevision: number().from('calc_revision'),
    updatedAt: number().from('updated_at'),
  })
  .primaryKey('workbookId', 'sheetName', 'address')

const definedNames = table('defined_names')
  .columns({
    workbookId: string().from('workbook_id'),
    name: string(),
    value: json(),
  })
  .primaryKey('workbookId', 'name')

const presenceCoarse = table('presence_coarse')
  .columns({
    workbookId: string().from('workbook_id'),
    sessionId: string().from('session_id'),
    userId: string().from('user_id'),
    presenceClientId: string().from('presence_client_id').optional(),
    sheetId: number().from('sheet_id').optional(),
    sheetName: string().from('sheet_name').optional(),
    address: string().optional(),
    selectionJson: json().from('selection_json').optional(),
    updatedAt: number().from('updated_at'),
  })
  .primaryKey('workbookId', 'sessionId')

const workbookChange = table('workbook_change')
  .columns({
    workbookId: string().from('workbook_id'),
    revision: number(),
    actorUserId: string().from('actor_user_id'),
    clientMutationId: string().from('client_mutation_id').optional(),
    eventKind: string().from('event_kind'),
    summary: string(),
    sheetId: number().from('sheet_id').optional(),
    sheetName: string().from('sheet_name').optional(),
    anchorAddress: string().from('anchor_address').optional(),
    rangeJson: json().from('range_json').optional(),
    undoBundleJson: json().from('undo_bundle_json').optional(),
    revertedByRevision: number().from('reverted_by_revision').optional(),
    revertsRevision: number().from('reverts_revision').optional(),
    createdAt: number().from('created_at'),
  })
  .primaryKey('workbookId', 'revision')

const workbookChatThread = table('workbook_chat_thread')
  .columns({
    workbookId: string().from('workbook_id'),
    threadId: string().from('thread_id'),
    ownerUserId: string().from('actor_user_id'),
    scope: string<'private' | 'shared'>(),
    executionPolicy: string<'autoApplySafe' | 'autoApplyAll' | 'ownerReview'>().from('execution_policy'),
    context: json().from('context_json').optional(),
    updatedAtUnixMs: number().from('updated_at_unix_ms'),
    entryCount: number().from('entry_count'),
    reviewQueueItemCount: number().from('review_queue_item_count'),
    latestEntryText: string().from('latest_entry_text').optional(),
  })
  .primaryKey('workbookId', 'threadId', 'ownerUserId')

const workbookChatItem = table('workbook_chat_item')
  .columns({
    workbookId: string().from('workbook_id'),
    threadId: string().from('thread_id'),
    actorUserId: string().from('actor_user_id'),
    entryId: string().from('entry_id'),
    sortOrder: number().from('sort_order'),
    turnId: string().from('turn_id').optional(),
    kind: string<'user' | 'assistant' | 'plan' | 'reasoning' | 'tool' | 'system'>(),
    text: string().optional(),
    phase: string().optional(),
    toolName: string().from('tool_name').optional(),
    toolStatus: string<'inProgress' | 'completed' | 'failed'>().from('tool_status').optional(),
    argumentsText: string().from('arguments_text').optional(),
    outputText: string().from('output_text').optional(),
    success: boolean().optional(),
    citations: json().from('citations_json').optional(),
  })
  .primaryKey('workbookId', 'threadId', 'actorUserId', 'entryId')

const workbookChatToolCall = table('workbook_chat_tool_call')
  .columns({
    workbookId: string().from('workbook_id'),
    threadId: string().from('thread_id'),
    actorUserId: string().from('actor_user_id'),
    entryId: string().from('entry_id'),
    sortOrder: number().from('sort_order'),
    turnId: string().from('turn_id').optional(),
    toolName: string().from('tool_name').optional(),
    toolStatus: string<'inProgress' | 'completed' | 'failed'>().from('tool_status').optional(),
    argumentsText: string().from('arguments_text').optional(),
    outputText: string().from('output_text').optional(),
    success: boolean().optional(),
  })
  .primaryKey('workbookId', 'threadId', 'actorUserId', 'entryId')

const workbookReviewQueueItem = table('workbook_review_queue_item')
  .columns({
    workbookId: string().from('workbook_id'),
    threadId: string().from('thread_id'),
    actorUserId: string().from('actor_user_id'),
    reviewItemId: string().from('review_item_id'),
    turnId: string().from('turn_id'),
    goalText: string().from('goal_text'),
    summary: string(),
    scope: string<'selection' | 'sheet' | 'workbook'>(),
    riskClass: string<'low' | 'medium' | 'high'>().from('risk_class'),
    reviewMode: string<'manual' | 'ownerReview'>().from('review_mode'),
    ownerUserId: string().from('owner_user_id').optional(),
    status: string<'pending' | 'approved' | 'rejected'>(),
    decidedByUserId: string().from('decided_by_user_id').optional(),
    decidedAtUnixMs: number().from('decided_at_unix_ms').optional(),
    baseRevision: number().from('base_revision'),
    createdAtUnixMs: number().from('created_at_unix_ms'),
    context: json().from('context_json').optional(),
    commands: json().from('commands_json'),
    affectedRanges: json().from('affected_ranges_json'),
    estimatedAffectedCells: number().from('estimated_affected_cells').optional(),
    recommendations: json().from('recommendations_json'),
  })
  .primaryKey('workbookId', 'threadId', 'actorUserId', 'reviewItemId')

const workbookAgentRun = table('workbook_agent_run')
  .columns({
    id: string(),
    bundleId: string().from('bundle_id'),
    workbookId: string().from('workbook_id'),
    threadId: string().from('thread_id'),
    turnId: string().from('turn_id'),
    actorUserId: string().from('actor_user_id'),
    goalText: string().from('goal_text'),
    planText: string().from('plan_text').optional(),
    summary: string(),
    scope: string<'selection' | 'sheet' | 'workbook'>(),
    riskClass: string<'low' | 'medium' | 'high'>().from('risk_class'),
    acceptedScope: string<'full' | 'partial'>().from('accepted_scope'),
    appliedBy: string<'user' | 'auto'>().from('applied_by'),
    baseRevision: number().from('base_revision'),
    appliedRevision: number().from('applied_revision'),
    createdAtUnixMs: number().from('created_at_unix_ms'),
    appliedAtUnixMs: number().from('applied_at_unix_ms'),
    context: json().from('context_json').optional(),
    commands: json().from('commands_json'),
    preview: json().from('preview_json').optional(),
  })
  .primaryKey('id')

const workbookWorkflowRun = table('workbook_workflow_run')
  .columns({
    runId: string().from('run_id'),
    workbookId: string().from('workbook_id'),
    threadId: string().from('thread_id'),
    startedByUserId: string().from('actor_user_id'),
    workflowTemplate: string().from('workflow_template'),
    title: string(),
    summary: string(),
    status: string(),
    createdAtUnixMs: number().from('created_at_unix_ms'),
    updatedAtUnixMs: number().from('updated_at_unix_ms'),
    completedAtUnixMs: number().from('completed_at_unix_ms').optional(),
    errorMessage: string().from('error_message').optional(),
    steps: json().from('steps_json'),
    artifact: json().from('artifact_json').optional(),
  })
  .primaryKey('runId')

const workbookWorkflowStep = table('workbook_workflow_step')
  .columns({
    workbookId: string().from('workbook_id'),
    runId: string().from('run_id'),
    stepId: string().from('step_id'),
    stepOrder: number().from('step_order'),
    label: string(),
    status: string<'pending' | 'running' | 'completed' | 'failed' | 'cancelled'>(),
    summary: string(),
    updatedAtUnixMs: number().from('updated_at_unix_ms'),
  })
  .primaryKey('runId', 'stepId')

const workbookWorkflowArtifact = table('workbook_workflow_artifact')
  .columns({
    runId: string().from('run_id'),
    workbookId: string().from('workbook_id'),
    kind: string<'markdown'>(),
    title: string(),
    text: string(),
    updatedAtUnixMs: number().from('updated_at_unix_ms'),
  })
  .primaryKey('runId')

const workbookAgentRunRelationships = relationships(workbookAgentRun, ({ many }) => ({
  ownerChatThreads: many({
    sourceField: ['workbookId', 'threadId', 'actorUserId'],
    destField: ['workbookId', 'threadId', 'ownerUserId'],
    destSchema: workbookChatThread,
  }),
}))

const workbookChatItemRelationships = relationships(workbookChatItem, ({ one }) => ({
  thread: one({
    sourceField: ['workbookId', 'threadId', 'actorUserId'],
    destField: ['workbookId', 'threadId', 'ownerUserId'],
    destSchema: workbookChatThread,
  }),
}))

const workbookChatToolCallRelationships = relationships(workbookChatToolCall, ({ one }) => ({
  thread: one({
    sourceField: ['workbookId', 'threadId', 'actorUserId'],
    destField: ['workbookId', 'threadId', 'ownerUserId'],
    destSchema: workbookChatThread,
  }),
}))

const workbookReviewQueueItemRelationships = relationships(workbookReviewQueueItem, ({ one }) => ({
  thread: one({
    sourceField: ['workbookId', 'threadId', 'actorUserId'],
    destField: ['workbookId', 'threadId', 'ownerUserId'],
    destSchema: workbookChatThread,
  }),
}))

const workbookWorkflowRunRelationships = relationships(workbookWorkflowRun, ({ many }) => ({
  chatThreads: many({
    sourceField: ['workbookId', 'threadId'],
    destField: ['workbookId', 'threadId'],
    destSchema: workbookChatThread,
  }),
  ownerChatThreads: many({
    sourceField: ['workbookId', 'threadId', 'startedByUserId'],
    destField: ['workbookId', 'threadId', 'ownerUserId'],
    destSchema: workbookChatThread,
  }),
}))

const workbookWorkflowStepRelationships = relationships(workbookWorkflowStep, ({ one }) => ({
  workflowRun: one({
    sourceField: ['runId'],
    destField: ['runId'],
    destSchema: workbookWorkflowRun,
  }),
}))

const workbookWorkflowArtifactRelationships = relationships(workbookWorkflowArtifact, ({ one }) => ({
  workflowRun: one({
    sourceField: ['runId'],
    destField: ['runId'],
    destSchema: workbookWorkflowRun,
  }),
}))

const cellRelationships = relationships(cells, ({ one }) => ({
  sheet: one({
    sourceField: ['workbookId', 'sheetName'],
    destField: ['workbookId', 'name'],
    destSchema: sheets,
  }),
}))

const cellEvalRelationships = relationships(cellEval, ({ one }) => ({
  sheet: one({
    sourceField: ['workbookId', 'sheetName'],
    destField: ['workbookId', 'name'],
    destSchema: sheets,
  }),
}))

const rowMetadataRelationships = relationships(rowMetadata, ({ one }) => ({
  sheet: one({
    sourceField: ['workbookId', 'sheetName'],
    destField: ['workbookId', 'name'],
    destSchema: sheets,
  }),
}))

const columnMetadataRelationships = relationships(columnMetadata, ({ one }) => ({
  sheet: one({
    sourceField: ['workbookId', 'sheetName'],
    destField: ['workbookId', 'name'],
    destSchema: sheets,
  }),
}))

export const schema = createSchema({
  tables: [
    workbooks,
    sheets,
    cellStyles,
    numberFormats,
    cells,
    rowMetadata,
    columnMetadata,
    cellEval,
    definedNames,
    presenceCoarse,
    workbookChange,
    workbookChatThread,
    workbookChatItem,
    workbookChatToolCall,
    workbookReviewQueueItem,
    workbookAgentRun,
    workbookWorkflowRun,
    workbookWorkflowStep,
    workbookWorkflowArtifact,
  ],
  relationships: [
    workbookChatItemRelationships,
    workbookChatToolCallRelationships,
    workbookReviewQueueItemRelationships,
    workbookAgentRunRelationships,
    workbookWorkflowRunRelationships,
    workbookWorkflowStepRelationships,
    workbookWorkflowArtifactRelationships,
    cellRelationships,
    cellEvalRelationships,
    rowMetadataRelationships,
    columnMetadataRelationships,
  ],
})

export const zeroSchemaTableNames = Object.keys(schema.tables)

export const zeroSchemaColumnNamesByTable: Readonly<Record<string, readonly string[]>> = Object.fromEntries(
  Object.entries(schema.tables).map(([tableName, tableSchema]) => [tableName, Object.keys(tableSchema.columns)]),
)

export const zeroSchemaServerColumnNamesByTable: Readonly<Record<string, readonly string[]>> = Object.fromEntries(
  Object.entries(schema.tables).map(([tableName, tableSchema]) => [
    tableName,
    Object.entries(tableSchema.columns).map(([columnName, columnSchema]) => columnSchema.serverName ?? columnName),
  ]),
)

export const sheetIdDependentTableNames = ['presence_coarse', 'workbook_change'] as const satisfies readonly (keyof typeof schema.tables)[]

declare module '@rocicorp/zero' {
  interface DefaultTypes {
    schema: typeof schema
  }
}
