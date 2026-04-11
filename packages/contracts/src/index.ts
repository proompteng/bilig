import { Schema } from "effect";

export const AuthSourceSchema = Schema.Literal("header", "cookie", "guest");
export type AuthSource = Schema.Schema.Type<typeof AuthSourceSchema>;

export const RuntimeSessionSchema = Schema.Struct({
  authToken: Schema.String,
  userId: Schema.String,
  roles: Schema.Array(Schema.String),
  isAuthenticated: Schema.Boolean,
  authSource: AuthSourceSchema,
});
export type RuntimeSession = Schema.Schema.Type<typeof RuntimeSessionSchema>;

export const ErrorEnvelopeSchema = Schema.Struct({
  error: Schema.String,
  message: Schema.String,
  retryable: Schema.Boolean,
});
export type ErrorEnvelope = Schema.Schema.Type<typeof ErrorEnvelopeSchema>;

export const DocumentStateSummarySchema = Schema.Struct({
  documentId: Schema.String,
  cursor: Schema.Number,
  owner: Schema.Union(Schema.String, Schema.Null),
  sessions: Schema.Array(Schema.String),
  latestSnapshotCursor: Schema.Union(Schema.Number, Schema.Null),
});
export type DocumentStateSummary = Schema.Schema.Type<typeof DocumentStateSummarySchema>;

export const SnapshotMetadataSchema = Schema.Struct({
  cursor: Schema.Number,
  contentType: Schema.String,
});
export type SnapshotMetadata = Schema.Schema.Type<typeof SnapshotMetadataSchema>;

export const WorkbookAgentUiSelectionSchema = Schema.Struct({
  sheetName: Schema.String,
  address: Schema.String,
});
export type WorkbookAgentUiSelection = Schema.Schema.Type<typeof WorkbookAgentUiSelectionSchema>;

export const WorkbookViewportSchema = Schema.Struct({
  rowStart: Schema.Number,
  rowEnd: Schema.Number,
  colStart: Schema.Number,
  colEnd: Schema.Number,
});
export type WorkbookViewport = Schema.Schema.Type<typeof WorkbookViewportSchema>;

export const WorkbookAgentUiContextSchema = Schema.Struct({
  selection: WorkbookAgentUiSelectionSchema,
  viewport: WorkbookViewportSchema,
});
export type WorkbookAgentUiContext = Schema.Schema.Type<typeof WorkbookAgentUiContextSchema>;

export const WorkbookAgentEntryKindSchema = Schema.Literal(
  "user",
  "assistant",
  "plan",
  "tool",
  "system",
);
export type WorkbookAgentEntryKind = Schema.Schema.Type<typeof WorkbookAgentEntryKindSchema>;

export const WorkbookAgentToolStatusSchema = Schema.Union(
  Schema.Literal("inProgress", "completed", "failed"),
  Schema.Null,
);
export type WorkbookAgentToolStatus = Schema.Schema.Type<typeof WorkbookAgentToolStatusSchema>;

export const WorkbookAgentRangeCitationRoleSchema = Schema.Literal("target", "source");
export type WorkbookAgentRangeCitationRole = Schema.Schema.Type<
  typeof WorkbookAgentRangeCitationRoleSchema
>;

export const WorkbookAgentRangeCitationSchema = Schema.Struct({
  kind: Schema.Literal("range"),
  sheetName: Schema.String,
  startAddress: Schema.String,
  endAddress: Schema.String,
  role: WorkbookAgentRangeCitationRoleSchema,
});
export type WorkbookAgentRangeCitation = Schema.Schema.Type<
  typeof WorkbookAgentRangeCitationSchema
>;

export const WorkbookAgentRevisionCitationSchema = Schema.Struct({
  kind: Schema.Literal("revision"),
  revision: Schema.Number,
});
export type WorkbookAgentRevisionCitation = Schema.Schema.Type<
  typeof WorkbookAgentRevisionCitationSchema
>;

export const WorkbookAgentTimelineCitationSchema = Schema.Union(
  WorkbookAgentRangeCitationSchema,
  WorkbookAgentRevisionCitationSchema,
);
export type WorkbookAgentTimelineCitation = Schema.Schema.Type<
  typeof WorkbookAgentTimelineCitationSchema
>;

export const WorkbookAgentTimelineEntrySchema = Schema.Struct({
  id: Schema.String,
  kind: WorkbookAgentEntryKindSchema,
  turnId: Schema.Union(Schema.String, Schema.Null),
  text: Schema.Union(Schema.String, Schema.Null),
  phase: Schema.Union(Schema.String, Schema.Null),
  toolName: Schema.Union(Schema.String, Schema.Null),
  toolStatus: WorkbookAgentToolStatusSchema,
  argumentsText: Schema.Union(Schema.String, Schema.Null),
  outputText: Schema.Union(Schema.String, Schema.Null),
  success: Schema.Union(Schema.Boolean, Schema.Null),
  citations: Schema.Array(WorkbookAgentTimelineCitationSchema),
});
export type WorkbookAgentTimelineEntry = Schema.Schema.Type<
  typeof WorkbookAgentTimelineEntrySchema
>;

export const WorkbookAgentSessionStatusSchema = Schema.Literal("idle", "inProgress", "failed");
export type WorkbookAgentSessionStatus = Schema.Schema.Type<
  typeof WorkbookAgentSessionStatusSchema
>;

export const WorkbookAgentThreadScopeSchema = Schema.Literal("private", "shared");
export type WorkbookAgentThreadScope = Schema.Schema.Type<typeof WorkbookAgentThreadScopeSchema>;

export const WorkbookAgentThreadSummarySchema = Schema.Struct({
  threadId: Schema.String,
  scope: WorkbookAgentThreadScopeSchema,
  ownerUserId: Schema.String,
  updatedAtUnixMs: Schema.Number,
  entryCount: Schema.Number,
  hasPendingBundle: Schema.Boolean,
  latestEntryText: Schema.Union(Schema.String, Schema.Null),
});
export type WorkbookAgentThreadSummary = Schema.Schema.Type<
  typeof WorkbookAgentThreadSummarySchema
>;

export const WorkbookAgentWorkflowTemplateSchema = Schema.Literal(
  "summarizeWorkbook",
  "summarizeCurrentSheet",
  "describeRecentChanges",
  "findFormulaIssues",
  "traceSelectionDependencies",
  "explainSelectionCell",
  "searchWorkbookQuery",
  "createSheet",
  "renameCurrentSheet",
  "hideCurrentRow",
  "hideCurrentColumn",
);
export type WorkbookAgentWorkflowTemplate = Schema.Schema.Type<
  typeof WorkbookAgentWorkflowTemplateSchema
>;

export const WorkbookAgentWorkflowStatusSchema = Schema.Literal(
  "running",
  "completed",
  "failed",
  "cancelled",
);
export type WorkbookAgentWorkflowStatus = Schema.Schema.Type<
  typeof WorkbookAgentWorkflowStatusSchema
>;

export const WorkbookAgentWorkflowStepStatusSchema = Schema.Literal(
  "pending",
  "running",
  "completed",
  "failed",
  "cancelled",
);
export type WorkbookAgentWorkflowStepStatus = Schema.Schema.Type<
  typeof WorkbookAgentWorkflowStepStatusSchema
>;

export const WorkbookAgentWorkflowStepSchema = Schema.Struct({
  stepId: Schema.String,
  label: Schema.String,
  status: WorkbookAgentWorkflowStepStatusSchema,
  summary: Schema.String,
  updatedAtUnixMs: Schema.Number,
});
export type WorkbookAgentWorkflowStep = Schema.Schema.Type<
  typeof WorkbookAgentWorkflowStepSchema
>;

export const WorkbookAgentWorkflowArtifactSchema = Schema.Struct({
  kind: Schema.Literal("markdown"),
  title: Schema.String,
  text: Schema.String,
});
export type WorkbookAgentWorkflowArtifact = Schema.Schema.Type<
  typeof WorkbookAgentWorkflowArtifactSchema
>;

export const WorkbookAgentWorkflowRunSchema = Schema.Struct({
  runId: Schema.String,
  threadId: Schema.String,
  startedByUserId: Schema.String,
  workflowTemplate: WorkbookAgentWorkflowTemplateSchema,
  title: Schema.String,
  summary: Schema.String,
  status: WorkbookAgentWorkflowStatusSchema,
  createdAtUnixMs: Schema.Number,
  updatedAtUnixMs: Schema.Number,
  completedAtUnixMs: Schema.Union(Schema.Number, Schema.Null),
  errorMessage: Schema.Union(Schema.String, Schema.Null),
  steps: Schema.Array(WorkbookAgentWorkflowStepSchema),
  artifact: Schema.Union(WorkbookAgentWorkflowArtifactSchema, Schema.Null),
});
export type WorkbookAgentWorkflowRun = Schema.Schema.Type<typeof WorkbookAgentWorkflowRunSchema>;

export const WorkbookAgentSessionSnapshotSchema = Schema.Struct({
  sessionId: Schema.String,
  documentId: Schema.String,
  threadId: Schema.String,
  scope: WorkbookAgentThreadScopeSchema,
  status: WorkbookAgentSessionStatusSchema,
  activeTurnId: Schema.Union(Schema.String, Schema.Null),
  lastError: Schema.Union(Schema.String, Schema.Null),
  context: Schema.Union(WorkbookAgentUiContextSchema, Schema.Null),
  entries: Schema.Array(WorkbookAgentTimelineEntrySchema),
  pendingBundle: Schema.Union(Schema.Unknown, Schema.Null),
  executionRecords: Schema.Array(Schema.Unknown),
  workflowRuns: Schema.Array(WorkbookAgentWorkflowRunSchema),
});
export type WorkbookAgentSessionSnapshot = Schema.Schema.Type<
  typeof WorkbookAgentSessionSnapshotSchema
>;

export const WorkbookAgentStreamEventSchema = Schema.Union(
  Schema.Struct({
    type: Schema.Literal("snapshot"),
    snapshot: WorkbookAgentSessionSnapshotSchema,
  }),
  Schema.Struct({
    type: Schema.Literal("assistantDelta"),
    itemId: Schema.String,
    delta: Schema.String,
  }),
  Schema.Struct({
    type: Schema.Literal("planDelta"),
    itemId: Schema.String,
    delta: Schema.String,
  }),
);
export type WorkbookAgentStreamEvent = Schema.Schema.Type<typeof WorkbookAgentStreamEventSchema>;

export function decodeUnknownSync<Decoded, Encoded>(
  schema: Schema.Schema<Decoded, Encoded>,
  input: unknown,
): Decoded {
  return Schema.decodeUnknownSync(schema)(input);
}
