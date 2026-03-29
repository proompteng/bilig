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

export function decodeUnknownSync<Decoded, Encoded>(
  schema: Schema.Schema<Decoded, Encoded>,
  input: unknown,
): Decoded {
  return Schema.decodeUnknownSync(schema)(input);
}
